# worker.py — descarga de PDFs + nombre "Num - Cliente - Importe" leyendo por-FACTURA desde XML
from __future__ import annotations

import os
import json
import re
import glob
import shutil
import traceback
import xml.etree.ElementTree as ET
from typing import Optional, Dict, List, Any, Iterable, Tuple

import pandas as pd

from PySide6.QtCore import QObject, Signal

# Debe existir un módulo pdf_downloader con una función:
# download_many(urls, dest_dir, browser, headless, name_func)
# que devuelva una lista de objetos con atributos: url, status ("ok" / "error"), error (str opcional)
from pdf_downloader import download_many


def detect_available_browser() -> Tuple[str, Optional[str]]:
    """Devuelve el navegador soportado detectado y, si se conoce, la ruta al ejecutable."""
    candidates = [
        ("chrome", [
            "chrome",
            "google-chrome",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        ]),
        ("edge", [
            "msedge",
            "microsoft-edge",
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ]),
    ]
    for code, options in candidates:
        for entry in options:
            if os.path.isabs(entry):
                if os.path.exists(entry):
                    return code, entry
            else:
                found = shutil.which(entry)
                if found:
                    return code, found
    # Fallback a Chrome si no se detecta ninguno
    return "chrome", None


class Worker(QObject):
    # Señales hacia la UI
    log_signal = Signal(str)
    finished = Signal()           # fin del proceso principal (envío/macro)
    downloads_done = Signal()     # fin de la descarga de PDFs (manual o auto)

    def __init__(self):
        super().__init__()
        # Entradas
        self._excel_path: Optional[str] = None
        self._post_macro_action: str = "MARK"  # MARK | DELETE_OK | NONE …

        # --- [NUEVO] Atributos para dataframes históricos ---
        self._df_factura_historico: Optional[pd.DataFrame] = None
        self._df_conceptos_historico: Optional[pd.DataFrame] = None
        # --- [FIN NUEVO] ---

        # Descarga de PDFs
        self._auto_download: bool = False
        self._pdf_dest_dir: str = r"C:\\FactuNabo\\FacturasPDF"
        self._pdf_browser, self._pdf_browser_path = detect_available_browser()
        self._pdf_headless: bool = True

    # ----------------- Setters llamados desde la UI -----------------
    def set_excel_path(self, path: str):
        self._excel_path = path

    def set_post_macro_action(self, action: str):
        self._post_macro_action = (action or "MARK").upper()

    def set_historical_data(self, df_factura_hist: pd.DataFrame, df_conceptos_hist: pd.DataFrame):
        """Recibe los dataframes históricos desde la UI."""
        self._df_factura_historico = df_factura_hist
        self._df_conceptos_historico = df_conceptos_hist

    def set_download_options(self, auto: bool, dest: str, browser: Optional[str] = None, headless: bool = True):
        self._auto_download = bool(auto)
        if dest:
            self._pdf_dest_dir = dest
        if browser:
            self._pdf_browser = (browser or "chrome").lower()
        else:
            self._pdf_browser, self._pdf_browser_path = detect_available_browser()
        self._pdf_headless = bool(headless)
        self._emit(f"Navegador seleccionado para descargas: {self._pdf_browser.upper()}")

    # ----------------- Utilidades internas -----------------
    def _emit(self, msg: str):
        """Emite log hacia la UI de forma segura."""
        try:
            self.log_signal.emit(str(msg))
        except Exception:
            pass

    @staticmethod
    def _read_summary(path: str) -> List[dict]:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)

        if isinstance(data, dict):
            if "proformas_procesadas" in data and isinstance(data["proformas_procesadas"], list):
                return data["proformas_procesadas"]
            if "items" in data and isinstance(data["items"], list):
                return data["items"]
            return [data]
        if isinstance(data, list):
            return data
        return []

    # --------- Búsqueda de URLs en items (robusta) ---------
    @staticmethod
    def _iter_scalars(obj: Any) -> Iterable[Any]:
        if isinstance(obj, dict):
            for v in obj.values():
                yield from Worker._iter_scalars(v)
        elif isinstance(obj, (list, tuple)):
            for v in obj:
                yield from Worker._iter_scalars(v)
        else:
            yield obj

    @staticmethod
    def _first_url_like(text: str) -> Optional[str]:
        if not isinstance(text, str):
            return None
        m = re.search(r"https?://[^\s\"'>)]+", text)
        if m:
            return m.group(0)
        return None

    @staticmethod
    def _looks_like_pdf_url(url: str) -> bool:
        if not isinstance(url, str):
            return False
        u = url.lower()
        return (".pdf" in u) or ("pdf" in u) or ("download" in u) or ("descarga" in u) or ("ver_afc_api.php" in u)

    @staticmethod
    def _extract_pdf_url(item: dict) -> Optional[str]:
        candidate_keys = [
            "url_descarga_pdf", "pdf_url", "pdf", "url_pdf", "enlace_pdf", "link_pdf",
            "download_url", "descarga", "href", "link", "pdfLink", "pdfUrl",
            "invoice_pdf", "factura_pdf", "pdf_enlace", "pdf_href"
        ]
        # 1) Claves directas
        for k in candidate_keys:
            if k in item:
                v = item.get(k)
                if isinstance(v, str) and v.strip():
                    if v.startswith("http"):
                        return v
                    maybe = Worker._first_url_like(v)
                    if maybe:
                        return maybe
        # 2) Dicts anidados
        for k in candidate_keys:
            v = item.get(k)
            if isinstance(v, dict):
                for s in Worker._iter_scalars(v):
                    if isinstance(s, str):
                        u = Worker._first_url_like(s)
                        if u and Worker._looks_like_pdf_url(u):
                            return u
        # 3) Búsqueda global
        for s in Worker._iter_scalars(item):
            if isinstance(s, str):
                u = Worker._first_url_like(s)
                if u and Worker._looks_like_pdf_url(u):
                    return u
        return None

    # --------- Normalización y formateo ---------
    @staticmethod
    def _normalize_invoice_id_value(x: Any) -> str:
        """Normaliza nº factura: '25042.0' -> '25042'; deja 'Int_25003' tal cual."""
        s = str(x).strip()
        if re.fullmatch(r"\d+(?:\.0+)?", s):
            try:
                return str(int(float(s)))
            except Exception:
                return s
        return s

    @staticmethod
    def _parse_amount(val: Any) -> Optional[float]:
        """Convierte importes '1.234,56' / '1,234.56' / 1234.56 -> float."""
        if val is None:
            return None
        try:
            if isinstance(val, (int, float)):
                return float(val)
            s = str(val).strip()
            if not s:
                return None
            if s.count(",") == 1 and s.count(".") >= 1:
                s = s.replace(".", "")
            s = s.replace(",", ".")
            return float(s)
        except Exception:
            return None

    @staticmethod
    def _format_eur(v: Optional[float]) -> str:
        """Formatea a es-ES: 1234.56 -> '1.234,56 €'"""
        if v is None:
            return ""
        s = f"{v:,.2f}"
        s = s.replace(",", "X").replace(".", ",").replace("X", ".")
        return s + " €"

    # --------- Lectura de XML por cada factura ---------
    @staticmethod
    def _text_of(root: ET.Element, *xpaths: str) -> Optional[str]:
        for xp in xpaths:
            el = root.find(xp)
            if el is not None and el.text:
                t = el.text.strip()
                if t:
                    return t
        return None

    def _xmls_sorted(self) -> List[str]:
        try:
            files = glob.glob(os.path.join("responses", "*.xml"))
            files.sort(key=os.path.getmtime, reverse=True)
            # Limitar a 500 para no eternizar (ajustable)
            return files[:500]
        except Exception:
            return []

    def _xml_context_for_item(self, item: dict) -> dict:
        """
        Localiza el XML correspondiente a la factura y devuelve {'cliente': ..., 'importe_total': ...}
        Criterios de matching:
          - Coincidencia por nº de factura (variantes): external_id, referencia, NumFactura, numero, etc.
          - Si no, coincidencia por nombre de cliente (normalizado)
        """
        # Nº factura desde summary
        num = self._normalize_invoice_id_value(
            item.get("id") or item.get("NumFactura") or item.get("num_factura") or item.get("numero") or ""
        )
        # Emisor desde summary
        emisor_item = (item.get("empresa") or item.get("empresa_emisora") or "").strip()
        emisor_norm = re.sub(r"\\s+", " ", emisor_item).lower()

        # Cliente desde summary
        cliente_item = (item.get("cliente") or item.get("nombre_cliente") or "").strip()
        cliente_norm = re.sub(r"\\s+", " ", cliente_item).lower()

        for fx in self._xmls_sorted():
            try:
                tree = ET.parse(fx)
                root = tree.getroot()

                # Nº factura en XML (añadimos external_id y referencia en minúscula)
                xml_num = self._text_of(
                    root,
                    ".//NumFactura", ".//numero", ".//Numero",
                    ".//IdFactura", ".//ExternalId", ".//FacturaNumero",
                    ".//external_id", ".//referencia"  # <-- NUEVO
                )
                xml_num_norm = self._normalize_invoice_id_value(xml_num) if xml_num else ""

                # Matching por número
                match_by_num = bool(num) and xml_num_norm and (xml_num_norm == num)

                # Emisor en XML
                xml_emisor = self._text_of(
                    root,
                    ".//empresa_emisora", ".//emisor", ".//EmisorNombre"
                )
                xml_emisor_norm = re.sub(r"\\s+", " ", xml_emisor).lower() if xml_emisor else ""
                match_by_emisor = bool(emisor_norm) and xml_emisor_norm and (xml_emisor_norm == emisor_norm)

                # Cliente en XML (añadimos cliente/nombre y variantes)
                xml_cliente = self._text_of(
                    root,
                    ".//Cliente", ".//customer", ".//ClienteNombre", ".//RazonSocial",
                    ".//cliente/nombre", ".//cliente/razon_social"  # <-- NUEVO
                )
                xml_cliente_norm = re.sub(r"\\s+", " ", xml_cliente).lower() if xml_cliente else ""
                match_by_cliente = bool(cliente_norm) and xml_cliente_norm and (xml_cliente_norm == cliente_norm)

                # --- LÓGICA DE MATCHING MEJORADA ---
                # Debe coincidir el número Y (el emisor O el cliente)
                if match_by_num and (match_by_emisor or match_by_cliente):
                    # Importe total en XML (añadimos importe_total y total_a_pagar)
                    imp_txt = self._text_of(
                        root,
                        ".//proforma/total_a_pagar", ".//proforma/importe_total",  # <-- priorizar totales a nivel proforma
                        ".//total_a_pagar", ".//ImporteTotal", ".//Total", ".//total", ".//TotalFactura",
                        ".//total_factura", ".//TotalConIVA", ".//ImporteConIVA",
                        ".//importe_total"  # (puede aparecer en conceptos e impuestos; por eso va al final)
                    )
                    imp_val = self._parse_amount(imp_txt) if imp_txt else None
                    return {
                        "cliente": xml_cliente or cliente_item,
                        "importe_total": imp_val
                    }
            except Exception:
                continue

        # Fallback si no encontramos XML exacto
        return {"cliente": cliente_item, "importe_total": None}

    # ----------------- Descarga de PDFs (reutilizable) -----------------
    def download_pdfs(self):
        """
        1) Lee responses/summary.json
        2) Detecta URLs PDF robustamente
        3) Para cada factura, busca su XML y extrae el importe (importe_total/total_a_pagar, etc.)
        4) Descarga y nombra: "Nº Factura - Nombre del cliente - Importe factura"
        """
        try:
            summary_path = os.path.join("responses", "summary.json")
            if not os.path.exists(summary_path):
                self._emit("⚠️ No se encontró responses/summary.json para descargar PDFs.")
                self.downloads_done.emit()
                return

            data = self._read_summary(summary_path)

            ok_statuses = {
                "ok", "success", "duplicate", "duplicado", "atencion", "atención",
                "exito", "éxito", "enviado"
            }

            items_with_url: List[dict] = []
            for x in data:
                url = self._extract_pdf_url(x)
                status_txt = str(x.get("status", "") or x.get("estado", "")).strip().lower()
                if url:
                    items_with_url.append({**x, "__pdf_url__": url})
                elif status_txt in ok_statuses:
                    pass  # informativo

            urls: List[str] = [it["__pdf_url__"] for it in items_with_url]

            if not urls:
                sample_keys = set()
                for x in data[:3]:
                    sample_keys.update(list(x.keys()))
                self._emit(
                    "ℹ️ No se detectaron URLs de PDF en el resumen. "
                    f"Claves ejemplo presentes: {sorted(sample_keys)}"
                )
                self._emit("Sugerencia: revisa las claves del resumen (p. ej. 'pdf_url', 'url_pdf', 'download_url').")
                self.downloads_done.emit()
                return

            # Enriquecer cada item con info del XML (cliente + importe)
            url_to_item: Dict[str, dict] = {}
            for it in items_with_url:
                ctx = self._xml_context_for_item(it)
                it_enriched = dict(it)
                it_enriched.setdefault("cliente", ctx.get("cliente"))
                it_enriched["__importe_total__"] = ctx.get("importe_total")
                url_to_item[it["__pdf_url__"]] = it_enriched

            def build_name_from_item(item: dict) -> str:
                # Nº Factura
                num = (
                    item.get("id")
                    or item.get("NumFactura")
                    or item.get("num_factura")
                    or item.get("numero")
                    or item.get("referencia")      # por si acaso viniera del summary
                    or item.get("external_id")     # por si acaso viniera del summary
                    or ""
                )
                num = self._normalize_invoice_id_value(num)

                # Cliente
                cliente = (
                    item.get("cliente")
                    or item.get("empresa")
                    or item.get("nombre_cliente")
                    or ""
                )
                cliente = str(cliente).strip()

                # Importe (preferir el que sacamos del XML)
                imp_val = item.get("__importe_total__")
                if imp_val is None:
                    imp_raw = (
                        item.get("importe_total")
                        or item.get("total_a_pagar")
                        or item.get("total_factura")
                        or item.get("total")
                        or item.get("importe")
                        or None
                    )
                    imp_val = self._parse_amount(imp_raw)
                importe_str = self._format_eur(imp_val) if imp_val is not None else ""

                # Ensamblado "Nº - Cliente - Importe"
                parts = [p for p in [num, cliente, importe_str] if str(p).strip() != ""]
                base = " - ".join(parts).strip()

                # Sanitizar para nombre de archivo
                base = re.sub(r"[\\/:*?\"<>|]+", "_", base).strip(" -_")
                if len(base) > 140:
                    base = base[:140].rstrip(" .-_")
                return base or "Factura"

            def name_func(url: str, idx: int) -> str:
                item = url_to_item.get(url, {})
                return build_name_from_item(item) if item else f"factura_{idx}"

            dest = self._pdf_dest_dir or r"C:\\FactuNabo\\FacturasPDF"
            os.makedirs(dest, exist_ok=True)
            self._emit(f"📥 Descargando {len(urls)} PDFs → {dest} ({self._pdf_browser}, headless={self._pdf_headless})")

            results = download_many(
                urls,
                dest_dir=dest,
                browser=self._pdf_browser,
                headless=self._pdf_headless,
                name_func=name_func,
            )

            ok = sum(1 for r in results if getattr(r, "status", "") == "ok")
            self._emit(f"✅ Descarga completada: {ok}/{len(results)} correctas.")
            if ok != len(results):
                errores = [
                    f"- {getattr(r, 'url', '')}: {getattr(r, 'error', 'error')}"
                    for r in results
                    if getattr(r, "status", "") != "ok"
                ]
                if errores:
                    self._emit("Algunas descargas fallaron:\\n" + "\\n".join(errores))

            download_map: Dict[str, str] = {
                getattr(r, "url", ""): getattr(r, "path", "")
                for r in results
                if getattr(r, "status", "") == "ok"
                and getattr(r, "url", "")
                and getattr(r, "path", "")
            }
            updated = False
            if download_map:
                for entry in data:
                    url = entry.get("pdf_url")
                    local_path = download_map.get(url)
                    if local_path:
                        entry["pdf_local_path"] = local_path
                        updated = True
                if updated:
                    try:
                        with open(summary_path, "w", encoding="utf-8") as f:
                            json.dump(data, f, indent=4, ensure_ascii=False)
                        self._emit("💾 Resumen actualizado con rutas locales de PDFs.")
                    except Exception as write_err:
                        self._emit(f"⚠️ No se pudo guardar la ruta local en summary.json: {write_err}")

        except Exception as e:
            self._emit(f"❌ Error en descarga de PDFs: {e}")
            traceback.print_exc()
        finally:
            self.downloads_done.emit()

    # ----------------- Flujo principal (hilo de trabajo) -----------------
    def process(self):
        """
        1) Llama a prueba.main() (pipeline macro/adaptación/envío)
        2) Si _auto_download = True, llama a download_pdfs()
        """
        try:
            import prueba as pro
            try:
                pro.set_gui_logger(self._emit)  # opcional
            except Exception:
                pass

            if not self._excel_path or not os.path.exists(self._excel_path):
                self._emit("❌ No hay Excel seleccionado.")
                self.finished.emit()
                return

            os.environ["EXCEL_PATH"] = self._excel_path
            os.environ["POST_MACRO_ACTION"] = self._post_macro_action

            self._emit(f"▶️ Iniciando envío con macro… (acción post-macro: {self._post_macro_action})")

            # --- [MODIFICADO] Pasar los dataframes históricos a pro.main() ---
            pro.main(
                df_factura_historico=self._df_factura_historico,
                df_conceptos_historico=self._df_conceptos_historico
            )
            # --- [FIN MODIFICADO] ---

            self._emit("✔️ Macro finalizada.")

            if self._auto_download:
                self.download_pdfs()

            # Señal de finalización del proceso principal
            self.finished.emit()

        except Exception as e:
            self._emit(f"💥 Error en proceso principal: {e}")
            traceback.print_exc()
            self.finished.emit()
