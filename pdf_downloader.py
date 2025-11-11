# pdf_downloader.py (refactor enfocado)
from __future__ import annotations

import os
import time
import pathlib
from contextlib import contextmanager
from dataclasses import dataclass
from typing import Iterable, List, Optional, Callable, Sequence, Tuple

from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

DEFAULT_SELECTORS: Sequence[Tuple[str, str]] = (
    (By.ID, "div_factura_cliente_descargar_pdf"),
    (By.CSS_SELECTOR, "#div_factura_cliente_descargar_pdf, a[href*='descargar_pdf'], button[id*='descargar'][id*='pdf']"),
)

@dataclass
class DownloadResult:
    url: str
    status: str  # "ok" | "error"
    path: Optional[str] = None
    error: Optional[str] = None

def _build_driver(browser: str, download_dir: str, headless: bool = True):
    os.makedirs(download_dir, exist_ok=True)
    prefs = {
        "download.default_directory": str(pathlib.Path(download_dir).resolve()),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "plugins.always_open_pdf_externally": True,
    }
    if browser.lower() == "edge":
        from selenium.webdriver.edge.options import Options as EdgeOptions
        opts = EdgeOptions()
        opts.add_experimental_option("prefs", prefs)
        if headless:
            opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")
        driver = webdriver.Edge(options=opts)
    else:
        from selenium.webdriver.chrome.options import Options as ChromeOptions
        opts = ChromeOptions()
        opts.add_experimental_option("prefs", prefs)
        if headless:
            opts.add_argument("--headless=new")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")
        driver = webdriver.Chrome(options=opts)
    driver.set_page_load_timeout(120)
    return driver

def _wait_new_pdf(download_dir: str, initial_pdfs: set, start_ts: float, timeout_s: int = 180) -> Optional[str]:
    deadline = time.time() + timeout_s
    seen = set(initial_pdfs)
    while time.time() < deadline:
        current = {f for f in os.listdir(download_dir) if f.lower().endswith(".pdf")}
        new_files = current - seen
        if new_files:
            candidates = [os.path.join(download_dir, f) for f in new_files]
            candidates = [c for c in candidates if os.path.getmtime(c) >= start_ts - 0.1 and os.path.getsize(c) > 0]
            if candidates:
                if not any(f.endswith(".crdownload") for f in os.listdir(download_dir)):
                    chosen = max(candidates, key=os.path.getmtime)
                    last = -1
                    for _ in range(8):
                        size = os.path.getsize(chosen)
                        if size == last and size > 0:
                            break
                        last = size
                        time.sleep(0.3)
                    if os.path.getsize(chosen) > 0:
                        return chosen
        time.sleep(0.3)
    return None

def _click_download(driver, selectors: Sequence[Tuple[str, str]], timeout_click: int = 30):
    last_exc = None
    for by, sel in selectors:
        try:
            btn = WebDriverWait(driver, timeout_click).until(EC.element_to_be_clickable((by, sel)))
            btn.click()
            return
        except Exception as e:
            last_exc = e
    raise RuntimeError(f"No se encontró botón/enlace de descarga. Último error: {last_exc}")

@contextmanager
def _safe_driver(browser: str, download_dir: str, headless: bool = True):
    drv = None
    try:
        drv = _build_driver(browser, download_dir, headless=headless)
        yield drv
    finally:
        if drv:
            try:
                drv.quit()
            except Exception:
                pass

def download_one(
    driver: webdriver.Remote,
    url: str,
    download_dir: str,
    name_base: str,
    selectors: Sequence[Tuple[str, str]] = DEFAULT_SELECTORS,
    timeout_click: int = 30,
    wait_download_s: int = 180,
) -> str:
    before = {f for f in os.listdir(download_dir) if f.lower().endswith(".pdf") and os.path.getsize(os.path.join(download_dir, f)) > 0}
    start_ts = time.time()
    driver.get(url)
    _click_download(driver, selectors, timeout_click=timeout_click)
    tmp_pdf = _wait_new_pdf(download_dir, before, start_ts, timeout_s=wait_download_s)
    if not tmp_pdf:
        raise RuntimeError("No se detectó la descarga del PDF en el tiempo esperado.")
    # renombrado seguro
    base = "".join(("_" if c in '\\/:*?"<>|' else c) for c in name_base).strip("_ .")
    destino = os.path.join(download_dir, f"{base}.pdf")
    i = 1
    while os.path.exists(destino):
        try:
            if os.path.getsize(destino) == 0:
                os.remove(destino)
                break
        except Exception:
            pass
        destino = os.path.join(download_dir, f"{base}_{i}.pdf")
        i += 1
    os.replace(tmp_pdf, destino)
    return destino

def download_many(
    urls: Iterable[str],
    dest_dir: str,
    *,
    prefix: str = "factura_",
    start_index: int = 1,
    browser: str = "chrome",
    headless: bool = True,
    selectors: Sequence[Tuple[str, str]] = DEFAULT_SELECTORS,
    timeout_click: int = 30,
    wait_download_s: int = 180,
    retry: int = 1,
    name_func: Optional[Callable[[str, int], str]] = None,
) -> List[DownloadResult]:
    urls = [u for u in urls if isinstance(u, str) and u.strip().lower().startswith("http")]
    if not urls:
        return []

    pathlib.Path(dest_dir).mkdir(parents=True, exist_ok=True)
    results: List[DownloadResult] = []

    with _safe_driver(browser, dest_dir, headless=headless) as driver:
        for idx, url in enumerate(urls, start=start_index):
            base = name_func(url, idx) if name_func else f"{prefix}{idx}"
            attempts = retry + 1
            last_err = None
            for _ in range(attempts):
                try:
                    path = download_one(
                        driver=driver,
                        url=url,
                        download_dir=dest_dir,
                        name_base=base,
                        selectors=selectors,
                        timeout_click=timeout_click,
                        wait_download_s=wait_download_s,
                    )
                    results.append(DownloadResult(url=url, status="ok", path=path))
                    break
                except (WebDriverException, RuntimeError) as e:
                    last_err = str(e)
                    time.sleep(1.0)
            else:
                results.append(DownloadResult(url=url, status="error", error=last_err or "error desconocido"))

    return results
