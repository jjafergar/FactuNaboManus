# FactuNabo – Integración MACRO ÚNICO + Emisores.xlsx

**Mapa de columnas (hoja Macro, datos desde fila 2):**
- A: Nº Factura
- B: Fecha de emisión
- E: CIF emisor (acepto variantes definidas en `cif_aliases` en Emisores.xlsx)
- G: Nombre cliente
- H: NIF cliente (limpio “CIF/NIF …”)
- I: Dirección cliente
- J: CP + Provincia cliente
- K/L .. Y/Z: 8 líneas de conceptos (texto/importe)
- AA: Gastos suplidos (solo facturas normales; sin IVA; suma a importe_total y total_a_pagar)
- AB: IBAN emisor por factura (si vacío → uso `iban_defecto` de Emisores.xlsx; si ambos vacíos → error)
- AD: Base imponible
- AH: Total con IVA

**Reglas clave:**
- Un solo emisor por Excel (E igual en todas las filas).
- IBAN obligatorio (Macro!AB o `iban_defecto` en Emisores.xlsx).
- BIC opcional (si falta, default).
- Tipos: `Int…` (retención 19%, IVA 0), `A…` (IVA 0), resto normal (IVA = (AH − AD − AA)/AD).
- Autónomos: `es_autonomo=TRUE` + `series_retencion` → retención 19% si el nº empieza por alguno.
- Método de pago: transferencia (beneficiario = `empresa_nombre`).

**Emisores.xlsx (misma carpeta):**
- `cif`, `cif_aliases`, `empresa_nombre`, `iban_defecto` (obligatorio), `bic` (opcional),
  `es_autonomo`, `series_retencion`, opcionales `unidad_medida_defecto`, `moneda`.

**Variables de entorno (opcionales):**
- `API_URL`, `API_TOKEN`.

Archivos: `macro_adapter.py`, `prueba.py`, `main.py`.
