# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import PySide6

project_root = Path.cwd()

data_files = [
    ('styles.qss', '.'),
    ('EsquemaProformas.xsd', '.'),
    ('users.json', '.'),
    ('factunabo_history.db', '.'),
    ('Plantilla_Facturas.xlsx', '.'),
    ('ESTADO_APLICACION.md', '.'),
    ('MANUAL_USUARIO.md', '.'),
    ('MANUAL_TECNICO.md', '.'),
    ('README.md', '.'),
    ('README_Integracion_Macro.md', '.'),
    ('Plantillas Facturas', 'Plantillas Facturas'),
    ('resources', 'resources'),
    ('logs', 'logs'),
    ('responses', 'responses'),
]

# Añadir traducciones de Qt necesarias para los diálogos en español
translations_dir = Path(PySide6.__file__).resolve().parent / "translations"
translations = []
for qm_name in ("qtbase_es.qm", "qt_es.qm"):
    qm_path = translations_dir / qm_name
    if qm_path.exists():
        translations.append((str(qm_path), f"translations/{qm_name}"))

datas = data_files + translations

a = Analysis(
    ['main.py'],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='FactuNabo',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[str(project_root / 'resources' / 'logo.ico')],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='FactuNabo',
)
