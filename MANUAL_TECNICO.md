# 🔧 Manual Técnico - Archivos del Proyecto FactuNabo

Este documento describe brevemente la función de cada archivo Python del proyecto, para facilitar la localización de código cuando sea necesario realizar modificaciones.

---

## 📁 Archivos Principales

### `main.py`
**Propósito**: Archivo principal de la aplicación. Contiene la ventana principal, la interfaz de usuario completa, y orquesta todas las funcionalidades.

**Contiene**:
- Clase `MainWindow`: Ventana principal con todas las páginas (Dashboard, Cargar Excel, Enviar Facturas, Histórico, Configuración)
- Componentes UI: Botones animados, tablas, diálogos, stepper de progreso
- Gestión de usuarios: Login, CRUD de usuarios
- Gestión de base de datos: Consultas y actualizaciones del historial
- Integración con Worker: Comunicación con el hilo de procesamiento
- Estilos y temas: Carga y aplicación de QSS

**Cuándo modificar**: Para cambios en la interfaz, navegación, gestión de usuarios, o integración de nuevas funcionalidades visuales.

---

### `prueba.py`
**Propósito**: Motor de procesamiento de facturas. Genera XMLs, valida contra XSD, y envía a la API de Facturantia.

**Contiene**:
- Función `main()`: Punto de entrada principal del procesamiento
- Generación de XML: Construcción de documentos XML según esquema XSD
- Validación XSD: Verificación de conformidad antes del envío
- Envío a API: Comunicación HTTP con Facturantia
- Normalización de datos: Limpieza y ajuste de tipos de factura (F1, F2, F3, R1-R5)
- Lógica de rectificativas: Detección automática de R1 vs R4
- Marcado en Excel: Actualización de estado en columna AC

**Cuándo modificar**: Para cambios en la lógica de generación de XML, validaciones, tipos de factura, o comunicación con la API.

---

### `macro_adapter.py`
**Propósito**: Adaptador que lee el Excel y convierte los datos a DataFrames estructurados para el procesamiento.

**Contiene**:
- Función `adapt_from_macro()`: Lee hoja "Macro" y "CLIENTES", produce 6 DataFrames
- Mapeo de columnas: Conversión de letras de columna (A, B, E...) a campos estructurados
- Normalización de datos: Limpieza de NIFs, CIFs, fechas, números
- Detección de tipos: Identificación de facturas normales, intereses, intracomunitarias
- Cálculo de IVA: Determinación automática de porcentajes de IVA
- Gestión de múltiples emisores: Agrupación y procesamiento por CIF emisor
- Lectura de historial: Procesamiento de hojas adicionales para facturas históricas

**Cuándo modificar**: Para cambios en la estructura del Excel, nuevas columnas, o lógica de lectura/transformación de datos.

---

### `worker.py`
**Propósito**: Worker que ejecuta el procesamiento en un hilo separado para no bloquear la interfaz.

**Contiene**:
- Clase `Worker`: Hereda de QObject, ejecuta en hilo de trabajo
- Método `process()`: Llama a `prueba.main()` y opcionalmente descarga PDFs
- Método `download_pdfs()`: Descarga masiva de PDFs desde URLs de la API
- Señales Qt: `log_signal`, `finished`, `downloads_done` para comunicación con UI
- Gestión de dataframes históricos: Pasa datos históricos a `prueba.py`

**Cuándo modificar**: Para cambios en el flujo de procesamiento en segundo plano, descarga de PDFs, o comunicación asíncrona con la UI.

---

## 🎨 Archivos de Interfaz

### `modern_dialogs.py`
**Propósito**: Implementa diálogos modernos frameless (sin bordes) con estilo consistente.

**Contiene**:
- Clase `ModernDialogBase`: Base para todos los diálogos modernos
- Clase `ConfirmDialog`: Diálogo de confirmación (Sí/No)
- Clase `TextInputDialog`: Diálogo para entrada de texto (con soporte para contraseñas)
- Función `show_info()`: Diálogo informativo
- Función `ask_yes_no()`: Diálogo de confirmación
- Función `ask_text()`: Diálogo de entrada de texto
- Efectos visuales: Sombras y estilos modernos

**Cuándo modificar**: Para cambios en el diseño de diálogos, añadir nuevos tipos de diálogos, o modificar estilos visuales.

---

### `dialog_shim.py`
**Propósito**: Intercepta llamadas a `QMessageBox` y `QInputDialog` estándar y las redirige a diálogos modernos.

**Contiene**:
- Funciones wrapper: `_question()`, `_information()`, `_warning()`, `_critical()`, `_getText()`
- Reemplazo de métodos estáticos: Sobrescribe métodos de `QMessageBox` y `QInputDialog`

**Cuándo modificar**: Para cambiar el comportamiento de diálogos del sistema o añadir nuevos tipos de interceptación.

---

### `login_dialog.py`
**Propósito**: Gestiona la autenticación de usuarios con almacenamiento seguro de contraseñas.

**Contiene**:
- Clase `UserStore`: Gestión de usuarios (lectura/escritura de `users.json`)
- Funciones de hash: `pbkdf2_hash()`, `pbkdf2_verify()` para contraseñas seguras
- Clase `LoginDialog`: Diálogo de inicio de sesión (aunque actualmente se usa el de `main.py`)

**Cuándo modificar**: Para cambios en el sistema de autenticación, algoritmo de hash, o formato de almacenamiento de usuarios.

---

## 📥 Archivos de Descarga

### `pdf_downloader.py`
**Propósito**: Descarga PDFs de facturas desde URLs usando Selenium (Chrome/Edge).

**Contiene**:
- Función `download_many()`: Descarga masiva de PDFs con nombres personalizados
- Clase `DownloadResult`: Dataclass para resultados de descarga
- Función `_build_driver()`: Configuración de Selenium WebDriver
- Selectores CSS: Para encontrar botones de descarga en páginas web
- Gestión de descargas: Espera de descargas, renombrado de archivos

**Cuándo modificar**: Para cambios en la lógica de descarga, soporte de nuevos navegadores, o modificación de nombres de archivo.

---

## 🛠️ Archivos de Utilidades

### `manual_save.py`
**Propósito**: Script de utilidad para guardar datos manualmente en la base de datos (usado para pruebas o mantenimiento).

**Contiene**:
- Ejecución manual: Llama a `window.on_finished()` sin interfaz gráfica
- Útil para: Procesar datos pendientes o corregir estados en la BD

**Cuándo modificar**: Para añadir nuevas funciones de mantenimiento manual o scripts de utilidad.

---

### `verify_db.py`
**Propósito**: Script de utilidad para verificar y mostrar el contenido de la base de datos SQLite.

**Contiene**:
- Lectura de BD: Conecta a `factunabo_history.db` y muestra contenido de tabla `envios`
- Útil para: Depuración, verificación de datos, o inspección manual

**Cuándo modificar**: Para añadir nuevas consultas de verificación o scripts de análisis de datos.

---

## 📊 Estructura de Dependencias

```
main.py
├── login_dialog.py (autenticación)
├── modern_dialogs.py (diálogos)
├── dialog_shim.py (interceptación)
├── worker.py (procesamiento en hilo)
│   ├── prueba.py (generación XML y envío)
│   │   └── macro_adapter.py (lectura Excel)
│   └── pdf_downloader.py (descarga PDFs)
└── verify_db.py (utilidad)
```

---

## 🔍 Guía Rápida: ¿Dónde buscar?

### Para modificar la interfaz visual:
→ **`main.py`** (páginas, layouts, widgets)

### Para cambiar cómo se lee el Excel:
→ **`macro_adapter.py`** (estructura de columnas, normalización)

### Para modificar la generación de XML:
→ **`prueba.py`** (construcción XML, validación XSD)

### Para cambiar el envío a la API:
→ **`prueba.py`** (función de envío HTTP)

### Para añadir nuevos tipos de factura:
→ **`prueba.py`** y **`macro_adapter.py`** (detección y procesamiento)

### Para modificar diálogos:
→ **`modern_dialogs.py`** (implementación) o **`dialog_shim.py`** (interceptación)

### Para cambiar el sistema de usuarios:
→ **`login_dialog.py`** (autenticación) o **`main.py`** (gestión CRUD)

### Para modificar la descarga de PDFs:
→ **`pdf_downloader.py`** (lógica Selenium) o **`worker.py`** (orquestación)

### Para cambiar estilos visuales:
→ **`styles.qss`** (no es .py, pero importante para UI)

### Para añadir nuevas páginas/secciones:
→ **`main.py`** (métodos `create_*_page()`)

### Para modificar la base de datos:
→ **`main.py`** (función `init_database()` y consultas SQL)

---

## 📝 Notas Importantes

- **`main.py`** es el archivo más grande y central. Contiene la mayoría de la lógica de UI.
- **`prueba.py`** y **`macro_adapter.py`** son los archivos más críticos para el procesamiento de facturas.
- **`worker.py`** actúa como puente entre la UI (main.py) y el procesamiento (prueba.py).
- Los archivos de diálogos (`modern_dialogs.py`, `dialog_shim.py`) son independientes y pueden modificarse sin afectar la lógica principal.
- **`pdf_downloader.py`** requiere Selenium y un navegador instalado (Chrome o Edge).

---

## 🚨 Archivos Críticos (modificar con precaución)

1. **`prueba.py`**: Cambios aquí afectan directamente el envío de facturas
2. **`macro_adapter.py`**: Cambios aquí pueden romper la lectura del Excel
3. **`main.py`**: Archivo muy grande, cambios pueden afectar múltiples funcionalidades
4. **`worker.py`**: Cambios aquí pueden afectar el procesamiento asíncrono

---

**Versión**: 1.0  
**Última actualización**: 2024

