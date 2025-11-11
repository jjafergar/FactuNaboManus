# 📘 Manual de Usuario - FactuNabo

## Índice
1. [Introducción](#introducción)
2. [Requisitos del Sistema](#requisitos-del-sistema)
3. [Primer Uso - Configuración Inicial](#primer-uso---configuración-inicial)
4. [Estructura del Excel de Facturas](#estructura-del-excel-de-facturas)
5. [Guía de Uso Paso a Paso](#guía-de-uso-paso-a-paso)
6. [Errores Comunes y Soluciones](#errores-comunes-y-soluciones)
7. [Preguntas Frecuentes](#preguntas-frecuentes)

---

## Introducción

**FactuNabo** es una aplicación de escritorio diseñada para gestionar y enviar facturas electrónicas a través de la API de Facturantia. La aplicación permite:

- ✅ Cargar facturas desde archivos Excel
- ✅ Validar facturas antes del envío
- ✅ Enviar facturas masivamente a la API
- ✅ Gestionar el historial de envíos
- ✅ Descargar PDFs de facturas emitidas
- ✅ Consultar estadísticas de facturación

---

## Requisitos del Sistema

- **Sistema Operativo**: Windows 10 o superior
- **Python**: Versión 3.8 o superior (incluido en la instalación)
- **Memoria RAM**: Mínimo 4 GB recomendado
- **Espacio en disco**: 500 MB libres
- **Conexión a Internet**: Requerida para el envío de facturas

---

## Primer Uso - Configuración Inicial

### 1. Inicio de Sesión

Al abrir la aplicación, verás una pantalla de inicio de sesión:

1. **Usuario**: Introduce tu nombre de usuario
2. **Contraseña**: Introduce tu contraseña
3. **Recordarme**: Marca esta opción si deseas que la aplicación recuerde tus credenciales
4. Haz clic en **"Entrar"**

> ⚠️ **Nota**: Si es tu primer uso, contacta con el administrador para obtener tus credenciales.

### 2. Configuración de la API

Antes de enviar facturas, debes configurar la conexión con Facturantia:

1. Ve a **"⚙️ Configuración"** en el menú lateral
2. En la sección **"Conexión API"**, haz clic en **"Configurar Parámetros API"**
3. Completa los siguientes campos:
   - **URL**: URL del endpoint de la API (normalmente proporcionada por Facturantia)
   - **Token**: Token de autenticación de la API
   - **Usuario**: Usuario de la API
   - **Timeout (seg)**: Tiempo de espera para las peticiones (por defecto: 30 segundos)
4. Haz clic en **"Guardar"**

> 💡 **Consejo**: Guarda estos datos en un lugar seguro. Los necesitarás cada vez que cambies de equipo.

---

## Estructura del Excel de Facturas

### Archivo Excel Requerido

El Excel debe contener **dos hojas obligatorias**:

1. **Hoja "Macro"** (o "MACRO", "Hoja1", "Resumen"): Contiene los datos de las facturas
2. **Hoja "CLIENTES"** (o "Clientes", "EMISORES"): Contiene la configuración de las empresas emisoras

### Hoja "Macro" - Estructura de Columnas

Los datos deben comenzar en la **fila 2** (la fila 1 puede contener encabezados, pero no se usa):

| Columna | Letra | Campo | Descripción | Obligatorio |
|---------|-------|-------|-------------|-------------|
| A | A | Nº Factura | Número de factura (ej: "123", "A25-001", "Int-001") | ✅ Sí |
| B | B | Fecha de emisión | Fecha en formato DD/MM/YYYY o YYYY-MM-DD | ✅ Sí |
| E | E | CIF emisor | CIF/NIF de la empresa que emite la factura | ✅ Sí |
| G | G | Nombre cliente | Nombre completo del cliente | ✅ Sí |
| H | H | NIF cliente | NIF/CIF del cliente (se limpia automáticamente) | ✅ Sí |
| I | I | Dirección cliente | Dirección completa del cliente | ✅ Sí |
| J | J | CP + Provincia | Código postal y provincia (ej: "41004 Sevilla") | ✅ Sí |
| K-L | K-L | Concepto 1 | Descripción (K) e Importe (L) del primer concepto | ⚠️ Al menos uno |
| M-N | M-N | Concepto 2 | Descripción (M) e Importe (N) del segundo concepto | ⚠️ Al menos uno |
| O-P | O-P | Concepto 3 | Descripción (O) e Importe (P) del tercer concepto | ⚠️ Al menos uno |
| Q-R | Q-R | Concepto 4 | Descripción (Q) e Importe (R) del cuarto concepto | ⚠️ Al menos uno |
| S-T | S-T | Concepto 5 | Descripción (S) e Importe (T) del quinto concepto | ⚠️ Al menos uno |
| U-V | U-V | Concepto 6 | Descripción (U) e Importe (V) del sexto concepto | ⚠️ Al menos uno |
| W-X | W-X | Concepto 7 | Descripción (W) e Importe (X) del séptimo concepto | ⚠️ Al menos uno |
| Y-Z | Y-Z | Concepto 8 | Descripción (Y) e Importe (Z) del octavo concepto | ⚠️ Al menos uno |
| AA | AA | Gastos suplidos | Importe de gastos suplidos (solo facturas normales) | ❌ No |
| AB | AB | IBAN emisor | IBAN de la empresa emisora (si está vacío, usa el de CLIENTES) | ⚠️ Si AB vacío, debe estar en CLIENTES |
| AC | AC | Estado | Estado de la factura (se marca automáticamente) | ❌ No |
| AD | AD | Base imponible | Base imponible total de la factura | ✅ Sí |
| AH | AH | Total con IVA | Total de la factura incluyendo IVA | ✅ Sí |
| AI | AI | Factura original | Número de factura original (para rectificativas) | ❌ No |

### Tipos de Factura según el Número

El sistema detecta automáticamente el tipo de factura según el número:

- **Facturas normales**: Cualquier número que no empiece por "Int" o "A" (ej: "123", "2024-001")
- **Facturas de intereses**: Números que empiezan por "Int" (ej: "Int-001", "Int2024-001")
  - Aplican retención IRPF del 19%
  - IVA = 0%
- **Facturas intracomunitarias**: Números que empiezan por "A" (ej: "A25-001", "A2024-001")
  - IVA = 0%
  - Requieren NIF-IVA del cliente (formato: "ES" + NIF)

### Hoja "CLIENTES" - Configuración de Emisores

Esta hoja contiene la información de las empresas que emiten facturas:

| Columna | Descripción | Obligatorio | Ejemplo |
|---------|-------------|-------------|---------|
| `cif` o `cif/nif` | CIF/NIF de la empresa emisora | ✅ Sí | "B12345678" |
| `cif_aliases` | Variantes del CIF (separadas por comas) | ❌ No | "B-12345678, B12345678" |
| `empresa_nombre` | Nombre legal de la empresa | ✅ Sí | "Mi Empresa SL" |
| `iban_defecto` | IBAN por defecto para facturas sin IBAN | ✅ Sí* | "ES1234567890123456789012" |
| `bic` | Código BIC del banco | ❌ No | "CAGLESMMXXX" |
| `es_autonomo` | TRUE si es autónomo (aplica retención) | ❌ No | "TRUE" o "FALSE" |
| `series_retencion` | Series que aplican retención (separadas por comas) | ❌ No | "AUT, AUT2" |
| `api_token` | Token de la API de Facturantia | ⚠️ Recomendado | "abc123..." |
| `api_email` | Email de la API | ⚠️ Recomendado | "usuario@facturantia.com" |
| `api_url` | URL del endpoint de la API | ❌ No | "https://..." |
| `unidad_medida_defecto` | Unidad de medida por defecto | ❌ No | "ud" |
| `moneda` | Moneda (por defecto EUR) | ❌ No | "EUR" |
| `plantilla_facturas_emitidas` | Plantilla para facturas emitidas | ❌ No | "Plantilla1" |
| `plantilla_facturas_proforma` | Plantilla para proformas | ❌ No | "Plantilla1" |

> ⚠️ **Importante**: El IBAN es obligatorio. Debe estar en la columna AB de cada factura O en `iban_defecto` de la hoja CLIENTES.

---

## Guía de Uso Paso a Paso

### Paso 1: Cargar el Excel

1. Abre la aplicación e inicia sesión
2. En el menú lateral, selecciona **"📁 Cargar Excel"**
3. Tienes dos opciones:
   - **Opción A**: Haz clic en **"Seleccionar Excel"** y busca el archivo
   - **Opción B**: Arrastra el archivo Excel directamente a la zona indicada
4. El sistema validará automáticamente el archivo

### Paso 2: Revisar y Validar

1. Una vez cargado, verás una tabla con todas las facturas
2. Revisa que los datos sean correctos:
   - Números de factura
   - Fechas
   - Clientes
   - Importes
3. Si hay errores, se mostrarán en rojo. Corrígelos en el Excel y vuelve a cargar

### Paso 3: Enviar Facturas

1. Ve a **"🚀 Enviar Facturas"** en el menú lateral
2. Verás una tabla con las facturas a enviar
3. Revisa la información mostrada
4. Haz clic en **"🚀 Iniciar Envío"**
5. El sistema mostrará el progreso:
   - **Paso 1**: Cargar Excel
   - **Paso 2**: Validar
   - **Paso 3**: Listo
6. Espera a que finalice el proceso

### Paso 4: Revisar Resultados

1. Una vez finalizado, verás los resultados:
   - ✅ **Éxito**: Facturas enviadas correctamente (verde)
   - ⚠️ **Duplicado**: Facturas ya enviadas anteriormente (naranja)
   - ❌ **Error**: Facturas con errores (rojo)
2. Puedes filtrar por estado usando los botones de filtro
3. Para ver detalles de una factura, haz clic en el botón **"Ver Factura"**

### Paso 5: Descargar PDFs (Opcional)

1. Después del envío, puedes descargar los PDFs de las facturas
2. Haz clic en **"📥 Guardar PDFs"**
3. Los PDFs se guardarán en: `C:\FactuNabo\FacturasPDF\`
4. El nombre del archivo será: `[Número] - [Cliente] - [Importe].pdf`

### Paso 6: Consultar Histórico

1. Ve a **"📜 Histórico"** en el menú lateral
2. Puedes consultar todas las facturas enviadas anteriormente
3. Usa los filtros para buscar por:
   - Empresa emisora
   - Período (trimestre)
4. Haz clic en **"Consultar"** para aplicar los filtros
5. Haz clic en **"🔄 Actualizar"** para refrescar los datos

---

## Errores Comunes y Soluciones

### ❌ Error: "La hoja 'Macro' está vacía o no tiene datos"

**Causa**: El Excel no tiene datos en la hoja "Macro" o la estructura es incorrecta.

**Solución**:
1. Verifica que la hoja se llame "Macro" (o "MACRO", "Hoja1", "Resumen")
2. Asegúrate de que los datos comienzan en la fila 2
3. Verifica que hay al menos una fila con datos

---

### ❌ Error: "No se encontró la hoja 'CLIENTES'"

**Causa**: Falta la hoja de configuración de emisores.

**Solución**:
1. Crea una hoja llamada "CLIENTES" (o "Clientes", "EMISORES")
2. Añade las columnas mínimas: `cif`, `empresa_nombre`, `iban_defecto`
3. Añade al menos una fila con los datos de tu empresa

---

### ❌ Error: "Falta IBAN para CIF [CIF] en filas Excel: [números]"

**Causa**: Las facturas no tienen IBAN y no hay `iban_defecto` en la hoja CLIENTES.

**Solución**:
1. **Opción A**: Añade el IBAN en la columna AB de cada factura
2. **Opción B**: Añade `iban_defecto` en la hoja CLIENTES para ese CIF

---

### ❌ Error: "Empresa no configurada en hoja CLIENTES para CIF: [CIF]"

**Causa**: El CIF del emisor en la columna E no coincide con ningún CIF en la hoja CLIENTES.

**Solución**:
1. Verifica que el CIF en la columna E coincida exactamente con el de CLIENTES
2. O añade el CIF como alias en `cif_aliases` en la hoja CLIENTES
3. Asegúrate de que no hay espacios extra o caracteres especiales

---

### ❌ Error: "Número de factura vacío" / "Empresa emisora vacía" / "Fecha de emisión vacía"

**Causa**: Faltan datos obligatorios en alguna fila.

**Solución**:
1. Revisa la fila indicada en el error
2. Completa los campos obligatorios:
   - Columna A: Número de factura
   - Columna E: CIF emisor
   - Columna B: Fecha de emisión

---

### ❌ Error: "Importe inválido (base_unidad <= 0)"

**Causa**: La factura no tiene conceptos con importe válido.

**Solución**:
1. Verifica que al menos un concepto (columnas L, N, P, R, T, V, X, Z) tenga un importe mayor que 0
2. Asegúrate de que los importes están en formato numérico (no texto)

---

### ❌ Error: "Error leyendo archivo (Macro)"

**Causa**: El archivo Excel está corrupto, abierto en otro programa, o tiene un formato no soportado.

**Solución**:
1. Cierra el Excel si está abierto en otro programa
2. Guarda el archivo como `.xlsx` (no `.xls` antiguo)
3. Verifica que el archivo no esté protegido con contraseña
4. Intenta abrir el archivo en Excel para verificar que no está corrupto

---

### ❌ Error: "Reindexing only valid with uniquely valued Index objects"

**Causa**: El Excel tiene filas duplicadas o índices problemáticos.

**Solución**:
1. Elimina filas duplicadas en el Excel
2. Asegúrate de que no hay filas completamente vacías entre los datos
3. Guarda el archivo y vuelve a cargarlo

---

### ❌ Error: "XSD Validation Error" al enviar

**Causa**: El XML generado no cumple con el esquema XSD requerido por Facturantia.

**Solución**:
1. Revisa los logs en la carpeta `logs/` para ver el error específico
2. Verifica que todos los campos obligatorios están completos
3. Para facturas rectificativas (R1, R4), asegúrate de que:
   - Existe `factura_original` en la columna AI
   - El tipo de factura es correcto (R1 para errores de IVA, R4 para otros)

---

### ❌ Error de conexión con la API

**Causa**: Problemas de conexión o credenciales incorrectas.

**Solución**:
1. Verifica tu conexión a Internet
2. Revisa la configuración de la API en "⚙️ Configuración"
3. Verifica que el Token y Usuario son correctos
4. Aumenta el Timeout si la conexión es lenta

---

### ⚠️ Advertencia: "No se pudo procesar el historial de facturas"

**Causa**: El sistema intenta leer hojas de historial pero hay un problema.

**Solución**:
- Esta advertencia no impide el funcionamiento normal
- Solo afecta a la búsqueda de facturas originales para rectificativas
- Si necesitas rectificativas, asegúrate de que la factura original está en la hoja "Macro"

---

## Preguntas Frecuentes

### ¿Puedo enviar facturas de múltiples empresas en un mismo Excel?

**Sí**. El sistema soporta múltiples emisores en un mismo Excel. Solo asegúrate de que:
- Cada emisor tiene su CIF en la columna E
- Cada CIF está configurado en la hoja CLIENTES
- Cada emisor tiene su IBAN (en columna AB o en `iban_defecto`)

---

### ¿Cómo funcionan las facturas rectificativas?

Las facturas rectificativas se detectan automáticamente cuando:
- El número de factura empieza por "R" (R1, R2, R3, R4, R5)
- O cuando hay una factura original en la columna AI

El sistema determina automáticamente:
- **R1**: Si hay errores de IVA detectados
- **R4**: Para otros tipos de rectificación

Asegúrate de incluir el número de factura original en la columna AI.

---

### ¿Qué pasa si una factura ya fue enviada?

El sistema detecta duplicados automáticamente. Si una factura ya fue enviada:
- Aparecerá con estado **"DUPLICADO"** (naranja)
- No se enviará de nuevo
- Puedes filtrar por "DUPLICADO" para verlas

---

### ¿Cómo cambio el tema (claro/oscuro)?

1. En el menú lateral, al final, encontrarás **"Modo Oscuro"**
2. Activa o desactiva el interruptor para cambiar entre tema claro y oscuro

---

### ¿Dónde se guardan los PDFs descargados?

Por defecto, los PDFs se guardan en:
```
C:\FactuNabo\FacturasPDF\
```

Puedes cambiar esta ruta en la configuración (si está disponible).

---

### ¿Puedo editar facturas después de cargarlas?

**No directamente en la aplicación**. Para editar facturas:
1. Edita el archivo Excel original
2. Vuelve a cargarlo en la aplicación
3. El sistema detectará los cambios

---

### ¿Qué formato de fecha debo usar?

El sistema acepta varios formatos:
- `DD/MM/YYYY` (ej: 15/03/2024)
- `YYYY-MM-DD` (ej: 2024-03-15)
- `DD-MM-YYYY` (ej: 15-03-2024)

---

### ¿Cómo sé si una factura se envió correctamente?

Después del envío:
1. Ve a la página de "🚀 Enviar Facturas"
2. Las facturas con estado **"ÉXITO"** (verde) se enviaron correctamente
3. Puedes verificar en el **"📜 Histórico"** que la factura aparece registrada

---

### ¿Qué hago si olvidé mi contraseña?

Contacta con el administrador del sistema para que te proporcione una nueva contraseña o restablezca la tuya.

---

### ¿Puedo exportar el historial?

Actualmente, el historial se muestra en la aplicación. Para exportar:
1. Usa la función de búsqueda y filtros
2. Toma capturas de pantalla si necesitas documentación
3. O contacta con el administrador para exportaciones masivas

---

## Contacto y Soporte

Si encuentras problemas no cubiertos en este manual:

1. Revisa los **logs** en la carpeta `logs/` de la aplicación
2. Consulta los mensajes de error en la interfaz
3. Contacta con el administrador del sistema

---

## Glosario de Términos

- **API**: Interfaz de programación de aplicaciones. En este caso, el servicio de Facturantia.
- **CIF**: Código de Identificación Fiscal (España).
- **IBAN**: International Bank Account Number (número de cuenta bancaria internacional).
- **BIC**: Bank Identifier Code (código de identificación bancaria).
- **IVA**: Impuesto sobre el Valor Añadido.
- **IRPF**: Impuesto sobre la Renta de las Personas Físicas (retención).
- **XSD**: XML Schema Definition (esquema de validación XML).
- **Rectificativa**: Factura que corrige o anula una factura anterior.

---

**Versión del Manual**: 1.0  
**Última actualización**: 2024  
**Aplicación**: FactuNabo

