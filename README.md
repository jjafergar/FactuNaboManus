# FactuNabo
Programa para emitir Facturas vía API.

## Construir ejecutable (.exe)

1. **Preparar entorno**
   - Instala Python 3.10+ en Windows.
   - Crea un entorno virtual y actívalo:
     ```powershell
     python -m venv .venv
     .\.venv\Scripts\activate
     ```
   - Instala dependencias:
     ```powershell
     pip install -r requirements.txt
     ```

2. **Actualizar traducciones de Qt (opcional)**  
   PyInstaller incluye automáticamente los ficheros `qtbase_es.qm` y `qt_es.qm` si existen en `Lib\site-packages\PySide6\translations`. Comprueba que están presentes; de lo contrario instala el paquete oficial de PySide6 para la misma versión.

3. **Compilar**
   - Ejecuta el script:
     ```powershell
     .\build.bat
     ```
   - El ejecutable quedará en `dist\FactuNabo\FactuNabo.exe`. Copia **toda** la carpeta cuando distribuyas la aplicación.

4. **Modo fichero único (opcional)**
   - Si necesitas un único `.exe`, usa:
     ```powershell
     pyinstaller FactuNabo.spec --noconfirm --clean --onefile
     ```
   - Ten en cuenta que el primer arranque es más lento porque PyInstaller descomprime recursos en tiempo de ejecución.

## Recomendaciones de rendimiento

- **Limpiar logs antes de compilar:** vacía la carpeta `logs/` para reducir el tamaño del paquete.
- **Mantener requirements mínimos:** elimina dependencias no utilizadas del `requirements.txt` antes de instalar.
- **Usar `--clean`:** el script ya lo hace para quitar artefactos intermedios.
- **Verificar rutas relativas:** el código usa `resource_path()` para localizar recursos, así que no hagas referencias absolutas en nuevos módulos.

## Soporte
Para dudas sobre la API de Facturantia (por ejemplo, caducidad de certificados), contacta con su equipo de soporte y actualiza `FactuNabo.spec` / `main.py` con los endpoints que te indiquen.
