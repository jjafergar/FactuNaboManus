# -*- coding: utf-8 -*-
"""
Módulo para comprimir logs antiguos automáticamente.
"""
import os
import gzip
import shutil
from datetime import datetime, timedelta
from pathlib import Path

LOG_DIR = "logs"
COMPRESSED_DIR = os.path.join(LOG_DIR, "compressed")
DAYS_TO_COMPRESS = 30  # Comprimir logs mayores a 30 días

def compress_old_logs(days=DAYS_TO_COMPRESS):
    """Comprime logs más antiguos que el número de días especificado."""
    if not os.path.exists(LOG_DIR):
        return 0
    
    os.makedirs(COMPRESSED_DIR, exist_ok=True)
    cutoff_date = datetime.now() - timedelta(days=days)
    compressed_count = 0
    
    for file_path in Path(LOG_DIR).glob("*.log"):
        if file_path.is_file():
            # Obtener fecha de modificación
            file_mtime = datetime.fromtimestamp(file_path.stat().st_mtime)
            
            if file_mtime < cutoff_date:
                # Comprimir archivo
                compressed_path = os.path.join(COMPRESSED_DIR, f"{file_path.name}.gz")
                
                try:
                    with open(file_path, 'rb') as f_in:
                        with gzip.open(compressed_path, 'wb') as f_out:
                            shutil.copyfileobj(f_in, f_out)
                    
                    # Eliminar archivo original después de comprimir
                    file_path.unlink()
                    compressed_count += 1
                except Exception as e:
                    print(f"Error comprimiendo {file_path}: {e}")
    
    return compressed_count

def compress_old_xmls(days=DAYS_TO_COMPRESS):
    """Comprime XMLs antiguos en el directorio de logs."""
    if not os.path.exists(LOG_DIR):
        return 0
    
    os.makedirs(COMPRESSED_DIR, exist_ok=True)
    cutoff_date = datetime.now() - timedelta(days=days)
    compressed_count = 0
    
    for file_path in Path(LOG_DIR).glob("*.xml"):
        if file_path.is_file():
            file_mtime = datetime.fromtimestamp(file_path.stat().st_mtime)
            
            if file_mtime < cutoff_date:
                compressed_path = os.path.join(COMPRESSED_DIR, f"{file_path.name}.gz")
                
                try:
                    with open(file_path, 'rb') as f_in:
                        with gzip.open(compressed_path, 'wb') as f_out:
                            shutil.copyfileobj(f_in, f_out)
                    
                    file_path.unlink()
                    compressed_count += 1
                except Exception as e:
                    print(f"Error comprimiendo XML {file_path}: {e}")
    
    return compressed_count

if __name__ == "__main__":
    # Ejecutar compresión si se llama directamente
    logs = compress_old_logs()
    xmls = compress_old_xmls()
    print(f"Comprimidos {logs} logs y {xmls} XMLs")

