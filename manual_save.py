# manual_save.py
import sys
from PySide6.QtWidgets import QApplication
from main import MainWindow

print("Iniciando la aplicación para la prueba...")
# Es necesario crear una instancia de QApplication para que los componentes de PySide6 funcionen
app = QApplication(sys.argv)

print("Creando una instancia de MainWindow...")
# Creamos la ventana principal, que contiene la lógica de 'on_finished'
window = MainWindow()

print("Invocando manualmente el método on_finished()...")
# Llamamos al método que contiene la lógica de guardado en la base de datos
window.on_finished()

print("Proceso de guardado manual completado.")
# No es necesario ejecutar app.exec() porque no necesitamos la interfaz interactiva
sys.exit(0)
