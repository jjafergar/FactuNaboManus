# dialog_shim.py — reemplaza QMessageBox/QInputDialog por diálogos modernos
from PySide6.QtWidgets import QMessageBox, QInputDialog
from modern_dialogs import ask_yes_no, ask_text, show_info

def _question(parent, title, text, *args, **kwargs):
    return QMessageBox.Yes if ask_yes_no(parent, title, text) else QMessageBox.No

def _information(parent, title, text, *args, **kwargs):
    show_info(parent, title, text)

def _warning(parent, title, text, *args, **kwargs):
    show_info(parent, title, text)

def _critical(parent, title, text, *args, **kwargs):
    show_info(parent, title, text)

def _getText(parent, title, label, echo=None, text="", *args, **kwargs):
    # echo puede ser QLineEdit.Password, QLineEdit.Normal, etc.
    # Si se pasa como argumento posicional, puede estar en text
    if isinstance(text, int):  # Si text es un EchoMode (int), intercambiar
        echo, text = text, ""
    ok, value = ask_text(parent, title, label, default=text or "", echo_mode=echo)
    return (value, ok)

QMessageBox.question = staticmethod(_question)
QMessageBox.information = staticmethod(_information)
QMessageBox.warning = staticmethod(_warning)
QMessageBox.critical = staticmethod(_critical)
QInputDialog.getText = staticmethod(_getText)