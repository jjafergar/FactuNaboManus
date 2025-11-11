# modern_dialogs.py — diálogos frameless estilo “login”
from __future__ import annotations
from PySide6.QtCore import Qt, QPoint
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame,
    QLineEdit, QWidget, QGraphicsDropShadowEffect
)

def _is_dark(widget: QWidget) -> bool:
    w = widget.window()
    if w and w.property("theme") == "dark":
        return True
    return widget.palette().window().color().lightness() < 128

def _shadow_effect(radius=30):
    eff = QGraphicsDropShadowEffect()
    eff.setBlurRadius(radius)
    eff.setOffset(0, 10)
    return eff

class ModernDialogBase(QDialog):
    def __init__(self, parent=None, title: str = ""):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self._drag_pos = None
        self.setModal(True)

        self.card = QFrame(self)
        self.card.setObjectName("DialogCard")
        self.card.setGraphicsEffect(_shadow_effect())

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.addWidget(self.card)

        self.vbox = QVBoxLayout(self.card)
        self.vbox.setContentsMargins(20, 20, 20, 20)
        self.vbox.setSpacing(16)

        title_bar = QHBoxLayout()
        self.title_lbl = QLabel(title or "")
        self.title_lbl.setObjectName("DialogTitle")
        title_bar.addWidget(self.title_lbl, 1)

        self.btn_close = QPushButton("✕")
        self.btn_close.setFixedSize(28, 28)
        self.btn_close.setObjectName("DialogCloseButton")
        self.btn_close.clicked.connect(self.reject)
        title_bar.addWidget(self.btn_close, 0)
        self.vbox.addLayout(title_bar)

        self.content_box = QVBoxLayout()
        self.vbox.addLayout(self.content_box)

        self.btn_row = QHBoxLayout()
        self.btn_row.addStretch(1)
        self.vbox.addLayout(self.btn_row)

        self._apply_inline_style()

    def _apply_inline_style(self):
        dark = _is_dark(self)
        bg = "#1C1C1E" if dark else "#FFFFFF"
        fg = "#FFFFFF" if dark else "#000000"
        subtle = "#38383A" if dark else "#E5E5EA"

        css = []
        css.append("QFrame#DialogCard {")
        css.append(f"background: {bg}; color: {fg}; border-radius: 16px; border: 1px solid {subtle};")
        css.append("}")
        css.append("QLabel#DialogTitle { font-weight: 700; font-size: 16px; }")
        css.append("QPushButton#DialogCloseButton { border: none; background: transparent; font-size: 14px; border-radius: 6px; }")
        self.card.setStyleSheet("\n".join(css))

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag_pos = e.globalPosition().toPoint() - self.frameGeometry().topLeft()
            e.accept()

    def mouseMoveEvent(self, e):
        if self._drag_pos and (e.buttons() & Qt.LeftButton):
            self.move(e.globalPosition().toPoint() - self._drag_pos)
            e.accept()

    def mouseReleaseEvent(self, e):
        self._drag_pos = None
        super().mouseReleaseEvent(e)

class ConfirmDialog(ModernDialogBase):
    def __init__(self, parent=None, title="Confirmar", message=""):
        super().__init__(parent, title)
        self.msg = QLabel(message)
        self.msg.setWordWrap(True)
        self.content_box.addWidget(self.msg)

        self.btn_yes = QPushButton("Sí")
        self.btn_yes.setObjectName("AnimatedButton")
        self.btn_yes.clicked.connect(self.accept)

        self.btn_no = QPushButton("No")
        self.btn_no.clicked.connect(self.reject)

        self.btn_row.addWidget(self.btn_no)
        self.btn_row.addWidget(self.btn_yes)
        self.btn_yes.setDefault(True)
        self.btn_yes.setAutoDefault(True)

class InfoDialog(ModernDialogBase):
    def __init__(self, parent=None, title="Información", message=""):
        super().__init__(parent, title)
        self.msg = QLabel(message)
        self.msg.setWordWrap(True)
        self.content_box.addWidget(self.msg)

        self.btn_ok = QPushButton("OK")
        self.btn_ok.setObjectName("AnimatedButton")
        self.btn_ok.clicked.connect(self.accept)
        self.btn_row.addWidget(self.btn_ok)

class TextInputDialog(ModernDialogBase):
    def __init__(self, parent=None, title="Nuevo usuario", label="", echo_mode=None):
        super().__init__(parent, title)
        self.label = QLabel(label)
        self.edit = QLineEdit()
        if echo_mode is not None:
            self.edit.setEchoMode(echo_mode)
        self.edit.setPlaceholderText("Escribe aquí…")
        self.content_box.addWidget(self.label)
        self.content_box.addWidget(self.edit)

        self.btn_ok = QPushButton("OK")
        self.btn_ok.setObjectName("AnimatedButton")
        self.btn_ok.clicked.connect(self._ok)

        self.btn_cancel = QPushButton("Cancelar")
        self.btn_cancel.clicked.connect(self.reject)

        self.btn_row.addWidget(self.btn_cancel)
        self.btn_row.addWidget(self.btn_ok)

        self.edit.returnPressed.connect(self._ok)
        self.edit.setFocus()

    def _ok(self):
        self.accept()

    def textValue(self) -> str:
        return self.edit.text()

def ask_yes_no(parent, title: str, text: str) -> bool:
    dlg = ConfirmDialog(parent, title=title, message=text)
    return dlg.exec() == QDialog.Accepted

def show_info(parent, title: str, text: str) -> None:
    dlg = InfoDialog(parent, title=title, message=text)
    dlg.exec()

def ask_text(parent, title: str, label: str, default: str = "", echo_mode=None) -> tuple[bool, str]:
    dlg = TextInputDialog(parent, title=title, label=label, echo_mode=echo_mode)
    if default:
        dlg.edit.setText(str(default))  # Asegurar que default sea string
    ok = dlg.exec() == QDialog.Accepted
    return ok, dlg.textValue()