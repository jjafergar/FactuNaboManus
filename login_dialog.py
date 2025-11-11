# login_dialog.py
import os
import json
import hashlib
import hmac
import base64
import secrets
from dataclasses import dataclass
from typing import Dict, Optional

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QFormLayout, QLineEdit,
    QCheckBox, QHBoxLayout, QPushButton, QMessageBox
)
from PySide6.QtCore import Qt, QSettings

USERS_PATH = os.path.join(os.path.dirname(__file__), "users.json")

# ========= Password utils (PBKDF2) =========
# Formato almacenado: "pbkdf2_sha256$<iter>$<salt_b64>$<hash_b64>"

PBK_ALGO = "pbkdf2_sha256"
PBK_ITER = 200_000
PBK_SALT_BYTES = 16
PBK_DKLEN = 32

def pbkdf2_hash(password: str, *, iterations: int = PBK_ITER) -> str:
    salt = secrets.token_bytes(PBK_SALT_BYTES)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations, dklen=PBK_DKLEN)
    return f"{PBK_ALGO}${iterations}${base64.b64encode(salt).decode()}${base64.b64encode(dk).decode()}"

def pbkdf2_verify(password: str, stored: str) -> bool:
    try:
        algo, iters, salt_b64, hash_b64 = stored.split("$", 3)
        if algo != PBK_ALGO:
            return False
        iterations = int(iters)
        salt = base64.b64decode(salt_b64)
        ref = base64.b64decode(hash_b64)
        dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations, dklen=len(ref))
        return hmac.compare_digest(dk, ref)
    except Exception:
        return False

def is_legacy_sha256(hex_str: str) -> bool:
    # Hash SHA-256 plano (64 hex)
    return (
        bool(hex_str)
        and len(hex_str) == 64
        and all(c in "0123456789abcdef" for c in hex_str.lower())
    )

def migrate_legacy_sha256(password: str, legacy_hex: str) -> Optional[str]:
    """Si el password coincide con el hash legacy (sha256 hex), devuelve PBKDF2 para migrar."""
    try:
        calc = hashlib.sha256(password.encode("utf-8")).hexdigest()
        if hmac.compare_digest(calc, legacy_hex):
            return pbkdf2_hash(password)
        return None
    except Exception:
        return None

# ========= User Store =========

@dataclass
class User:
    username: str
    password_hash: str  # PBKDF2 (preferido) o legacy (sha256 hex, se migra al validar)

class UserStore:
    def __init__(self, path: str = USERS_PATH):
        self.path = path
        self._users: Dict[str, User] = {}
        self._load()

    def _ensure_file(self):
        if not os.path.exists(self.path):
            # crea usuario admin por defecto con PBKDF2
            data = {"users": [{"username": "admin", "password_hash": pbkdf2_hash("admin")}]}
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

    def _load(self):
        self._ensure_file()
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                raw = json.load(f)
        except Exception:
            raw = {"users": []}
        users = {}
        for it in raw.get("users", []):
            u = str(it.get("username", "")).strip()
            ph = str(it.get("password_hash", "")).strip()
            if u and ph:
                users[u.lower()] = User(username=u, password_hash=ph)
        self._users = users

    def _save(self):
        data = {"users": [{"username": u.username, "password_hash": u.password_hash} for u in self._users.values()]}
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        os.replace(tmp, self.path)

    def verify(self, username: str, password: str) -> bool:
        u = self._users.get(username.lower())
        if not u:
            return False
        stored = u.password_hash

        # Ruta preferente (PBKDF2)
        if stored.startswith(f"{PBK_ALGO}$"):
            return pbkdf2_verify(password, stored)

        # Migración automática desde legacy sha256
        if is_legacy_sha256(stored):
            migrated = migrate_legacy_sha256(password, stored)
            if migrated:
                self._users[username.lower()] = User(username=u.username, password_hash=migrated)
                self._save()
                return True
        return False

# ========= Diálogo de Login =========

class LoginDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Acceso a FactuNabo")
        self.setModal(True)
        self.setMinimumWidth(420)
        self.settings = QSettings("FactuNabo", "Login")
        self.store = UserStore(USERS_PATH)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout(self)
        title = QLabel("FactuNabo – Inicio de sesión")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 20px; font-weight: 700; margin-bottom: 12px;")
        layout.addWidget(title)

        form = QFormLayout()
        self.username = QLineEdit(); self.username.setPlaceholderText("Usuario")
        self.password = QLineEdit(); self.password.setPlaceholderText("Contraseña"); self.password.setEchoMode(QLineEdit.Password)
        form.addRow("Usuario:", self.username)
        form.addRow("Contraseña:", self.password)
        layout.addLayout(form)

        self.remember = QCheckBox("Recordarme")
        layout.addWidget(self.remember)

        last_user = self.settings.value("last_user", "")
        if last_user:
            self.username.setText(last_user)
            self.remember.setChecked(True)

        btns = QHBoxLayout()
        btns.addStretch()
        btn_cancel = QPushButton("Cancelar")
        btn_login  = QPushButton("Entrar")
        btn_cancel.clicked.connect(self.reject)
        btn_login.clicked.connect(self.do_login)
        btns.addWidget(btn_cancel)
        btns.addWidget(btn_login)
        layout.addLayout(btns)

    def do_login(self):
        u = (self.username.text() or "").strip()
        p = self.password.text() or ""
        if not u or not p:
            QMessageBox.warning(self, "Login", "Indica usuario y contraseña.")
            return

        if not self.store.verify(u, p):
            QMessageBox.critical(self, "Login", "Usuario o contraseña incorrectos.")
            self.password.clear()
            self.password.setFocus()
            return

        if self.remember.isChecked():
            self.settings.setValue("last_user", u)
        else:
            self.settings.remove("last_user")

        # Opcional: exponer el usuario autenticado como variable de entorno para logs
        os.environ["FACTUNABO_USER"] = u
        self.accept()