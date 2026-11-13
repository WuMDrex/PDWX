import os
import json
import uuid
from datetime import datetime, timedelta
from cryptography.fernet import Fernet
import requests

# ==========================================
# CONFIGURACIÓN DEL SISTEMA Y RUTAS
# ==========================================
APP_FOLDER = os.path.join(os.getenv("APPDATA"), "SystemCache")
LICENSE_FILE = os.path.join(APP_FOLDER, "Sharedstorage.dat")

SECRET_KEY = b'Z3JYV3p0V2hQd0xqM1JtS3l4eFl2dGx5VjJpV0hVZ3A='  # NO CAMBIAR
cipher = Fernet(SECRET_KEY)

# Fecha de vencimiento base para nuevos registros
DEFAULT_EXPIRATION_DATE = "2027-02-15"

# URL de tu despliegue de Apps Script (Aplicación web)
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwuhcNdpLzIQB6kXUHMGfo2QW-bVZqPGNNec1gEzO-BFI2FEd3a4_YsEQLhdKKMyxZN/exec"

# Genera un ID único real por computadora (Basado en la dirección MAC)
def get_hardware_id():
    mac = uuid.getnode()
    return '-'.join(('%012X' % mac)[i:i+2] for i in range(0, 12, 2))

MY_HARDWARE_ID = get_hardware_id()


def _ensure_folder():
    os.makedirs(APP_FOLDER, exist_ok=True)


def _create_local_license(expiry, usuario="", password=""):
    data = {"expiry": expiry, "usuario": usuario, "password": password}
    encrypted = cipher.encrypt(json.dumps(data).encode('utf-8'))
    with open(LICENSE_FILE, "wb") as f:
        f.write(encrypted)


def _read_local_license():
    try:
        with open(LICENSE_FILE, "rb") as f:
            encrypted = f.read()
        decrypted = cipher.decrypt(encrypted)
        return json.loads(decrypted.decode())
    except Exception:
        return None


def _is_active(expiry_str):
    try:
        expiry = datetime.strptime(expiry_str, "%Y-%m-%d")
        return datetime.now() <= expiry
    except Exception:
        return False


def import_license(file_path):
    _ensure_folder()
    try:
        with open(file_path, "rb") as f:
            content = f.read()
        with open(LICENSE_FILE, "wb") as f:
            f.write(content)
        return True
    except Exception:
        return False


def check_license():
    """Verifica la licencia online (Apps Script) o localmente (respaldo)."""
    _ensure_folder()

    # --- PASO 1: INTENTO ONLINE ---
    try:
        response = requests.get(APPS_SCRIPT_URL, params={"id": MY_HARDWARE_ID}, timeout=10)
        if response.status_code == 200:
            api = response.json()
            print(f"[Licencias] Respuesta del servidor: {api.get('mensaje', 'sin mensaje')}")
            usuario = api.get("usuario", "")
            password = api.get("password", "")
            expiry = api.get("expiry_date", DEFAULT_EXPIRATION_DATE)
            estado = api.get("estado", "")
            active = str(estado).strip().lower() == "activo" and _is_active(expiry)

            _create_local_license(expiry=expiry, usuario=usuario, password=password)
            print(f"[Licencias] Archivo local actualizado. Estado: {estado}. Vence: {expiry}.")

            if not active:
                print("[Licencias] Licencia inactiva o expirada. No se permite iniciar.")
                return False
            return True
    except requests.RequestException:
        print("[Licencias] Sin conexión al servidor de licencias. Usando respaldo local.")

    # --- PASO 2: RESPALDO LOCAL ---
    local = _read_local_license()
    if local is None:
        # Día de gracia: sin internet y sin archivo local
        grace_date = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
        _create_local_license(expiry=grace_date)
        print(f"[Licencias] Sin conexión y sin registro previo. Se otorga 1 día de gracia hasta {grace_date}.")
        return True

    expiry = local.get("expiry", DEFAULT_EXPIRATION_DATE)
    if _is_active(expiry):
        print(f"[Licencias] Licencia local válida hasta {expiry}.")
        return True

    print(f"[Licencias] Licencia local expirada (vencía {expiry}). No se permite iniciar.")
    return False


if __name__ == "__main__":
    print(f"Tu Hardware ID único es: {MY_HARDWARE_ID}")
    print("Iniciando validación...")
    if check_license():
        print("[OK] Licencia válida. Puedes continuar.")
    else:
        print("[ERROR] Licencia expirada o bloqueada.")
