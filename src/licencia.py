"""
Módulo de validación de licencias de KINETICA.

Flujo:
  1. Al iniciar, se carga la caché local (license.dat) si existe.
  2. Caché fresca (< CACHE_TTL_DAYS días): se permite el acceso sin tocar la red.
  3. Caché vencida (> CACHE_TTL_DAYS días): se re-valida online.
  4. Sin internet: período de gracia de OFFLINE_GRACE_DAYS días desde la última
     verificación exitosa. Al superar ese límite, se bloquea hasta recuperar conexión.
  5. Sin caché: la GUI solicita la clave al usuario y valida online.

La caché está protegida con HMAC para evitar edición manual trivial.
"""

import hashlib
import hmac
import json
import re
import uuid
import winreg
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Optional

import requests

# ── Endpoint del servidor de licencias ───────────────────────────────────────
# La URL real vive en src/_config.py (gitignoreado).
# Para desarrollo: copiá src/_config.py.example → src/_config.py y completá la URL.
try:
    from _config import LICENSE_ENDPOINT as _ENDPOINT
except ImportError:
    _ENDPOINT = ""

# ── Parámetros ────────────────────────────────────────────────────────────────
CACHE_TTL_DAYS     = 1   # días antes de re-validar online
OFFLINE_GRACE_DAYS = 7   # días sin internet antes de bloquear

_KEY_RE   = re.compile(r"^KINE-[A-HJ-NP-Z2-9]{4}-[A-HJ-NP-Z2-9]{4}-[A-HJ-NP-Z2-9]{4}$")
_HMAC_KEY = b"kinetica-lic-v1-\x3a\x7f\x21\xb8"  # ofuscación básica del caché local


# ── LicenseResult ─────────────────────────────────────────────────────────────

@dataclass
class LicenseResult:
    valid:            bool
    reason:           str  = ""    # key_not_found | expired | machine_mismatch |
                                   # offline_expired | connection_error | ok
    days_left:        int  = 0
    expires:          str  = ""    # ISO: YYYY-MM-DD
    first_activation: bool = False
    offline:          bool = False # True si se usó el período de gracia sin internet

    @property
    def expires_fmt(self) -> str:
        """Fecha de expiración en formato DD/MM/AAAA."""
        try:
            return datetime.strptime(self.expires, "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            return self.expires


# ── Machine ID ────────────────────────────────────────────────────────────────

def get_machine_id() -> str:
    """
    Genera un identificador estable y anónimo de la máquina.
    Usa el MachineGuid de Windows (registro) como fuente primaria.
    """
    parts: list[str] = []
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Cryptography",
        ) as k:
            parts.append(winreg.QueryValueEx(k, "MachineGuid")[0])
    except OSError:
        parts.append(str(uuid.getnode()))  # MAC address como fallback

    combined = "|".join(filter(None, parts)) or "fallback"
    return hashlib.sha256(combined.encode()).hexdigest()[:32]


# ── Clave: normalización y validación de formato ──────────────────────────────

def normalize_key(raw: str) -> str:
    """
    Normaliza la clave ingresada por el usuario:
    convierte a mayúsculas y reconstruye los guiones si los omitió.
    """
    clean = raw.strip().upper().replace(" ", "")
    if _KEY_RE.match(clean):
        return clean
    stripped = clean.replace("-", "")
    if stripped.startswith("KINE"):
        stripped = stripped[4:]
    if len(stripped) == 12:
        return f"KINE-{stripped[:4]}-{stripped[4:8]}-{stripped[8:]}"
    return clean  # lo devolvemos limpio; la validación online lo rechazará si es incorrecto

def is_valid_format(key: str) -> bool:
    return bool(_KEY_RE.match(key))


# ── Caché local ───────────────────────────────────────────────────────────────

def _sign(payload: dict) -> str:
    msg = json.dumps(payload, sort_keys=True, ensure_ascii=True).encode()
    return hmac.new(_HMAC_KEY, msg, hashlib.sha256).hexdigest()

def load_cache(path: Path) -> Optional[dict]:
    """Carga la caché y verifica su integridad. Retorna None si falta o es inválida."""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        sig  = data.pop("sig", "")
        if hmac.compare_digest(_sign(data), sig):
            return {**data, "sig": sig}
    except Exception:
        pass
    return None

def _save_cache(path: Path, key: str, expires: str, machine_id: str) -> None:
    payload = {
        "key":        key,
        "machine_id": machine_id,
        "expires":    expires,
        "last_check": date.today().isoformat(),
    }
    payload["sig"] = _sign({k: v for k, v in payload.items()})
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

def clear_cache(path: Path) -> None:
    path.unlink(missing_ok=True)


# ── Validación online ─────────────────────────────────────────────────────────

class ServerError(Exception):
    pass

def _validate_online(key: str, machine_id: str) -> LicenseResult:
    """Llama al servidor de licencias. Lanza ServerError si no hay conexión."""
    try:
        resp = requests.post(
            _ENDPOINT,
            json={"key": key, "machine_id": machine_id},
            timeout=12,
        )
        resp.raise_for_status()
        data = resp.json()
    except requests.exceptions.Timeout:
        raise ServerError("El servidor no respondió (timeout). Verificá tu conexión.")
    except requests.exceptions.ConnectionError:
        raise ServerError("Sin conexión a internet.")
    except Exception as e:
        raise ServerError(f"Error al contactar el servidor de licencias: {e}")

    if data.get("valid"):
        return LicenseResult(
            valid=True,
            days_left=int(data.get("days_left", 0)),
            expires=str(data.get("expires", "")),
            first_activation=bool(data.get("first_activation", False)),
        )
    return LicenseResult(valid=False, reason=data.get("reason", "unknown"))


# ── Verificación principal ────────────────────────────────────────────────────

def check(key: str, cache_path: Path, machine_id: Optional[str] = None) -> LicenseResult:
    """
    Verifica la licencia usando la caché cuando es posible.

    Prioridades:
      1. Caché fresca y válida → permite el acceso sin red.
      2. Caché vencida → re-valida online; actualiza la caché si sigue siendo válida.
      3. Sin internet → período de gracia de OFFLINE_GRACE_DAYS días.
      4. Sin caché válida → valida online obligatoriamente.
    """
    if machine_id is None:
        machine_id = get_machine_id()

    today = date.today()
    cache = load_cache(cache_path)

    # ── Intentar usar la caché existente ────────────────────────────────────
    if (
        cache
        and cache.get("key")        == key
        and cache.get("machine_id") == machine_id
    ):
        try:
            exp  = date.fromisoformat(cache["expires"])
            last = date.fromisoformat(cache["last_check"])
        except (KeyError, ValueError):
            cache = None  # caché corrupta, se ignora

        if cache:
            if today > exp:
                return LicenseResult(valid=False, reason="expired", expires=cache["expires"])

            days_left = (exp - today).days

            if (today - last).days < CACHE_TTL_DAYS:
                # Caché fresca: no hace falta tocar la red
                return LicenseResult(valid=True, days_left=days_left, expires=cache["expires"])

            # Caché vencida: intentar re-validación online
            try:
                result = _validate_online(key, machine_id)
                if result.valid:
                    _save_cache(cache_path, key, result.expires, machine_id)
                return result
            except ServerError:
                # Sin internet: conceder período de gracia
                stale_days = (today - last).days
                if stale_days <= OFFLINE_GRACE_DAYS:
                    return LicenseResult(
                        valid=True,
                        days_left=days_left,
                        expires=cache["expires"],
                        offline=True,
                    )
                return LicenseResult(valid=False, reason="offline_expired")

    # ── Sin caché válida: validación online obligatoria ──────────────────────
    try:
        result = _validate_online(key, machine_id)
    except ServerError:
        return LicenseResult(valid=False, reason="connection_error")

    if result.valid:
        _save_cache(cache_path, key, result.expires, machine_id)

    return result
