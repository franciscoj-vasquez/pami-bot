"""
Verificación de afiliados PAMI.

Permite consultar el nombre de cada afiliado usando una única sesión headless de
Playwright, pagando el costo de login y navegación una sola vez para toda la lista:

    login → ambulatorio → ALTA → [búsqueda × N] → cancelar → cerrar

Cada búsqueda individual abre el popup de afiliado, lee el primer resultado y cierra
el popup con Escape sin seleccionarlo, dejando el formulario en estado limpio para
la siguiente consulta.
"""

from __future__ import annotations

import random
import re
import threading
import time
from dataclasses import dataclass
from typing import Callable, Optional

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout


_URL_LOGIN       = "https://efectoresweb.pami.org.ar/EfectoresWeb/login.isp"
_URL_AMBULATORIO = "https://efectoresweb.pami.org.ar/EfectoresWeb/ambulatorio.isp"

_SLOW_MO   = 150
_TYPING_MS = 80
_PAUSA_MIN = 0.4
_PAUSA_MAX = 1.0
_CORTA_MIN = 0.2
_CORTA_MAX = 0.5


def _pausa(mn: float = None, mx: float = None) -> None:
    time.sleep(random.uniform(
        mn if mn is not None else _PAUSA_MIN,
        mx if mx is not None else _PAUSA_MAX,
    ))


def _pausa_corta() -> None:
    time.sleep(random.uniform(_CORTA_MIN, _CORTA_MAX))


# ── Tipos ─────────────────────────────────────────────────────────────────────

@dataclass
class ResultadoAfiliado:
    """Resultado de la verificación de un único afiliado."""

    beneficio:  str
    parentesco: str
    nombre:     Optional[str] = None  # None → no encontrado o error técnico
    error:      Optional[str] = None  # mensaje de error; None si la búsqueda fue exitosa

    @property
    def encontrado(self) -> bool:
        return self.nombre is not None


class VerificacionError(Exception):
    """Error recuperable durante la verificación de un paciente."""


class LoginError(VerificacionError):
    """Fallo de autenticación — detiene toda la sesión."""


ProgressCallback = Callable[[int, int, ResultadoAfiliado], None]


# ── Interacción con PAMI ──────────────────────────────────────────────────────

def _login(page, usuario: str, clave: str) -> None:
    page.goto(_URL_LOGIN)
    page.wait_for_selector('input[type="text"]', timeout=15000)
    page.locator('input[type="text"]').first.fill(usuario)
    page.locator('input[type="password"]').first.fill(clave)
    page.get_by_role("button", name="INICIAR SESION").click()
    try:
        page.wait_for_selector("text=Usuario y/o contraseña incorrecta.", timeout=4000)
        page.locator(".z-messagebox-button").click()
        raise LoginError("Usuario y/o contraseña incorrecta. Verificá las credenciales.")
    except PWTimeout:
        pass  # login exitoso


def _abrir_formulario_alta(page) -> None:
    page.goto(_URL_AMBULATORIO)
    page.wait_for_selector("text=ALTA", timeout=15000)
    _pausa()
    page.locator("text=ALTA").first.click()
    page.wait_for_selector("#zk_comp_130-btn", state="visible", timeout=20000)
    _pausa()


def _cancelar_alta(page) -> None:
    try:
        page.locator("#zk_comp_318").click()
        page.wait_for_selector("text=ALTA", state="visible", timeout=10000)
    except Exception:
        pass


def _consultar_nombre(
    page,
    beneficio:  str,
    parentesco: str,
    stop:       threading.Event,
) -> Optional[str]:
    """
    Abre el popup de búsqueda de afiliado, ingresa beneficio + parentesco, busca,
    y retorna el texto del primer resultado sin seleccionarlo.

    Cierra el popup con Escape para no alterar el estado del formulario, permitiendo
    reutilizarlo para la siguiente consulta sin cancelar la orden.

    Returns:
        Texto de las celdas del primer resultado (celdas separadas por "  |  "),
        o None si el afiliado no se encontró.
    Raises:
        VerificacionError: si hay un problema técnico con la UI de PAMI.
    """
    if stop.is_set():
        return None

    page.locator("#zk_comp_130-btn").click()
    popup = page.locator("#zk_comp_130-pp")
    try:
        popup.wait_for(state="visible", timeout=10000)
    except PWTimeout:
        raise VerificacionError("El panel de búsqueda de afiliado no respondió.")
    _pausa()

    if stop.is_set():
        page.keyboard.press("Escape")
        return None

    # Ingresar número de beneficio
    campo = page.locator("#zk_comp_153")
    campo.click()
    campo.press("Control+a")
    campo.press("Delete")
    campo.type(beneficio, delay=_TYPING_MS)
    _pausa_corta()

    # Seleccionar parentesco en el combobox
    cod_par = parentesco.split(" - ")[0] if " - " in parentesco else parentesco
    cod_par = cod_par.strip().zfill(2)

    par_btn = popup.locator("tr").filter(has_text="Parentesco").locator(".z-combobox-button")
    try:
        par_btn.wait_for(state="visible", timeout=10000)
    except PWTimeout:
        page.keyboard.press("Escape")
        raise VerificacionError("El selector de parentesco no respondió.")
    par_btn.click()
    _pausa_corta()

    dropdown = page.locator(".z-combobox-popup.z-combobox-open")
    try:
        dropdown.wait_for(state="visible", timeout=5000)
    except PWTimeout:
        page.keyboard.press("Escape")
        raise VerificacionError("El dropdown de parentesco no se abrió.")

    items_par = dropdown.locator(
        ".z-comboitem-text", has_text=re.compile(rf"^{cod_par}")
    )
    if items_par.count() == 0:
        page.keyboard.press("Escape")
        _pausa_corta()
        raise VerificacionError(f"Parentesco '{cod_par}' no reconocido por PAMI.")
    items_par.first.click()
    _pausa()

    if stop.is_set():
        page.keyboard.press("Escape")
        return None

    popup.get_by_role("button", name="Buscar").click()

    primer_item = popup.locator(".z-listitem").first
    try:
        primer_item.wait_for(state="visible", timeout=12000)
    except PWTimeout:
        # El afiliado no fue encontrado — no es un error técnico
        page.keyboard.press("Escape")
        _pausa_corta()
        return None

    # Leer todas las celdas del primer resultado y combinarlas
    cells = primer_item.locator(".z-listcell")
    nombre = "  |  ".join(
        cells.nth(i).inner_text().strip()
        for i in range(cells.count())
        if cells.nth(i).inner_text().strip()
    )

    page.keyboard.press("Escape")
    _pausa_corta()
    return nombre or None


# ── API pública ───────────────────────────────────────────────────────────────

def verificar_afiliados_batch(
    usuario:     str,
    clave:       str,
    pacientes:   list[dict],
    on_progress: Optional[ProgressCallback] = None,
    stop:        Optional[threading.Event] = None,
) -> list[ResultadoAfiliado]:
    """
    Verifica el nombre de todos los afiliados en una única sesión de browser.

    Flujo: login → ambulatorio → ALTA → [búsqueda × N] → cancelar → cerrar.
    El overhead de login y navegación se paga una sola vez para toda la lista.

    Args:
        usuario:     credencial PAMI.
        clave:       credencial PAMI.
        pacientes:   lista de dicts con claves ``"beneficio"`` y ``"parentesco"``.
        on_progress: callback(idx_0based, total, resultado) invocado tras cada verificación.
        stop:        Event para cancelación externa desde otro thread.

    Returns:
        Lista de :class:`ResultadoAfiliado` en el mismo orden que ``pacientes``.

    Raises:
        LoginError: si las credenciales son incorrectas (antes de cualquier búsqueda).
        VerificacionError: si ocurre un error fatal antes de comenzar las búsquedas.
    """
    if stop is None:
        stop = threading.Event()

    resultados: list[ResultadoAfiliado] = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True, slow_mo=_SLOW_MO)
        page    = browser.new_context().new_page()

        try:
            _login(page, usuario, clave)

            if stop.is_set():
                return resultados

            _abrir_formulario_alta(page)

            for idx, p in enumerate(pacientes):
                if stop.is_set():
                    break

                beneficio  = str(p.get("beneficio", "")).strip()
                parentesco = str(p.get("parentesco", "")).strip()
                resultado  = ResultadoAfiliado(beneficio=beneficio, parentesco=parentesco)

                try:
                    nombre = _consultar_nombre(page, beneficio, parentesco, stop)
                    if nombre:
                        resultado.nombre = nombre
                    else:
                        resultado.error = "No encontrado en PAMI"
                except VerificacionError as e:
                    resultado.error = str(e)

                resultados.append(resultado)
                if on_progress:
                    on_progress(idx, len(pacientes), resultado)

            if not stop.is_set():
                _cancelar_alta(page)

        finally:
            browser.close()

    return resultados
