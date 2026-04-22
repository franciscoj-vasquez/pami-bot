"""
Tests para las funciones puras de gui.py:
  - generar_fechas
  - cargar_feriados / guardar_feriados
"""
from datetime import datetime, date, timedelta

import gui as gui_module
from gui import cargar_feriados, generar_fechas, guardar_feriados


# ── generar_fechas ────────────────────────────────────────────────────────────

class TestGenerarFechas:
    def test_cantidad_exacta(self):
        inicio = datetime(2025, 1, 6)  # lunes
        assert len(generar_fechas(inicio, 10, [])) == 10

    def test_sin_sabados_ni_domingos(self):
        inicio = datetime(2025, 1, 6)
        for f in generar_fechas(inicio, 20, []):
            assert f.weekday() < 5, f"{f} es fin de semana"

    def test_inicio_sabado_arranca_el_lunes(self):
        sabado = datetime(2025, 1, 4)
        fechas = generar_fechas(sabado, 1, [])
        assert fechas[0] == date(2025, 1, 6)

    def test_inicio_domingo_arranca_el_lunes(self):
        domingo = datetime(2025, 1, 5)
        fechas = generar_fechas(domingo, 1, [])
        assert fechas[0] == date(2025, 1, 6)

    def test_excluye_feriado_unico(self):
        lunes = datetime(2025, 1, 6)
        fechas = generar_fechas(lunes, 1, ["06/01/2025"])
        assert fechas[0] == date(2025, 1, 7)  # saltea el lunes → martes

    def test_excluye_feriados_consecutivos(self):
        # Semana santa: lunes a jueves feriados
        lunes = datetime(2025, 4, 14)
        feriados = ["14/04/2025", "15/04/2025", "16/04/2025", "17/04/2025"]
        fechas = generar_fechas(lunes, 1, feriados)
        # Viernes 18 libre (no feriado en esta lista), debería ser la primera fecha
        assert fechas[0] == date(2025, 4, 18)

    def test_feriado_formato_invalido_ignorado(self):
        inicio = datetime(2025, 1, 6)
        # No debe lanzar excepción; los feriados inválidos se ignoran
        fechas = generar_fechas(inicio, 3, ["no_es_fecha", "99/99/9999", "2025-01-06"])
        assert len(fechas) == 3

    def test_sesiones_cero_retorna_lista_vacia(self):
        assert generar_fechas(datetime(2025, 1, 6), 0, []) == []

    def test_fechas_en_orden_ascendente(self):
        inicio = datetime(2025, 1, 6)
        fechas = generar_fechas(inicio, 5, [])
        assert fechas == sorted(fechas)

    def test_sin_duplicados(self):
        inicio = datetime(2025, 1, 6)
        fechas = generar_fechas(inicio, 10, [])
        assert len(fechas) == len(set(fechas))


# ── cargar_feriados / guardar_feriados ────────────────────────────────────────

class TestFeriados:
    def test_archivo_inexistente_retorna_lista_vacia(self, tmp_path, monkeypatch):
        monkeypatch.setattr(gui_module, "FERIADOS_FILE", tmp_path / "no_existe.json")
        assert cargar_feriados() == []

    def test_json_corrupto_retorna_lista_vacia(self, tmp_path, monkeypatch):
        f = tmp_path / "feriados.json"
        f.write_text("{json inválido")
        monkeypatch.setattr(gui_module, "FERIADOS_FILE", f)
        assert cargar_feriados() == []

    def test_json_es_dict_retorna_lista_vacia(self, tmp_path, monkeypatch):
        f = tmp_path / "feriados.json"
        f.write_text('{"fecha": "01/01/2025"}')
        monkeypatch.setattr(gui_module, "FERIADOS_FILE", f)
        assert cargar_feriados() == []

    def test_json_vacio_retorna_lista_vacia(self, tmp_path, monkeypatch):
        f = tmp_path / "feriados.json"
        f.write_text("[]")
        monkeypatch.setattr(gui_module, "FERIADOS_FILE", f)
        assert cargar_feriados() == []

    def test_carga_lista_valida(self, tmp_path, monkeypatch):
        f = tmp_path / "feriados.json"
        f.write_text('["01/01/2025", "25/12/2025"]')
        monkeypatch.setattr(gui_module, "FERIADOS_FILE", f)
        assert cargar_feriados() == ["01/01/2025", "25/12/2025"]

    def test_guardar_y_releer_round_trip(self, tmp_path, monkeypatch):
        f = tmp_path / "feriados.json"
        monkeypatch.setattr(gui_module, "FERIADOS_FILE", f)
        payload = ["01/05/2025", "25/12/2025"]
        guardar_feriados(payload)
        assert cargar_feriados() == payload

    def test_feriado_excluido_de_generar_fechas(self, tmp_path, monkeypatch):
        f = tmp_path / "feriados.json"
        monkeypatch.setattr(gui_module, "FERIADOS_FILE", f)
        guardar_feriados(["06/01/2025"])
        feriados = cargar_feriados()
        fechas = generar_fechas(datetime(2025, 1, 6), 1, feriados)
        assert date(2025, 1, 6) not in fechas
