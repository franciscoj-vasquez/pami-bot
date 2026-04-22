"""
Tests para la lógica de bot.py que no requiere playwright ni PAMI:
  - leer_pacientes   (filtrado de filas vacías / NaN)
  - filtrado de NaN en prácticas (misma lógica que nueva_orden)
  - detección de fechas futuras (misma condición que el loop principal)
"""
from datetime import date, timedelta

import pandas as pd
import pytest

import bot as bot_module
from bot import leer_pacientes


# ── leer_pacientes ────────────────────────────────────────────────────────────

class TestLeerPacientes:
    def _excel(self, tmp_path, rows):
        path = tmp_path / "pacientes.xlsx"
        pd.DataFrame(rows).to_excel(path, index=False)
        return path

    def test_lee_filas_normales(self, tmp_path, monkeypatch):
        path = self._excel(tmp_path, [
            {"Beneficio": "111", "Fecha": "01/04/2025"},
            {"Beneficio": "222", "Fecha": "02/04/2025"},
        ])
        monkeypatch.setattr(bot_module, "EXCEL_PATH", path)
        assert len(leer_pacientes()) == 2

    def test_filtra_fila_beneficio_nan(self, tmp_path, monkeypatch):
        path = self._excel(tmp_path, [
            {"Beneficio": "111", "Fecha": "01/04/2025"},
            {"Beneficio": None,  "Fecha": None},
        ])
        monkeypatch.setattr(bot_module, "EXCEL_PATH", path)
        df = leer_pacientes()
        assert len(df) == 1
        assert df.iloc[0]["Beneficio"] == "111"

    def test_filtra_beneficio_solo_espacios(self, tmp_path, monkeypatch):
        path = self._excel(tmp_path, [
            {"Beneficio": "111", "Fecha": "01/04/2025"},
            {"Beneficio": "   ", "Fecha": "02/04/2025"},
        ])
        monkeypatch.setattr(bot_module, "EXCEL_PATH", path)
        assert len(leer_pacientes()) == 1

    def test_filtra_multiples_filas_vacias(self, tmp_path, monkeypatch):
        path = self._excel(tmp_path, [
            {"Beneficio": "111", "Fecha": "01/04/2025"},
            {"Beneficio": None,  "Fecha": None},
            {"Beneficio": None,  "Fecha": None},
            {"Beneficio": "222", "Fecha": "02/04/2025"},
        ])
        monkeypatch.setattr(bot_module, "EXCEL_PATH", path)
        assert len(leer_pacientes()) == 2

    def test_index_es_continuo_tras_filtrado(self, tmp_path, monkeypatch):
        path = self._excel(tmp_path, [
            {"Beneficio": "111", "Fecha": "01/04/2025"},
            {"Beneficio": None,  "Fecha": None},
            {"Beneficio": "222", "Fecha": "02/04/2025"},
        ])
        monkeypatch.setattr(bot_module, "EXCEL_PATH", path)
        df = leer_pacientes()
        assert list(df.index) == list(range(len(df)))

    def test_columnas_sin_espacios_extra(self, tmp_path, monkeypatch):
        # Simula columnas con espacios (caso real de Excel exportado)
        df_raw = pd.DataFrame([{"Beneficio": "111", "Fecha": "01/04/2025"}])
        df_raw.columns = [" Beneficio ", " Fecha "]
        path = tmp_path / "pacientes.xlsx"
        df_raw.to_excel(path, index=False)
        monkeypatch.setattr(bot_module, "EXCEL_PATH", path)
        df = leer_pacientes()
        assert "Beneficio" in df.columns
        assert " Beneficio " not in df.columns


# ── filtrado de NaN en prácticas ──────────────────────────────────────────────

def _practicas_validas(fila: pd.Series) -> list:
    """Replica exacta de la lógica en nueva_orden."""
    cols = sorted(c for c in fila.index if c.startswith("Cod_Practica"))
    return [str(fila[col]).strip() for col in cols
            if pd.notna(fila[col]) and str(fila[col]).strip()]


class TestFiltrarPracticas:
    def test_nan_excluido(self):
        fila = pd.Series({"Cod_Practica1": "250101", "Cod_Practica2": float("nan")})
        assert _practicas_validas(fila) == ["250101"]

    def test_cadena_vacia_excluida(self):
        fila = pd.Series({"Cod_Practica1": "250101", "Cod_Practica2": ""})
        assert _practicas_validas(fila) == ["250101"]

    def test_espacios_excluidos(self):
        fila = pd.Series({"Cod_Practica1": "250101", "Cod_Practica2": "   "})
        assert _practicas_validas(fila) == ["250101"]

    def test_multiples_validas(self):
        fila = pd.Series({
            "Cod_Practica1": "250101",
            "Cod_Practica2": "250102",
            "Cod_Practica3": "250103",
        })
        assert _practicas_validas(fila) == ["250101", "250102", "250103"]

    def test_todas_nan_retorna_vacio(self):
        fila = pd.Series({
            "Cod_Practica1": float("nan"),
            "Cod_Practica2": float("nan"),
        })
        assert _practicas_validas(fila) == []

    def test_mezcla_nan_y_validas(self):
        fila = pd.Series({
            "Cod_Practica1": "250101",
            "Cod_Practica2": float("nan"),
            "Cod_Practica3": "250103",
        })
        assert _practicas_validas(fila) == ["250101", "250103"]

    def test_orden_por_nombre_de_columna(self):
        # El sort es lexicográfico sobre el nombre de columna
        fila = pd.Series({
            "Cod_Practica3": "TRES",
            "Cod_Practica1": "UNO",
            "Cod_Practica2": "DOS",
        })
        assert _practicas_validas(fila) == ["UNO", "DOS", "TRES"]


# ── detección de fechas futuras ───────────────────────────────────────────────

def _es_futura(fecha_str: str) -> bool:
    """Replica exacta de la condición en el loop de run()."""
    return pd.to_datetime(fecha_str.strip(), dayfirst=True).date() > date.today()


class TestDeteccionFechaFutura:
    def test_manana_es_futura(self):
        manana = (date.today() + timedelta(days=1)).strftime("%d/%m/%Y")
        assert _es_futura(manana)

    def test_hoy_no_es_futura(self):
        hoy = date.today().strftime("%d/%m/%Y")
        assert not _es_futura(hoy)

    def test_ayer_no_es_futura(self):
        ayer = (date.today() - timedelta(days=1)).strftime("%d/%m/%Y")
        assert not _es_futura(ayer)

    def test_fecha_lejana_pasada_no_es_futura(self):
        assert not _es_futura("01/01/2020")

    def test_fecha_lejana_futura_es_futura(self):
        assert _es_futura("01/01/2099")
