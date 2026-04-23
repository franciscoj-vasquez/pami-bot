import customtkinter as ctk
import json
import keyring
import os
import re
import subprocess
import sys
import threading
import tkinter as tk
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import messagebox

KEYRING_SERVICE = "pami_bot"

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

DATA_DIR          = Path(__file__).parent.parent / "data"
DATA_DIR.mkdir(exist_ok=True)
FERIADOS_FILE     = DATA_DIR / "feriados.json"
PACIENTES_FILE    = DATA_DIR / "pacientes_estado.json"
EXCEL_PATH        = DATA_DIR / "pacientes.xlsx"
STOP_FLAG         = DATA_DIR / "stop.flag"
DEFAULT_PRACTICAS    = ["250101", "250102"]
COLUMNAS_REQUERIDAS  = {"Beneficio", "Parentesco", "Fecha", "Cod_Diagnostico"}

# ── Helpers ───────────────────────────────────────────────────────────────────

def cargar_feriados():
    if FERIADOS_FILE.exists():
        try:
            with open(FERIADOS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, list) else []
        except (json.JSONDecodeError, OSError):
            return []
    return []

def guardar_feriados(feriados):
    with open(FERIADOS_FILE, "w", encoding="utf-8") as f:
        json.dump(feriados, f)

def cargar_pacientes_guardados():
    if not PACIENTES_FILE.exists():
        return []
    try:
        with open(PACIENTES_FILE, "r") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return []

    if not isinstance(data, list):
        return []

    resultado = []
    for p in data:
        try:
            p["fecha_inicio"] = datetime.strptime(p["fecha_inicio"], "%d/%m/%Y")
            resultado.append(p)
        except (ValueError, KeyError):
            print(f"[AVISO] Paciente ignorado por fecha inválida: {p.get('beneficio', '?')}")
    return resultado

def validar_excel(path):
    try:
        df = pd.read_excel(path, dtype=str)
    except Exception as e:
        return False, f"No se pudo leer el archivo: {e}"

    df.columns = df.columns.str.strip()
    cols = set(df.columns)

    faltantes = COLUMNAS_REQUERIDAS - cols
    if faltantes:
        return False, f"Columnas faltantes: {', '.join(sorted(faltantes))}"

    if not any(c.startswith("Cod_Practica") for c in cols):
        return False, "No se encontró ninguna columna 'Cod_Practica'."

    filas_validas = df[df["Beneficio"].notna() & (df["Beneficio"].str.strip() != "")]
    if filas_validas.empty:
        return False, "El archivo no contiene filas con datos válidos."

    return True, ""

def guardar_pacientes_estado(pacientes):
    data = []
    for p in pacientes:
        entry = p.copy()
        entry["fecha_inicio"] = p["fecha_inicio"].strftime("%d/%m/%Y")
        data.append(entry)
    with open(PACIENTES_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def generar_fechas(fecha_inicio: datetime, cantidad: int, feriados: list) -> list:
    feriados_dt = set()
    for f in feriados:
        try:
            feriados_dt.add(datetime.strptime(f, "%d/%m/%Y").date())
        except ValueError:
            pass

    fechas, fecha = [], fecha_inicio.date()
    while len(fechas) < cantidad:
        if fecha.weekday() < 5 and fecha not in feriados_dt:
            fechas.append(fecha)
        fecha += timedelta(days=1)
    return fechas

# ── Tooltip ───────────────────────────────────────────────────────────────────

class Tooltip:
    def __init__(self, widget, text):
        self._widget = widget
        self._text   = text
        self._tip    = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _=None):
        if self._tip:
            self._tip.destroy()
        x = self._widget.winfo_rootx() + 20
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 4
        self._tip = tk.Toplevel(self._widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self._tip, text=self._text, background="#2b2b2b",
                 foreground="white", relief="flat", padx=8, pady=4,
                 font=("Segoe UI", 10)).pack()

    def _hide(self, _=None):
        if self._tip:
            self._tip.destroy()
            self._tip = None

# ── Widget: Fecha con tres campos ─────────────────────────────────────────────

class DateEntry(ctk.CTkFrame):
    def __init__(self, parent, width=160, **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)

        self._vars    = [ctk.StringVar(), ctk.StringVar(), ctk.StringVar()]
        self._maxlens = [2, 2, 4]
        self._fields  = []

        for i, (w, ml, var) in enumerate(zip([36, 36, 56], self._maxlens, self._vars)):
            entry = ctk.CTkEntry(self, width=w, justify="center", textvariable=var)
            entry.pack(side="left")
            if i < 2:
                ctk.CTkLabel(self, text="/", width=10,
                             text_color="gray60").pack(side="left")
            self._fields.append(entry)
            var.trace_add("write", lambda *_, idx=i: self._on_change(idx))

        for i in range(1, 3):
            self._fields[i]._entry.bind(
                "<BackSpace>", lambda e, idx=i: self._on_backspace(idx))

    def _on_change(self, i):
        var = self._vars[i]
        val = ''.join(c for c in var.get() if c.isdigit())[:self._maxlens[i]]
        if val != var.get():
            var.set(val)
            return
        if len(val) == self._maxlens[i] and i < 2:
            self._fields[i + 1]._entry.focus_set()
            self._fields[i + 1]._entry.select_range(0, "end")

    def _on_backspace(self, i):
        if not self._vars[i].get():
            self._fields[i - 1]._entry.focus_set()
            self._fields[i - 1]._entry.icursor("end")

    def get(self):
        return "/".join(v.get() for v in self._vars)

    def set(self, value):
        parts = (value or "").split("/")
        if len(parts) == 3:
            for var, part in zip(self._vars, parts):
                var.set(part)

# ── Widget: Parentesco ────────────────────────────────────────────────────────

class ParentescoWidget(ctk.CTkFrame):
    DESCRIPCIONES = {
        "00": "TITULAR",
        "01": "ESPOSO/A",
        **{f"{i:02d}": "HIJO/A" for i in range(2, 16)},
        **{f"{i:02d}": "HIJO/A DISCAPACITADO" for i in range(16, 21)},
        "21": "PADRE/MADRE",
        "27": "SOBRINO/A",
        "38": "NIETO DISCAPACITADO/A",
        "39": "2da UNION CONVIVENCIAL",
        "40": "MENOR BAJO GUARDA",
        "42": "HIJO/A DEL CONYUGE",
        "44": "PADRE/MADRE",
        "68": "HIJO DISCAPACITADO DEL CONYUGE",
        "69": "HIJO/A DE HIJO/A MENOR A CARGO",
        "70": "HIJO DISCAPACITADO DEL CONYUGE",
        "74": "NIETO/A DISCAPACITADO",
        "75": "SEGUNDA ESPOSO/A",
        "76": "MENOR BAJO GUARDA",
        "77": "MENOR BAJO GUARDA",
        "78": "HIJO/A",
        "82": "HIJO/A DISCAPACITADO",
        "83": "HIJO/A DISCAPACITADO",
        "85": "HIJO/A DISCAPACITADO",
        "87": "HIJO/A DEL CONYUGE",
        "88": "HIJO/A DEL CONYUGE",
        "89": "HIJO/A DEL CONYUGE",
        **{f"{i:02d}": "MENOR BAJO GUARDA" for i in range(90, 96)},
        **{f"{i:02d}": "PUPILO/TUTELADO SUJ. A CURATELA" for i in range(96, 99)},
        "99": "UNION CONVIVENCIAL",
    }

    _MAX = 16

    def __init__(self, parent, default="00", **kwargs):
        super().__init__(parent, fg_color="transparent", **kwargs)

        self._var = ctk.StringVar(value=default)
        self._var.trace_add("write", self._actualizar_desc)

        self.entry = ctk.CTkEntry(self, width=45, textvariable=self._var)
        self.entry.pack(side="left")

        display, full = self._desc(default)
        self._label = ctk.CTkLabel(self, text=display, text_color="gray55",
                                   anchor="w", width=107)
        self._label.pack(side="left", padx=(8, 0))
        self._tooltip = Tooltip(self._label, full) if len(full) > self._MAX else None

    def _desc(self, cod):
        d = self.DESCRIPCIONES.get(cod.strip(), "")
        if not d:
            return "", ""
        truncado = d if len(d) <= self._MAX else d[:self._MAX - 1] + "…"
        return f"— {truncado}", d

    def _actualizar_desc(self, *_):
        display, full = self._desc(self._var.get())
        self._label.configure(text=display)
        if len(full) > self._MAX:
            self._tooltip = Tooltip(self._label, full)
        else:
            self._tooltip = None

    def get(self):
        return self._var.get().strip()

    def set(self, value):
        cod = value.split(" - ")[0] if " - " in value else value
        self._var.set(cod)

# ── Diálogo: Credenciales ─────────────────────────────────────────────────────

class CredencialesDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Credenciales PAMI")
        self.geometry("340x260")
        self.resizable(False, False)
        self.grab_set()

        guardadas = keyring.get_password(KEYRING_SERVICE, "usuario") is not None

        ctk.CTkLabel(self, text="Usuario:").grid(row=0, column=0, padx=20, pady=(20,5), sticky="w")
        self.entry_usuario = ctk.CTkEntry(self, width=200)
        self.entry_usuario.grid(row=0, column=1, padx=10, pady=(20,5))
        self.entry_usuario.insert(0, parent.usuario)

        ctk.CTkLabel(self, text="Contraseña:").grid(row=1, column=0, padx=20, pady=5, sticky="w")
        frame_clave = ctk.CTkFrame(self, fg_color="transparent")
        frame_clave.grid(row=1, column=1, padx=10, pady=5)
        self.entry_clave = ctk.CTkEntry(frame_clave, width=162, show="*")
        self.entry_clave.pack(side="left")
        self.entry_clave.insert(0, parent.clave)
        ctk.CTkButton(frame_clave, text="👁", width=30, fg_color="transparent",
                      border_width=1, hover_color="#444",
                      command=self._toggle_clave).pack(side="left", padx=(4,0))

        self.recordar_var = ctk.BooleanVar(value=guardadas)
        ctk.CTkCheckBox(self, text="Recordar credenciales", variable=self.recordar_var).grid(
            row=2, column=0, columnspan=2, pady=(10,0))

        ctk.CTkButton(self, text="Guardar", command=self.guardar).grid(
            row=3, column=0, columnspan=2, pady=(10,4))

        ctk.CTkButton(self, text="Olvidar credenciales", fg_color="#c0392b",
                      hover_color="#922b21", command=self.olvidar).grid(
                      row=4, column=0, columnspan=2, pady=(0,16))

    def _toggle_clave(self):
        self.entry_clave.configure(show="" if self.entry_clave.cget("show") == "*" else "*")

    def guardar(self):
        self.parent.usuario = self.entry_usuario.get().strip()
        self.parent.clave   = self.entry_clave.get()
        if self.recordar_var.get():
            keyring.set_password(KEYRING_SERVICE, "usuario", self.parent.usuario)
            keyring.set_password(KEYRING_SERVICE, "clave",   self.parent.clave)
            self.parent.log_append("Credenciales guardadas en el sistema.\n")
        else:
            self._borrar_keyring()
        self.destroy()

    def olvidar(self):
        self._borrar_keyring()
        self.parent.usuario = ""
        self.parent.clave   = ""
        self.parent.log_append("Credenciales olvidadas.\n")
        self.destroy()

    def _borrar_keyring(self):
        for key in ("usuario", "clave"):
            try:
                keyring.delete_password(KEYRING_SERVICE, key)
            except keyring.errors.PasswordDeleteError:
                pass

# ── Diálogo: Feriados ─────────────────────────────────────────────────────────

class FeriadosDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Días a excluir")
        self.geometry("320x420")
        self.resizable(False, False)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.cerrar)

        self.feriados = cargar_feriados()

        ctk.CTkLabel(self, text="Días a excluir (DD/MM/AAAA):").pack(pady=(15,5))

        self.lista = ctk.CTkScrollableFrame(self, height=220)
        self.lista.pack(fill="x", padx=15)
        self._renderizar_lista()

        frame_add = ctk.CTkFrame(self, fg_color="transparent")
        frame_add.pack(pady=10)
        self.entry_fecha = DateEntry(frame_add, width=140)
        self.entry_fecha.set(datetime.today().strftime("%d/%m/%Y"))
        self.entry_fecha.pack(side="left", padx=(0,8))
        ctk.CTkButton(frame_add, text="Agregar", width=80, command=self.agregar).pack(side="left")

        ctk.CTkButton(self, text="Cerrar", command=self.cerrar).pack(pady=10)

    def _renderizar_lista(self):
        for w in self.lista.winfo_children():
            w.destroy()
        for fecha in self.feriados:
            fila = ctk.CTkFrame(self.lista, fg_color="transparent")
            fila.pack(fill="x", pady=2)
            ctk.CTkLabel(fila, text=fecha, width=180).pack(side="left")
            ctk.CTkButton(fila, text="✕", width=30, fg_color="#c0392b",
                          command=lambda f=fecha: self.quitar(f)).pack(side="right")

    def agregar(self):
        fecha = self.entry_fecha.get().strip()
        try:
            datetime.strptime(fecha, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error", "Formato inválido. Usá DD/MM/AAAA", parent=self)
            return
        if fecha not in self.feriados:
            self.feriados.append(fecha)
            self.feriados.sort(key=lambda f: datetime.strptime(f, "%d/%m/%Y"))
        self.entry_fecha.set(datetime.today().strftime("%d/%m/%Y"))
        self._renderizar_lista()

    def quitar(self, fecha):
        self.feriados.remove(fecha)
        self._renderizar_lista()

    def cerrar(self):
        guardar_feriados(self.feriados)
        self.destroy()

# ── Diálogo: Agregar / Editar Paciente ───────────────────────────────────────

class PacienteDialog(ctk.CTkToplevel):
    def __init__(self, parent, idx=None):
        super().__init__(parent)
        self.parent = parent
        self.idx    = idx
        self._practica_rows = []

        es_edicion = idx is not None
        self.title("Editar Paciente" if es_edicion else "Agregar Paciente")
        self.geometry("420x440")
        self.resizable(False, True)
        self.grab_set()

        hoy = datetime.today().strftime("%d/%m/%Y")
        if es_edicion:
            p    = parent.pacientes[idx]
            vals = {
                "beneficio":   p["beneficio"],
                "parentesco":  p["parentesco"],
                "diagnostico": p["diagnostico"],
                "fecha":       p["fecha_inicio"].strftime("%d/%m/%Y"),
                "sesiones":    str(p["sesiones"]),
                "practicas":   p.get("practicas", DEFAULT_PRACTICAS),
            }
        else:
            vals = {
                "beneficio":   "",
                "parentesco":  "00",
                "diagnostico": "",
                "fecha":       hoy,
                "sesiones":    "",
                "practicas":   DEFAULT_PRACTICAS,
            }

        campos = [
            ("Beneficio:",            "entry_beneficio",   vals["beneficio"],   "entry"),
            ("Parentesco:",           "entry_parentesco",  vals["parentesco"],  "parentesco"),
            ("Cód. Diagnóstico:",     "entry_diagnostico", vals["diagnostico"], "entry"),
            ("Fecha inicio:",         "entry_fecha",       vals["fecha"],       "date"),
            ("Cantidad de sesiones:", "entry_sesiones",    vals["sesiones"],    "entry"),
        ]

        for i, (label, attr, default, tipo) in enumerate(campos):
            ctk.CTkLabel(self, text=label, anchor="w", width=200).grid(
                row=i, column=0, padx=(20,5), pady=(12,0), sticky="w")
            if tipo == "parentesco":
                widget = ParentescoWidget(self, default=default)
            elif tipo == "date":
                widget = DateEntry(self, width=160)
                widget.set(default)
            else:
                widget = ctk.CTkEntry(self, width=160)
                widget.insert(0, default)
            widget.grid(row=i, column=1, padx=(0,20), pady=(12,0), sticky="w")
            setattr(self, attr, widget)

        n = len(campos)
        ctk.CTkLabel(self, text="Códigos de práctica:", anchor="w", width=200).grid(
            row=n, column=0, padx=(20,5), pady=(12,0), sticky="nw")
        self.practicas_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.practicas_frame.grid(row=n, column=1, padx=(0,20), pady=(12,0), sticky="w")

        self.btn_agregar_cod = ctk.CTkButton(
            self.practicas_frame, text="+ Agregar código", width=145,
            command=self._agregar_entrada_practica)
        self.btn_agregar_cod.pack(pady=(2,0))

        for cod in vals["practicas"]:
            self._agregar_entrada_practica(cod)

        label_btn = "Guardar" if es_edicion else "Agregar"
        ctk.CTkButton(self, text=label_btn, command=self._guardar, width=160).grid(
            row=n+1, column=0, columnspan=2, pady=16)

    def _agregar_entrada_practica(self, cod=""):
        fila = ctk.CTkFrame(self.practicas_frame, fg_color="transparent")
        fila.pack(fill="x", pady=2, before=self.btn_agregar_cod)
        entry = ctk.CTkEntry(fila, width=110)
        entry.insert(0, cod)
        entry.pack(side="left")
        ref = [fila, entry]
        self._practica_rows.append(ref)
        ctk.CTkButton(fila, text="✕", width=28, fg_color="#c0392b",
                      command=lambda r=ref: self._quitar_practica(r)).pack(side="left", padx=(4,0))

    def _quitar_practica(self, ref):
        ref[0].destroy()
        self._practica_rows.remove(ref)

    def _guardar(self):
        beneficio   = self.entry_beneficio.get().strip()
        parentesco  = self.entry_parentesco.get().strip()
        diagnostico = self.entry_diagnostico.get().strip().upper()
        fecha_str   = self.entry_fecha.get().strip()
        sesiones_str = self.entry_sesiones.get().strip()

        try:
            sesiones = int(sesiones_str)
        except ValueError:
            messagebox.showerror("Error", "La cantidad de sesiones debe ser un número entero.", parent=self)
            return

        try:
            fecha = datetime.strptime(fecha_str, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error", f"La fecha '{fecha_str}' no es válida. Usá el formato DD/MM/AAAA.", parent=self)
            return

        practicas = [r[1].get().strip() for r in self._practica_rows if r[1].get().strip()]

        if sesiones < 1:
            messagebox.showerror("Error", "La cantidad de sesiones debe ser al menos 1.", parent=self)
            return

        if not all([beneficio, parentesco, diagnostico]) or not practicas:
            messagebox.showerror("Error", "Completá todos los campos e ingresá al menos una práctica.", parent=self)
            return

        hace_3_meses = datetime.today() - timedelta(days=90)
        if fecha < hace_3_meses:
            if not messagebox.askyesno(
                "Fecha lejana en el pasado",
                f"La fecha de inicio es {fecha_str}, hace más de 3 meses.\n\n¿Es correcta?",
                parent=self
            ):
                return

        paciente = {
            "beneficio":    beneficio,
            "parentesco":   parentesco,
            "diagnostico":  diagnostico,
            "fecha_inicio": fecha,
            "sesiones":     sesiones,
            "practicas":    practicas,
        }
        if self.idx is None:
            self.parent.pacientes.append(paciente)
        else:
            self.parent.pacientes[self.idx] = paciente

        self.parent.actualizar_tabla()
        self.destroy()

# ── Diálogo: Fechas Futuras ───────────────────────────────────────────────────

class FechasFuturasDialog(ctk.CTkToplevel):
    def __init__(self, parent, por_beneficio, max_fecha, on_confirm):
        super().__init__(parent)
        self._on_confirm = on_confirm
        self.title("Sesiones con fecha futura")
        self.geometry("440x460")
        self.resizable(False, True)
        self.grab_set()

        total = sum(len(v) for v in por_beneficio.values())

        ctk.CTkLabel(self,
                     text=f"{total} sesión(es) de {len(por_beneficio)} paciente(s) tienen fecha futura.",
                     wraplength=400, font=ctk.CTkFont(weight="bold")).pack(pady=(18, 4), padx=20)

        ctk.CTkLabel(self,
                     text=f"Podrás procesar el Excel completo a partir del {max_fecha.strftime('%d/%m/%Y')}.",
                     wraplength=400, text_color="#e67e22").pack(pady=(0, 10), padx=20)

        lista = ctk.CTkScrollableFrame(self, height=140)
        lista.pack(fill="x", padx=20, pady=(0, 8))
        for ben, fechas in por_beneficio.items():
            texto = (f"• Beneficio {ben}: {len(fechas)} sesión(es),"
                     f" hasta el {max(fechas).strftime('%d/%m/%Y')}")
            ctk.CTkLabel(lista, text=texto, anchor="w").pack(fill="x", padx=6, pady=2)

        ctk.CTkLabel(self,
                     text="Estas sesiones se omitirán automáticamente al ejecutar el bot.",
                     wraplength=400, text_color="gray60",
                     font=ctk.CTkFont(size=11)).pack(pady=(0, 14), padx=20)

        frame_btns = ctk.CTkFrame(self, fg_color="transparent")
        frame_btns.pack(pady=(0, 18))
        ctk.CTkButton(frame_btns, text="Generar de todas formas",
                      command=self._confirmar).pack(side="left", padx=8)
        ctk.CTkButton(frame_btns, text="Cancelar", fg_color="#c0392b",
                      hover_color="#922b21", command=self.destroy).pack(side="left", padx=8)

    def _confirmar(self):
        self.destroy()
        self._on_confirm()

# ── Ventana Principal ─────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("PAMI - Carga de Órdenes")
        self.geometry("700x580")
        self.minsize(700, 580)
        self.resizable(True, True)

        self.usuario   = keyring.get_password(KEYRING_SERVICE, "usuario") or ""
        self.clave     = keyring.get_password(KEYRING_SERVICE, "clave")   or ""
        self.pacientes = cargar_pacientes_guardados()
        self._proc     = None
        self._hide_after_id = None

        self._build_ui()

        if self.pacientes:
            self.actualizar_tabla()

        if self.usuario:
            self.log_append(f"Credenciales cargadas para '{self.usuario}'.\n")

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=2)  # tabla crece más
        self.grid_rowconfigure(8, weight=1)  # log crece menos

        # ── Título
        ctk.CTkLabel(self, text="PAMI — Carga de Órdenes",
                     font=ctk.CTkFont(size=18, weight="bold")).grid(
                     row=0, column=0, pady=(20,5))

        # ── Botones de configuración
        frame_config = ctk.CTkFrame(self, fg_color="transparent")
        frame_config.grid(row=1, column=0, pady=5)
        ctk.CTkButton(frame_config, text="Credenciales",
                      command=self.abrir_credenciales).pack(side="left", padx=8)
        ctk.CTkButton(frame_config, text="Días a excluir",
                      command=self.abrir_feriados).pack(side="left", padx=8)

        # ── Tabla de pacientes
        frame_tabla_header = ctk.CTkFrame(self, fg_color="transparent")
        frame_tabla_header.grid(row=2, column=0, sticky="ew", padx=20, pady=(15,2))
        ctk.CTkLabel(frame_tabla_header, text="Pacientes cargados:", anchor="w").pack(side="left")
        ctk.CTkButton(frame_tabla_header, text="Limpiar todo", width=90, height=22,
                      fg_color="transparent", border_width=1, hover_color="#444",
                      text_color="#e67e22", font=ctk.CTkFont(size=11),
                      command=self.limpiar_pacientes).pack(side="right")

        self.tabla_frame = ctk.CTkScrollableFrame(self, height=180)
        self.tabla_frame.grid(row=3, column=0, sticky="nsew", padx=20)
        self._headers_tabla()

        # ── Botones de tabla (agregar, limpiar, generar)
        frame_tabla_btns = ctk.CTkFrame(self, fg_color="transparent")
        frame_tabla_btns.grid(row=4, column=0, pady=(8, 4))
        ctk.CTkButton(frame_tabla_btns, text="+ Agregar Paciente",
                      command=self.abrir_agregar_paciente).pack(side="left", padx=8)
        ctk.CTkButton(frame_tabla_btns, text="Generar Excel",
                      command=self.generar_excel).pack(side="left", padx=8)

        # ── Sección Ejecutar Bot + Modo prueba
        frame_bot = ctk.CTkFrame(self, fg_color="transparent")
        frame_bot.grid(row=5, column=0, pady=(14, 4))
        self.btn_ejecutar = ctk.CTkButton(frame_bot, text="Ejecutar Bot", fg_color="#27ae60",
                      hover_color="#1e8449", command=self.ejecutar_bot)
        self.btn_ejecutar.pack(side="left", padx=(8, 16))
        self.dry_run_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(frame_bot, text="Modo prueba (no guarda órdenes)",
                        variable=self.dry_run_var,
                        text_color="#e67e22",
                        fg_color="#e67e22", hover_color="#ca6f1e",
                        checkmark_color="white").pack(side="left")

        # ── Barra de progreso
        self.frame_progreso = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_progreso.grid(row=6, column=0, sticky="ew", padx=20, pady=(4, 2))
        self.frame_progreso.columnconfigure(0, weight=1)
        self.frame_progreso.columnconfigure(1, weight=0)
        self.progress_bar = ctk.CTkProgressBar(self.frame_progreso)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=0, column=0, sticky="ew")
        self.progress_label = ctk.CTkLabel(self.frame_progreso, text="", text_color="gray60")
        self.progress_label.grid(row=1, column=0)
        self.btn_detener = ctk.CTkButton(
            self.frame_progreso, text="Detener", width=80,
            fg_color="#c0392b", hover_color="#922b21",
            command=self._detener_bot)
        self.btn_detener.grid(row=0, column=1, rowspan=2, padx=(10, 0))
        self.frame_progreso.grid_remove()

        # ── Log
        frame_log_header = ctk.CTkFrame(self, fg_color="transparent")
        frame_log_header.grid(row=7, column=0, sticky="ew", padx=20, pady=(10,2))
        ctk.CTkLabel(frame_log_header, text="Log:", anchor="w").pack(side="left")
        ctk.CTkButton(frame_log_header, text="Limpiar", width=70, height=22,
                      fg_color="transparent", border_width=1, hover_color="#444",
                      font=ctk.CTkFont(size=11),
                      command=self._limpiar_log).pack(side="right")
        self.log = ctk.CTkTextbox(self, height=120, state="disabled")
        self.log.grid(row=8, column=0, sticky="nsew", padx=20, pady=(0,15))

    def _headers_tabla(self):
        for col in range(5):
            self.tabla_frame.columnconfigure(col, weight=1)
        self.tabla_frame.columnconfigure(5, weight=0, minsize=44)
        self.tabla_frame.columnconfigure(6, weight=0, minsize=44)

        headers = ["Beneficio", "Parentesco", "Diagnóstico", "Inicio", "Sesiones", "", ""]
        for col, h in enumerate(headers):
            ctk.CTkLabel(self.tabla_frame, text=h,
                         font=ctk.CTkFont(weight="bold")).grid(
                         row=0, column=col, padx=4, sticky="ew")

    def actualizar_tabla(self):
        guardar_pacientes_estado(self.pacientes)
        for w in self.tabla_frame.winfo_children():
            w.destroy()
        self._headers_tabla()
        for i, p in enumerate(self.pacientes):
            datos = [
                p["beneficio"], p["parentesco"], p["diagnostico"],
                p["fecha_inicio"].strftime("%d/%m/%Y"), str(p["sesiones"])
            ]
            for col, val in enumerate(datos):
                ctk.CTkLabel(self.tabla_frame, text=val).grid(
                    row=i+1, column=col, padx=4, pady=1, sticky="ew")
            ctk.CTkButton(self.tabla_frame, text="✎", width=40,
                          font=ctk.CTkFont(size=16),
                          command=lambda idx=i: self.abrir_editar_paciente(idx)).grid(
                          row=i+1, column=5, padx=2)
            ctk.CTkButton(self.tabla_frame, text="✕", width=40, fg_color="#c0392b",
                          command=lambda idx=i: self.quitar_paciente(idx)).grid(
                          row=i+1, column=6, padx=2)

    def quitar_paciente(self, idx):
        if not messagebox.askyesno("Confirmar", f"¿Eliminás al paciente con beneficio {self.pacientes[idx]['beneficio']}?"):
            return
        self.pacientes.pop(idx)
        self.actualizar_tabla()

    def abrir_editar_paciente(self, idx):
        PacienteDialog(self, idx)

    def limpiar_pacientes(self):
        if not self.pacientes:
            return
        if not messagebox.askyesno("Confirmar", "¿Eliminás todos los pacientes cargados?"):
            return
        self.pacientes.clear()
        self.actualizar_tabla()

    def abrir_credenciales(self):
        CredencialesDialog(self)

    def abrir_feriados(self):
        FeriadosDialog(self)

    def abrir_agregar_paciente(self):
        PacienteDialog(self)

    def generar_excel(self):
        if not self.pacientes:
            messagebox.showwarning("Aviso", "No hay pacientes cargados.")
            return

        feriados = cargar_feriados()
        hoy = datetime.today().date()
        rows = []
        por_beneficio = {}

        for p in self.pacientes:
            fechas = generar_fechas(p["fecha_inicio"], p["sesiones"], feriados)
            practicas = p.get("practicas", DEFAULT_PRACTICAS)
            futuras = [f for f in fechas if f > hoy]
            if futuras:
                por_beneficio[p["beneficio"]] = futuras
            for fecha in fechas:
                fila = {
                    "Beneficio":       p["beneficio"],
                    "Parentesco":      p["parentesco"],
                    "Fecha":           fecha.strftime("%d/%m/%Y"),
                    "Cod_Diagnostico": p["diagnostico"],
                }
                for j, cod in enumerate(practicas, 1):
                    fila[f"Cod_Practica{j}"] = cod
                rows.append(fila)

        if por_beneficio:
            max_fecha = max(f for fechas in por_beneficio.values() for f in fechas)
            FechasFuturasDialog(
                self, por_beneficio, max_fecha,
                on_confirm=lambda: self._escribir_excel(rows, por_beneficio, max_fecha),
            )
        else:
            self._escribir_excel(rows, por_beneficio, max_fecha=None)

    def _escribir_excel(self, rows, por_beneficio, max_fecha):
        try:
            df = pd.DataFrame(rows)
            df.to_excel(EXCEL_PATH, index=False)
        except PermissionError:
            messagebox.showerror("Error", "No se pudo escribir el archivo. ¿Está abierto en Excel?")
            return
        except OSError as e:
            messagebox.showerror("Error", f"No se pudo escribir el archivo: {e}")
            return

        n_futuras = sum(len(v) for v in por_beneficio.values())
        msg = f"Excel generado: {len(rows)} filas para {len(self.pacientes)} paciente(s).\n"
        if por_beneficio:
            msg += f"⚠ {n_futuras} sesión(es) futuras — procesables a partir del {max_fecha.strftime('%d/%m/%Y')}.\n"
        self.log_append(msg)
        messagebox.showinfo("Listo", "pacientes.xlsx generado correctamente.")

    def _update_progress(self, actual, total):
        self.progress_bar.set((actual - 1) / total)
        self.progress_label.configure(text=f"Fila {actual} / {total}", text_color="gray60")

    def ejecutar_bot(self):
        if not self.usuario or not self.clave:
            messagebox.showwarning("Aviso", "Configurá las credenciales primero.")
            return
        if not EXCEL_PATH.exists():
            messagebox.showwarning("Aviso", "Generá el Excel primero.")
            return

        ok, error = validar_excel(EXCEL_PATH)
        if not ok:
            messagebox.showerror(
                "Excel inválido",
                f"El archivo pacientes.xlsx tiene un problema:\n\n{error}\n\nRegeneralo desde la aplicación."
            )
            return

        if STOP_FLAG.exists():
            STOP_FLAG.unlink(missing_ok=True)

        if self._hide_after_id is not None:
            self.frame_progreso.after_cancel(self._hide_after_id)
            self._hide_after_id = None
        self.btn_ejecutar.configure(state="disabled")
        self.btn_detener.configure(state="normal", text="Detener")
        self.btn_detener.grid(row=0, column=1, rowspan=2, padx=(10, 0))
        self.frame_progreso.grid()
        self.progress_bar.set(0)
        self.progress_label.configure(text="", text_color="gray60")
        modo = " [MODO PRUEBA]" if self.dry_run_var.get() else ""
        self.log_append(f"Iniciando bot{modo}...\n")
        env = os.environ.copy()
        env["PAMI_USER"]        = self.usuario
        env["PAMI_PASS"]        = self.clave
        env["PAMI_DRY_RUN"]     = "1" if self.dry_run_var.get() else ""
        env["PYTHONUNBUFFERED"] = "1"

        def correr():
            resumen = {"ok": 0, "omit": 0, "det": 0, "err": 0, "reporte": None, "detenido": False, "reporte_error": False}

            proc = subprocess.Popen(
                [sys.executable, Path(__file__).parent / "bot.py"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                env=env,
                cwd=str(DATA_DIR),
            )
            self._proc = proc
            for linea in proc.stdout:
                self.after(0, lambda l=linea: self.log_append(l))
                m = re.search(r"Fila (\d+) de (\d+)", linea)
                if m:
                    actual, total = int(m.group(1)), int(m.group(2))
                    self.after(0, lambda a=actual, t=total: self._update_progress(a, t))
                m_res = re.search(r"Total: \d+ \| OK: (\d+) \| Omitidos: (\d+) \| Detenidos: (\d+) \| Errores: (\d+)", linea)
                if m_res:
                    resumen["ok"]   = int(m_res.group(1))
                    resumen["omit"] = int(m_res.group(2))
                    resumen["det"]  = int(m_res.group(3))
                    resumen["err"]  = int(m_res.group(4))
                m_rep = re.search(r"Reporte guardado en (.+\.xlsx)", linea)
                if m_rep:
                    resumen["reporte"] = m_rep.group(1).strip()
                if "[DETENIDO]" in linea:
                    resumen["detenido"] = True
                if "[ERROR REPORTE]" in linea:
                    resumen["reporte_error"] = True
            proc.wait()
            self._proc = None

            def _mostrar_resultado():
                self.progress_bar.set(1.0)
                self.btn_detener.grid_remove()
                if self._hide_after_id is not None:
                    self.frame_progreso.after_cancel(self._hide_after_id)
                self._hide_after_id = self.frame_progreso.after(1200, self._hide_progreso)
                if resumen["detenido"]:
                    partes_det = [f"{resumen['det']} detenida(s)"]
                    if resumen["ok"]:
                        partes_det.insert(0, f"{resumen['ok']} OK")
                    if resumen["err"]:
                        partes_det.append(f"{resumen['err']} error(es)")
                    self.progress_label.configure(
                        text=f"Detenido — {', '.join(partes_det)}", text_color="#f39c12"
                    )
                    if resumen["reporte"]:
                        messagebox.showwarning(
                            "Bot detenido",
                            f"Se procesaron {resumen['ok']} orden(es) correctamente.\n"
                            f"{resumen['det']} fila(s) marcada(s) como DETENIDO.\n"
                            + (f"{resumen['err']} error(es) encontrado(s).\n" if resumen["err"] else "")
                            + f"\nRevisá el reporte en:\n{resumen['reporte']}"
                        )
                elif resumen["err"] == 0 and not resumen["reporte_error"]:
                    sufijo = f" ({resumen['omit']} omitido(s))" if resumen["omit"] else ""
                    self.progress_label.configure(text=f"Completado{sufijo}", text_color="#27ae60")
                elif resumen["reporte_error"] and resumen["err"] == 0:
                    sufijo = f" ({resumen['omit']} omitido(s))" if resumen["omit"] else ""
                    self.progress_label.configure(text=f"Completado{sufijo} — reporte no guardado", text_color="#e67e22")
                    messagebox.showwarning(
                        "Reporte no guardado",
                        f"Las {resumen['ok']} orden(es) se procesaron correctamente,\n"
                        "pero no se pudo guardar el reporte Excel.\n\n"
                        "Revisá el log para más detalles (¿archivo abierto en Excel?)."
                    )
                else:
                    parte_omit = f", {resumen['omit']} omitido(s)" if resumen["omit"] else ""
                    self.progress_label.configure(
                        text=f"Finalizado con {resumen['err']} error(es){parte_omit}",
                        text_color="#e67e22"
                    )
                    if resumen["reporte"]:
                        messagebox.showwarning(
                            "Bot finalizado con errores",
                            f"Se procesaron {resumen['ok']} orden(es) correctamente.\n"
                            f"{resumen['err']} error(es) encontrado(s).\n\n"
                            f"Revisá el reporte en:\n{resumen['reporte']}"
                        )
                self.log_append("Bot finalizado.\n")
                self.btn_ejecutar.configure(state="normal")

            self.after(0, _mostrar_resultado)

        threading.Thread(target=correr, daemon=True).start()

    def _hide_progreso(self):
        self._hide_after_id = None
        self.frame_progreso.grid_remove()

    def _detener_bot(self):
        if not STOP_FLAG.exists():
            # Primer clic: parada cooperativa — el bot termina el paso actual y genera el reporte
            STOP_FLAG.write_text("")
            self.btn_detener.configure(text="Forzar cierre")
        else:
            # Segundo clic: bot colgado — matar el proceso (sin reporte)
            self.btn_detener.configure(state="disabled", text="Cerrando...")
            if self._proc is not None:
                self._proc.terminate()

    def _limpiar_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def log_append(self, texto):
        self.log.configure(state="normal")
        self.log.insert("end", texto)
        self.log.see("end")
        self.log.configure(state="disabled")

if __name__ == "__main__":
    App().mainloop()
