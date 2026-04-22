# pami-bot

Bot de automatización para la carga de órdenes ambulatorias en el sistema web de PAMI (efectoresweb.pami.org.ar).

Incluye una GUI de escritorio para gestionar pacientes, generar el Excel de trabajo y ejecutar el bot con seguimiento en tiempo real.

---

## Características

- Carga automática de órdenes ambulatorias en PAMI
- GUI con lista de pacientes, credenciales y log en vivo
- Generación de fechas de sesiones excluyendo fines de semana y feriados
- Soporte para múltiples prácticas por orden
- Modo prueba (DRY_RUN): navega el formulario pero no guarda
- Reporte Excel con colores al finalizar (verde = OK, rojo = error)
- Credenciales guardadas de forma segura con `keyring`

---

## Requisitos

- Python 3.10+
- Playwright (con Chromium)

```bash
pip install -r requirements.txt
playwright install chromium
```

---

## Uso

### GUI

```bash
python gui.py
```

1. Abrí **Credenciales** e ingresá usuario y contraseña de PAMI
2. Opcionalmente configurá **Días a excluir** (feriados u otros días)
3. Agregá pacientes con **+ Agregar Paciente**
4. Hacé clic en **Generar Excel** y luego en **Ejecutar Bot**


---

## Notas

- El sistema PAMI no permite cargar órdenes con fechas futuras
- El bot detecta y reporta sesiones duplicadas sin interrumpir el proceso
- `pacientes.xlsx` está en `.gitignore` por contener datos de pacientes
