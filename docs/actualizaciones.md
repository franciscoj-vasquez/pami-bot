# Publicar una Actualización — KINETICA

Este documento describe el proceso completo para compilar una nueva versión de KINETICA
y distribuirla a todos los clientes que tienen la app instalada.

---

## Cómo funciona el sistema

1. Al abrir la app, se realiza una consulta al servidor (Google Apps Script).
2. El servidor devuelve la versión más reciente y la URL de descarga del instalador.
3. Si la versión del servidor es mayor a la instalada en el equipo, aparece un banner
   verde en la parte superior de la app: **"Nueva versión disponible"**.
4. El cliente hace clic en **Actualizar**:
   - Se descarga el nuevo instalador desde GitHub.
   - Se lanza el instalador en modo silencioso (sin ventanas ni confirmaciones).
   - La app espera a que el instalador termine y luego se cierra sola.
5. El cliente vuelve a abrir la app y ya tiene la versión nueva.
   La licencia, las credenciales y los datos de pacientes se conservan intactos.

---

## Paso 1 — Modificar el código

Realizar los cambios necesarios en el código fuente (`src/bot.py`, `src/gui.py`, etc.).

---

## Paso 2 — Actualizar el número de versión

Abrir `installer/setup.iss` y modificar la línea `MyAppVersion`:

```ini
; Antes:
#define MyAppVersion   "1.0.0"

; Después (ejemplo):
#define MyAppVersion   "1.1.0"
```

Seguir el esquema **MAJOR.MINOR.PATCH**:

| Tipo de cambio | Qué incrementar | Ejemplo |
|----------------|-----------------|---------|
| Corrección de un selector o interacción con PAMI | PATCH | `1.0.0` → `1.0.1` |
| Nueva funcionalidad menor | MINOR | `1.0.0` → `1.1.0` |
| Cambio grande o incompatible | MAJOR | `1.0.0` → `2.0.0` |

---

## Paso 3 — Compilar

Desde la raíz del proyecto, ejecutar el script de build:

```
installer\build.bat
```

El script realiza dos pasos automáticamente:

1. **PyInstaller** — empaqueta `gui.py` y `bot.py` en ejecutables `.exe`.
2. **Inno Setup** — genera el instalador final `dist\KINETICA_setup.exe`.

Al finalizar se debe ver:

```
[1/2] OK - Ejecutables generados en dist\KINETICA\
[2/2] OK - Instalador generado en: dist\KINETICA_setup.exe
```

> Si Inno Setup no está instalado, descargarlo desde https://jrsoftware.org/isdl.php

---

## Paso 4 — Publicar el release en GitHub

1. Ir al repositorio en GitHub: `github.com/franciscoj-vasquez/pami-bot`
2. En el panel lateral derecho, hacer clic en **Releases** → **Draft a new release**.
3. En **Choose a tag**, escribir el tag de la nueva versión (ej: `v1.1.0`) y seleccionar
   **"Create new tag: v1.1.0 on publish"**.
4. En **Release title**, escribir `v1.1.0`.
5. En la sección de archivos adjuntos, subir el archivo `dist\KINETICA_setup.exe`.
   El nombre del archivo debe ser exactamente **`KINETICA_setup.exe`**.
6. Hacer clic en **Publish release**.

Para verificar que el archivo es accesible, pegar la siguiente URL en el navegador
y confirmar que se inicia la descarga:

```
https://github.com/franciscoj-vasquez/pami-bot/releases/latest/download/KINETICA_setup.exe
```

---

## Paso 5 — Actualizar el Google Sheet

1. Abrir el Google Sheet del servidor de licencias.
2. Navegar a la hoja **`config`**.
3. Modificar el valor de la fila `version` con el nuevo número:

```
Antes:  1.0.0
Después: 1.1.0
```

La columna `download_url` **no requiere cambios** — siempre apunta a la última
release de GitHub (`/releases/latest/`).

**A partir de este momento**, todos los clientes que abran la app verán el banner
de actualización.

---

## Ejemplo completo — corrección de un selector PAMI

Escenario: PAMI cambió el ID de un botón y el bot deja de funcionar.
Se corrige en `src/bot.py` y se distribuye la actualización.

```
1. Editar src/bot.py con la corrección.

2. En installer/setup.iss:
   #define MyAppVersion   "1.0.1"

3. Ejecutar installer\build.bat
   → genera dist\KINETICA_setup.exe

4. En GitHub:
   → New release → tag: v1.0.1
   → Subir KINETICA_setup.exe
   → Publish release

5. En Google Sheet → hoja config:
   → version: 1.0.1

6. Los clientes abren la app → ven el banner → hacen clic en Actualizar
   → la app se actualiza sola y se reinicia.
```

---

## Resumen del flujo

```
Cambio en el código
       ↓
Actualizar versión en setup.iss
       ↓
Ejecutar installer\build.bat
       ↓
Publicar release en GitHub con KINETICA_setup.exe
       ↓
Actualizar versión en Google Sheet (hoja "config")
       ↓
Los clientes reciben la actualización automáticamente
```

---

## Consideraciones

- **La URL de descarga no cambia nunca.** GitHub redirige `/releases/latest/` siempre
  al release más reciente. El nombre del archivo debe ser siempre `KINETICA_setup.exe`.
- **Los datos del cliente no se borran.** La licencia, las credenciales y los pacientes
  se almacenan en `%LOCALAPPDATA%\KINETICA\data\`, que el instalador no toca.
- **El cliente no necesita reinstalar manualmente.** El proceso es completamente
  automático desde el clic en "Actualizar".
- **Si un cliente no abre la app por varios días**, verá el banner igualmente la próxima
  vez que la abra, sin importar cuánto tiempo haya pasado.
