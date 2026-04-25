# Gestión de Licencias — KINETICA

Cada cliente necesita una clave de licencia para usar la app. Este documento explica
cómo generarla, registrarla, renovarla y revocarla.

---

## Cómo funciona el sistema

1. El cliente instala la app e ingresa su clave la primera vez.
2. La app realiza un POST al servidor (Google Apps Script) con la clave y un ID de máquina.
3. El servidor verifica la clave en el Google Sheet y, si es válida, registra el equipo.
4. La app guarda una caché local cifrada para no consultar el servidor en cada apertura.
5. La caché se re-valida online cada 1 día. Sin internet, hay un período de gracia de 7 días.

---

## Paso 1 — Generar la clave

Desde la raíz del proyecto, con el entorno virtual activado:

```bash
# Activar venv (Windows)
venv\Scripts\activate

# Generar 1 clave
python tools/generar_key.py

# Ejemplo de salida:
# KINE-R7TM-4PXQ-9HBZ
```

```bash
# Generar varias claves a la vez
python tools/generar_key.py -n 5

# Ejemplo de salida:
# KINE-R7TM-4PXQ-9HBZ
# KINE-2VCN-EFWQ-7KJP
# KINE-9YBD-MRXZ-4TUG
# KINE-5HKP-QZ3N-8WCV
# KINE-J6NT-2DBR-VMYX
```

```bash
# Generar en formato CSV listo para pegar en el Sheet (validez 1 año por defecto)
python tools/generar_key.py --csv

# Ejemplo de salida:
# key,cliente,expiracion
# KINE-R7TM-4PXQ-9HBZ,,2027-04-25

# Con validez personalizada (ej: 6 meses = 180 días)
python tools/generar_key.py --csv --dias 180
```

---

## Paso 2 — Registrar la clave en Google Sheets

Abrir el Google Sheet del servidor de licencias y navegar a la hoja **`licencias`**.

La estructura de la hoja es:

| key | cliente | expiracion | machine_id | activado_en | notas |
|-----|---------|------------|------------|-------------|-------|
| KINE-R7TM-4PXQ-9HBZ | Juan Pérez | 2027-04-25 | | | Kinesiólogo, Córdoba |

**Columnas a completar al crear la fila:**

| Columna | Descripción |
|---------|-------------|
| `key` | La clave generada en el Paso 1 |
| `cliente` | Nombre del cliente (referencia interna) |
| `expiracion` | Fecha de vencimiento en formato `YYYY-MM-DD` (ej: `2027-04-25`) |
| `notas` | Información adicional: ciudad, teléfono, etc. |

**Columnas que se completan automáticamente:**

| Columna | Cuándo |
|---------|--------|
| `machine_id` | La primera vez que el cliente activa la app en su equipo |
| `activado_en` | La primera vez que el cliente activa la app en su equipo |

> **Importante:** dejar `machine_id` y `activado_en` en blanco al crear la fila.
> Si se completan manualmente, la activación del cliente fallará.

---

## Paso 3 — Enviar la clave al cliente

Enviar la clave al cliente por el medio preferido (WhatsApp, correo, etc.):

```
Su clave de licencia KINETICA es:

KINE-R7TM-4PXQ-9HBZ

Ingresarla la primera vez que se abra la aplicación.
```

---

## Renovar una licencia vencida

Cuando vence la licencia, el cliente ve un mensaje en la app indicando que debe
contactar al administrador. Para renovarla:

1. Abrir el Google Sheet → hoja `licencias`.
2. Localizar la fila del cliente correspondiente.
3. Modificar la columna `expiracion` con la nueva fecha:

```
Antes:  2027-04-25
Después: 2028-04-25
```

La app detecta el cambio en la próxima apertura. No se requiere ninguna acción adicional.

---

## Cliente que cambió de computadora

La licencia está vinculada al equipo donde fue activada. Si el cliente formateó o
cambió de PC, realizar los siguientes pasos:

1. Abrir el Google Sheet → hoja `licencias`.
2. Localizar la fila del cliente.
3. Borrar el contenido de las celdas `machine_id` y `activado_en` (dejarlas vacías).

La próxima vez que el cliente abra la app en el nuevo equipo, el sistema registrará
la nueva máquina automáticamente.

---

## Revocar una licencia

Para bloquear el acceso de un cliente de forma inmediata:

1. Abrir el Google Sheet → hoja `licencias`.
2. Localizar la fila del cliente.
3. Cambiar la columna `expiracion` a una fecha pasada:

```
2000-01-01
```

La app detecta el cambio en la próxima apertura (o al vencer la caché de 1 día) y
muestra el mensaje de licencia vencida.

---

## Resumen rápido

| Situación | Acción |
|-----------|--------|
| Cliente nuevo | Generar clave → agregar fila en Sheet → enviar clave al cliente |
| Renovar licencia | Modificar `expiracion` en Sheet |
| Cliente cambió de PC | Borrar `machine_id` y `activado_en` en Sheet |
| Revocar acceso | Cambiar `expiracion` a `2000-01-01` |
