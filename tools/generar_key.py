#!/usr/bin/env python3
"""
Generador de claves de licencia para KINETICA.

Uso:
    python tools/generar_key.py           # genera 1 clave
    python tools/generar_key.py -n 10     # genera 10 claves
    python tools/generar_key.py -n 5 --csv # genera 5 claves en formato CSV (para pegar en el Sheet)

Formato de clave: KINE-XXXX-XXXX-XXXX
Alfabeto: excluye caracteres ambiguos (0, O, 1, I, L) para facilitar la transcripción.
"""

import argparse
import random
import secrets
import sys
from datetime import date, timedelta

ALPHABET = "ABCDEFGHJKMNPQRSTUVWXYZ23456789"


def generar_key() -> str:
    grupos = ["".join(secrets.choice(ALPHABET) for _ in range(4)) for _ in range(3)]
    return "KINE-" + "-".join(grupos)


def main():
    parser = argparse.ArgumentParser(
        description="Generador de claves de licencia KINETICA",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "-n", "--cantidad",
        type=int, default=1,
        help="Cantidad de claves a generar (default: 1)",
    )
    parser.add_argument(
        "--csv",
        action="store_true",
        help=(
            "Salida en formato CSV listo para pegar en el Google Sheet.\n"
            "Incluye columnas: key,cliente,expiracion (fecha 1 año desde hoy)"
        ),
    )
    parser.add_argument(
        "--dias",
        type=int, default=365,
        help="Días de validez para las claves generadas en modo --csv (default: 365)",
    )
    args = parser.parse_args()

    if args.cantidad < 1:
        parser.error("La cantidad debe ser al menos 1.")

    if args.csv:
        expiracion = (date.today() + timedelta(days=args.dias)).strftime("%Y-%m-%d")
        print("key,cliente,expiracion")
        for _ in range(args.cantidad):
            print(f"{generar_key()},,{expiracion}")
    else:
        for _ in range(args.cantidad):
            print(generar_key())


if __name__ == "__main__":
    main()
