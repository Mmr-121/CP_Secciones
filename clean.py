"""
clean.py
Borra los ficheros de salida generados por ejecuciones anteriores del
pipeline, para no arrastrar datos obsoletos de una edición del INE a otra.
"""

from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

FICHEROS_A_BORRAR = [
    "Relacion_Secciones_CP.xlsx",
    "Informe_Renta_Urbanizacion_por_CP.xlsx",
]


def main():
    for nombre in FICHEROS_A_BORRAR:
        f = BASE_DIR / nombre
        if f.exists():
            f.unlink()
            print(f"Borrado: {nombre}")
        else:
            print(f"(no existía, se omite) {nombre}")


if __name__ == "__main__":
    main()