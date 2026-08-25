"""
actualizar_todo.py

Orquestador del pipeline de renta y urbanizaciones por Código Postal (Alicante).

Ejecuta en orden:
    1. clean.py                          -> borra los ficheros de salida anteriores
    2. generar_relacion_secciones_cp.py  -> tabla de traslación sección censal -> CP
    3. Renta_segun_CP.py                 -> informe agregado de renta y % urbanizaciones por CP


Antes de ejecutar, sobrescribe la carpeta caj_esp_072026 con los ficheros
VIAS/TRAM más recientes del INE, y coloca el Excel de renta (30833.xlsx o el
que corresponda) en esta misma carpeta.
"""

import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

PIPELINE = [
    "clean.py",
    "generar_relacion_secciones_cp.py",
    "Renta_segun_CP.py",
    "actualizar_ranking_plantilla.py",
]


def ejecutar(script_name: str) -> None:
    script_path = BASE_DIR / script_name
    if not script_path.exists():
        print(f"[AVISO] No se encontró {script_name}, se omite.")
        return

    print(f"\n{'=' * 60}")
    print(f"Ejecutando: {script_name}")
    print(f"{'=' * 60}")

    resultado = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=BASE_DIR,
    )

    if resultado.returncode != 0:
        print(f"\nERROR: {script_name} terminó con código {resultado.returncode}.")
        print("Se detiene el pipeline para que puedas revisar el problema.")
        sys.exit(resultado.returncode)


def main():
    print("Iniciando actualización completa del pipeline de renta y urbanizaciones...")
    for script in PIPELINE:
        ejecutar(script)

    print(f"\n{'=' * 60}")
    print("Pipeline completado correctamente.")
    print("Ficheros generados:")
    print("  - Relacion_Secciones_CP.xlsx")
    print("  - Informe_Renta_Urbanizacion_por_CP.xlsx")
    print("  - Plantilla_Profesional_IPJ_Alicante.xlsx actualizada")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()