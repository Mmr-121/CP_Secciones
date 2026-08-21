"""
Vuelca el Informe_Renta_Urbanizacion_por_CP.xlsx (salida de Renta_segun_CP.py)
en la hoja Ranking_CP de Plantilla_Profesional_IPJ_Alicante.xlsx, para que la
plantilla de campaña quede al día tras cada ejecución del pipeline.

Columnas rellenadas por este script:
    Código Postal, Municipio, Codigo INE, Renta Bruta media, Renta Neta media,
    Urbanizaciones (%), Observaciones (aviso de fiabilidad).

Columnas que NO toca: Evolución, Habitantes, % Unifamiliares, Competencia, IPJ, Prioridad SEO/Ads.
"""

from pathlib import Path
import pandas as pd
from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent

INFORME = BASE_DIR / "Informe_Renta_Urbanizacion_por_CP.xlsx"
PLANTILLA = BASE_DIR / "Plantilla_Profesional_IPJ_Alicante.xlsx"
HOJA = "Ranking_CP"

COLUMNAS = [
    "Ranking", "Código Postal", "Municipio", "Codigo INE",
    "Renta Bruta media (€)", "Evolución bruta", "Renta Neta media (€)",
    "Evolucion Neta", "Habitantes", "% Unifamiliares", "Urbanizaciones",
    "Competencia", "IPJ", "Prioridad SEO", "Prioridad Ads", "Observaciones",
]
COL_IDX = {nombre: i + 1 for i, nombre in enumerate(COLUMNAS)}  # 1-based


def main():
    if not INFORME.exists():
        raise FileNotFoundError(
            f"No se encontró {INFORME.name}. Ejecuta antes Renta_segun_CP.py."
        )
    if not PLANTILLA.exists():
        raise FileNotFoundError(f"No se encontró {PLANTILLA.name} en {BASE_DIR}.")

    print("1. Leyendo informe de renta/urbanizaciones por CP...")
    df = pd.read_excel(INFORME, dtype={"Codigo_Postal": str, "Codigo_INE_Principal": str})
    df = df.sort_values(
        ["Renta_Bruta_Media", "Nº_Tramos_Urbanizacion"], ascending=[False, False]
    ).reset_index(drop=True)

    print(f"2. Abriendo '{PLANTILLA.name}'...")
    wb = load_workbook(PLANTILLA)
    if HOJA not in wb.sheetnames:
        raise ValueError(f"La plantilla no tiene una hoja llamada '{HOJA}'.")
    ws = wb[HOJA]

    print("3. Limpiando filas de datos anteriores (se conserva la cabecera)...")
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    print(f"4. Escribiendo {len(df)} códigos postales...")
    for i, fila in df.iterrows():
        r = i + 2  # fila 1 = cabecera
        cobertura_baja = fila.get("Pct_Cobertura_Renta", 100) < 50
        observacion = (
            f"Cobertura de renta {fila.get('Pct_Cobertura_Renta', 0):.0f}% "
            f"(secreto estadístico en parte de las secciones) — dato poco fiable, revisar antes de segmentar."
            if cobertura_baja else ""
        )

        ws.cell(r, COL_IDX["Ranking"], i + 1)
        ws.cell(r, COL_IDX["Código Postal"], str(fila["Codigo_Postal"]).zfill(5))
        ws.cell(r, COL_IDX["Municipio"], fila["Municipio"])
        ws.cell(r, COL_IDX["Codigo INE"], fila.get("Codigo_INE_Principal", ""))
        ws.cell(r, COL_IDX["Renta Bruta media (€)"], fila["Renta_Bruta_Media"])
        ws.cell(r, COL_IDX["Renta Neta media (€)"], fila["Renta_Neta_Media"])
        ws.cell(r, COL_IDX["Urbanizaciones"], fila["Pct_Urbanizaciones"])
        ws.cell(r, COL_IDX["Observaciones"], observacion)

    print(f"5. Guardando '{PLANTILLA.name}'...")
    wb.save(PLANTILLA)

    print("\nCompletado.")
    print(f"  CP volcados en Ranking_CP: {len(df)}")


if __name__ == "__main__":
    main()