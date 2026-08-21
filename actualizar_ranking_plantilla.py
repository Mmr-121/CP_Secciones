"""
actualizar_ranking_plantilla.py

Vulca el informe generado en Ranking_CP conservando el mapeo completo de columnas,
incluyendo Evolución Bruta y Evolución Neta.
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
COL_IDX = {nombre: i + 1 for i, nombre in enumerate(COLUMNAS)}


def main():
    if not INFORME.exists():
        raise FileNotFoundError(f"No se encontró {INFORME.name}.")
    if not PLANTILLA.exists():
        raise FileNotFoundError(f"No se encontró {PLANTILLA.name}.")

    df = pd.read_excel(INFORME, dtype={"Codigo_Postal": str, "Codigo_INE_Principal": str})
    
    wb = load_workbook(PLANTILLA)
    if HOJA not in wb.sheetnames:
        raise ValueError(f"La plantilla no contiene la hoja '{HOJA}'.")
    ws = wb[HOJA]

    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    for i, fila in df.iterrows():
        r = i + 2
        cobertura_baja = fila.get("Pct_Cobertura_Renta", 100) < 50
        observacion = (
            f"Cobertura {fila.get('Pct_Cobertura_Renta', 0):.0f}% "
            f"(secreto estadístico parcial) — verificar antes de segmentar."
            if cobertura_baja else ""
        )

        ws.cell(r, COL_IDX["Ranking"], i + 1)
        ws.cell(r, COL_IDX["Código Postal"], str(fila["Codigo_Postal"]).zfill(5))
        ws.cell(r, COL_IDX["Municipio"], fila["Municipio"])
        ws.cell(r, COL_IDX["Codigo INE"], fila.get("Codigo_INE_Principal", ""))
        
        # Renta por hogar y evoluciones
        ws.cell(r, COL_IDX["Renta Bruta media (€)"], fila.get("Renta_Bruta_Hogar_Media"))
        ws.cell(r, COL_IDX["Evolución bruta"], fila.get("Evolución_Bruta_%"))
        ws.cell(r, COL_IDX["Renta Neta media (€)"], fila.get("Renta_Neta_Hogar_Media"))
        ws.cell(r, COL_IDX["Evolucion Neta"], fila.get("Evolución_Neta_%"))
        
        ws.cell(r, COL_IDX["Urbanizaciones"], fila.get("Pct_Urbanizaciones"))
        ws.cell(r, COL_IDX["Observaciones"], observacion)

    wb.save(PLANTILLA)
    print(f"Plantilla '{PLANTILLA.name}' actualizada correctamente con {len(df)} CP.")


if __name__ == "__main__":
    main()