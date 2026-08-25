"""
Genera un Excel con la relación Sección censal (CUSEC) <-> Código Postal:
    nombre_municipio | codigo_ine_municipio | codigo_seccion_cusec | seccion_completa | codigo_postal

Cada sección puede tener tramos en más de un CP; aquí se asigna el CP
"mayoritario" (el que más tramos tiene dentro de esa sección).
"""

from pathlib import Path
from collections import defaultdict
import glob
import re
import warnings

import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# RUTAS DE LOS ARCHIVOS
BASE_DIR = Path(__file__).resolve().parent.parent / "datos/entrada"
BASE_DIR_SALIDA = Path(__file__).resolve().parent.parent / "datos/salida"


def _localizar_carpeta_datos() -> Path:
    """Busca la carpeta de la edición del INE descargada (caj_esp_XXXXXX/caj_esp_XXXXXX)."""
    candidatas = sorted(BASE_DIR.glob("caj_esp_*"))
    candidatas = [c for c in candidatas if c.is_dir()]
    if not candidatas:
        raise FileNotFoundError(
            f"No se encontró ninguna carpeta 'caj_esp_*' con los datos del INE dentro de {BASE_DIR}. "
            "Descomprime aquí el zip de la última edición (VIAS/TRAM) antes de ejecutar."
        )
    carpeta = candidatas[-1]
    anidada = carpeta / carpeta.name
    return anidada if anidada.is_dir() else carpeta


FOLDER_DATOS = _localizar_carpeta_datos()

SALIDA = BASE_DIR_SALIDA / "Relacion_Secciones_CP.xlsx"

PROVINCIA = "03"  # Alicante


def encontrar_fichero_tram() -> Path:
    candidatos = sorted(glob.glob(str(FOLDER_DATOS / "TRAM*")))
    candidatos = [c for c in candidatos if not c.endswith(".xlsx")]
    if not candidatos:
        raise FileNotFoundError(
            f"No se encontró ningún fichero TRAM en {FOLDER_DATOS}"
        )
    return Path(candidatos[0])


def encontrar_excel_renta() -> Path:
    candidatos = sorted(glob.glob(str(BASE_DIR / "*.xlsx")))
    candidatos = [
        c for c in candidatos
        if "Informe_Renta" not in c and "Relacion_Secciones" not in c and "Plantilla" not in c
    ]
    if not candidatos:
        raise FileNotFoundError(f"No se encontró el Excel de renta del INE en {BASE_DIR}")
    if len(candidatos) > 1:
        raise ValueError(f"Hay más de un Excel candidato a renta del INE en {BASE_DIR}: {candidatos}")
    return Path(candidatos[0])


def main():
    fichero_tram = encontrar_fichero_tram()
    archivo_nombres = encontrar_excel_renta()

    # TRAM: nº de tramos por (CUSEC, CP), para quedarnos con el CP mayoritario de cada sección.
    print(f"1. Leyendo TRAM ({fichero_tram.name}, puede tardar unos minutos, es un fichero grande)...")
    cusec_cp_tramos = defaultdict(lambda: defaultdict(int))

    with open(fichero_tram, "r", encoding="latin1") as f:
        for linea in f:
            if len(linea) < 47 or linea[0:2] != PROVINCIA:
                continue
            cusec = linea[0:10].strip()
            cpos = linea[42:47].strip()
            if cpos.isdigit() and len(cpos) == 5:
                cusec_cp_tramos[cusec][cpos] += 1

    cp_mayoritario_por_cusec = {
        cusec: max(cp_counts.items(), key=lambda kv: kv[1])[0]
        for cusec, cp_counts in cusec_cp_tramos.items()
    }

    print(f"2. Leyendo nombres de municipios y secciones ({archivo_nombres.name})...")
    df_bruto = pd.read_excel(archivo_nombres, header=None)

    fila_cabecera = -1
    for i, row in df_bruto.iterrows():
        if row.astype(str).str.contains("renta neta", case=False, na=False).any():
            fila_cabecera = i
            break
    if fila_cabecera == -1:
        raise ValueError("No se encontró la fila de cabecera en el Excel del INE.")

    df_nombres = df_bruto.iloc[fila_cabecera + 2:, [0]].copy()
    df_nombres.columns = ["Territorio"]
    df_nombres["CUSEC"] = df_nombres["Territorio"].astype(str).str.extract(r"^(\d{10})")
    df_nombres = df_nombres.dropna(subset=["CUSEC"]).copy()

    def extraer_campos(texto):
        cod_mun = texto[:5]
        resto = texto[10:].strip()
        match_secc = re.search(r"\s+sección\s+.*$", resto, re.IGNORECASE)
        nombre_mun = resto[: match_secc.start()].strip() if match_secc else resto
        return pd.Series([cod_mun, nombre_mun])

    df_nombres[["codigo_ine_municipio", "nombre_municipio"]] = df_nombres["Territorio"].apply(extraer_campos)
    nombre_municipio_por_codigo = dict(
        df_nombres.drop_duplicates("codigo_ine_municipio")[["codigo_ine_municipio", "nombre_municipio"]].values
    )

    # Cruce y exportación
    print("3. Cruzando secciones con su CP mayoritario...")
    df_secc = pd.DataFrame({"CUSEC": list(cp_mayoritario_por_cusec.keys())})
    df_secc["Codigo_Postal"] = df_secc["CUSEC"].map(cp_mayoritario_por_cusec)
    df_secc["codigo_ine_municipio"] = df_secc["CUSEC"].str[:5]
    df_secc["nombre_municipio"] = df_secc["codigo_ine_municipio"].map(nombre_municipio_por_codigo).fillna("?")

    df_secc["codigo_distrito"] = df_secc["CUSEC"].str.zfill(10).str[5:7]
    df_secc["codigo_seccion"] = df_secc["CUSEC"].str.zfill(10).str[7:10]
    df_secc["Territorio"] = (
        df_secc["CUSEC"] + " " + df_secc["nombre_municipio"] + " sección " + df_secc["codigo_seccion"]
    )

    df_exportar = df_secc[[
        "nombre_municipio",
        "codigo_ine_municipio",
        "codigo_distrito",
        "codigo_seccion",
        "CUSEC",
        "Territorio",
        "Codigo_Postal",
    ]].rename(columns={
        "CUSEC": "codigo_seccion_cusec",
        "Territorio": "seccion_completa",
        "Codigo_Postal": "codigo_postal",
    })
    df_exportar["codigo_seccion_cusec"] = df_exportar["codigo_seccion_cusec"].astype(str).str.zfill(10)
    df_exportar["codigo_postal"] = df_exportar["codigo_postal"].astype(str).str.zfill(5)

    print(f"4. Guardando '{SALIDA.name}'...")
    df_exportar.to_excel(SALIDA, index=False, sheet_name="Relacion_Secciones_CP")

    print("\nCompletado.")
    print(f"  Secciones exportadas : {len(df_exportar)}")
    print(f"  Fichero de salida    : {SALIDA}")


if __name__ == "__main__":
    main()