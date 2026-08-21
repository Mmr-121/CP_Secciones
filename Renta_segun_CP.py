"""
Renta_segun_CP.py / generar_informe_renta_cp.py

Genera un Excel con, para cada Código Postal de la provincia de Alicante:
  - Renta neta media por HOGAR (actual y evolución %)
  - Renta bruta media por HOGAR (actual y evolución %)
  - % de tramos de tipo "urbanización" (URB/URBAT)
  - % de cobertura de renta y fiabilidad
"""

from pathlib import Path
from collections import defaultdict
import glob
import re
import warnings

import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

BASE_DIR = Path(__file__).resolve().parent


def _localizar_carpeta_datos() -> Path:
    candidatas = sorted(BASE_DIR.glob("caj_esp_*"))
    candidatas = [c for c in candidatas if c.is_dir()]
    if not candidatas:
        raise FileNotFoundError(
            f"No se encontró ninguna carpeta 'caj_esp_*' en {BASE_DIR}."
        )
    carpeta = candidatas[-1]
    anidada = carpeta / carpeta.name
    return anidada if anidada.is_dir() else carpeta


FOLDER_DATOS = _localizar_carpeta_datos()
PROVINCIA = "03"  # Alicante
TIPOS_VIA_URBANIZACION = {"URB", "URBAT"}

SALIDA = BASE_DIR / "Informe_Renta_Urbanizacion_por_CP.xlsx"


def encontrar_fichero(patron: str) -> Path:
    candidatos = sorted(glob.glob(str(FOLDER_DATOS / f"{patron}*")))
    candidatos = [c for c in candidatos if not c.endswith(".xlsx")]
    if not candidatos:
        raise FileNotFoundError(f"No se encontró fichero '{patron}' en {FOLDER_DATOS}")
    return Path(candidatos[0])


def encontrar_excel_renta() -> Path:
    candidatos = sorted(glob.glob(str(BASE_DIR / "*.xlsx")))
    candidatos = [
        c for c in candidatos
        if "Informe_Renta" not in c and "Relacion_Secciones" not in c
    ]
    if not candidatos:
        raise FileNotFoundError(f"No se encontró el Excel de renta del INE en {BASE_DIR}")
    return Path(candidatos[0])


def main():
    fichero_vias = encontrar_fichero("VIAS")
    fichero_tram = encontrar_fichero("TRAM")
    fichero_renta = encontrar_excel_renta()

    print("1. Leyendo VIAS para clasificar tipos de vía...")
    tipo_via_dict = {}
    with open(fichero_vias, "r", encoding="latin1") as f:
        for linea in f:
            if linea[0:2] != PROVINCIA:
                continue
            mun = linea[0:5]
            vcode = linea[5:10]
            idx = linea.find("20260630")
            if idx == -1:
                m = re.search(r"\d{8}", linea)
                if not m:
                    continue
                idx = m.start()
            tipo = linea[idx + 14: idx + 19].strip()
            tipo_via_dict[(mun, vcode)] = tipo

    print("2. Leyendo TRAM...")
    cusec_cp_tramos = defaultdict(lambda: defaultdict(int))
    total_tramos_cp = defaultdict(int)
    urba_tramos_cp = defaultdict(int)
    mun_tramos_cp = defaultdict(lambda: defaultdict(int))

    with open(fichero_tram, "r", encoding="latin1") as f:
        for linea in f:
            if len(linea) < 165 or linea[0:2] != PROVINCIA:
                continue
            cusec = linea[0:10].strip()
            mun = linea[0:5]
            vcode = linea[160:165]
            cpos = linea[42:47].strip()

            if not (cpos.isdigit() and len(cpos) == 5):
                continue

            cusec_cp_tramos[cusec][cpos] += 1
            total_tramos_cp[cpos] += 1
            mun_tramos_cp[cpos][mun] += 1

            if tipo_via_dict.get((mun, vcode)) in TIPOS_VIA_URBANIZACION:
                urba_tramos_cp[cpos] += 1

    print("3. Leyendo Excel de renta por HOGAR del INE...")
    df_bruto = pd.read_excel(fichero_renta, header=None)

    fila_cabecera = -1
    for i, row in df_bruto.iterrows():
        if row.astype(str).str.contains("renta neta", case=False, na=False).any():
            fila_cabecera = i
            break
    if fila_cabecera == -1:
        raise ValueError("No se encontró la cabecera de renta en el Excel del INE.")

    # NUEVO: Propagamos los nombres de las celdas combinadas hacia la derecha (forward fill)
    row_indicadores = df_bruto.iloc[fila_cabecera].ffill().astype(str).str.lower()
    
    # Identificar las posiciones de Renta Neta y Bruta por Hogar (devuelve todas las columnas, una por año)
    cols_neta = [i for i, ind in enumerate(row_indicadores) if "renta neta media por hogar" in ind or "renta neta por hogar" in ind]
    cols_bruta = [i for i, ind in enumerate(row_indicadores) if "renta bruta media por hogar" in ind or "renta bruta por hogar" in ind]

    if not cols_neta or not cols_bruta:
        cols_neta = [i for i, ind in enumerate(row_indicadores) if "renta neta" in ind]
        cols_bruta = [i for i, ind in enumerate(row_indicadores) if "renta bruta" in ind]

    # Asignamos la primera columna encontrada como el año actual y la segunda como el año anterior
    col_neta_act = cols_neta[0]
    col_bruta_act = cols_bruta[0]
    
    col_neta_ant = cols_neta[1] if len(cols_neta) > 1 else None
    col_bruta_ant = cols_bruta[1] if len(cols_bruta) > 1 else None

    indices = [0, col_neta_act, col_bruta_act]
    if col_neta_ant: indices.append(col_neta_ant)
    if col_bruta_ant: indices.append(col_bruta_ant)

    df_renta = df_bruto.iloc[fila_cabecera + 2:, indices].copy()
    
    col_names = ["Territorio", "Renta_Neta_Hogar", "Renta_Bruta_Hogar"]
    if col_neta_ant: col_names.append("Renta_Neta_Hogar_Ant")
    if col_bruta_ant: col_names.append("Renta_Bruta_Hogar_Ant")
    df_renta.columns = col_names

    df_renta["CUSEC"] = df_renta["Territorio"].astype(str).str.extract(r"^(\d{10})")
    df_renta = df_renta.dropna(subset=["CUSEC"]).copy()

    for col in df_renta.columns:
        if col not in ["Territorio", "CUSEC"]:
            df_renta[col] = pd.to_numeric(df_renta[col], errors="coerce")

    renta_por_cusec = df_renta.set_index("CUSEC").to_dict("index")

    nombre_mun_por_codigo = {}
    for territorio in df_renta["Territorio"].astype(str):
        cod_mun = territorio[:5]
        if cod_mun not in nombre_mun_por_codigo:
            resto = territorio[10:].strip()
            match_secc = re.search(r"\s+(sección|distrito)\s+.*$", resto, re.IGNORECASE)
            corte = match_secc.start() if match_secc else None
            nombre_mun_por_codigo[cod_mun] = resto[:corte].strip() if corte else resto

    print("4. Ponderando renta del hogar y evolución por CP...")
    suma_neta_act, suma_bruta_act = defaultdict(float), defaultdict(float)
    suma_neta_ant, suma_bruta_ant = defaultdict(float), defaultdict(float)
    peso_neta_act, peso_bruta_act = defaultdict(float), defaultdict(float)
    peso_neta_ant, peso_bruta_ant = defaultdict(float), defaultdict(float)
    secciones_por_cp = defaultdict(set)

    for cusec, cp_counts in cusec_cp_tramos.items():
        renta = renta_por_cusec.get(cusec)
        if not renta:
            continue
        
        r_neta = renta.get("Renta_Neta_Hogar")
        r_bruta = renta.get("Renta_Bruta_Hogar")
        r_neta_ant = renta.get("Renta_Neta_Hogar_Ant")
        r_bruta_ant = renta.get("Renta_Bruta_Hogar_Ant")

        for cp, peso in cp_counts.items():
            secciones_por_cp[cp].add(cusec)
            if pd.notna(r_neta):
                suma_neta_act[cp] += peso * r_neta
                peso_neta_act[cp] += peso
            if pd.notna(r_neta_ant):
                suma_neta_ant[cp] += peso * r_neta_ant
                peso_neta_ant[cp] += peso

            if pd.notna(r_bruta):
                suma_bruta_act[cp] += peso * r_bruta
                peso_bruta_act[cp] += peso
            if pd.notna(r_bruta_ant):
                suma_bruta_ant[cp] += peso * r_bruta_ant
                peso_bruta_ant[cp] += peso

    print("5. Construyendo informe final...")
    filas = []
    for cp, total_tramos in total_tramos_cp.items():
        pn_act = peso_neta_act.get(cp, 0)
        pn_ant = peso_neta_ant.get(cp, 0)
        pb_act = peso_bruta_act.get(cp, 0)
        pb_ant = peso_bruta_ant.get(cp, 0)

        rn_act = (suma_neta_act[cp] / pn_act) if pn_act > 0 else None
        rn_ant = (suma_neta_ant[cp] / pn_ant) if pn_ant > 0 else None
        rb_act = (suma_bruta_act[cp] / pb_act) if pb_act > 0 else None
        rb_ant = (suma_bruta_ant[cp] / pb_ant) if pb_ant > 0 else None

        evo_neta = (((rn_act - rn_ant) / rn_ant) * 100) if (rn_act is not None and rn_ant is not None and rn_ant != 0) else None
        evo_bruta = (((rb_act - rb_ant) / rb_ant) * 100) if (rb_act is not None and rb_ant is not None and rb_ant != 0) else None

        n_urba = urba_tramos_cp.get(cp, 0)
        pct_urba = (n_urba / total_tramos * 100) if total_tramos else 0
        cobertura = (pb_act / total_tramos * 100) if total_tramos else 0

        reparto_mun = sorted(mun_tramos_cp.get(cp, {}).items(), key=lambda kv: kv[1], reverse=True)
        codigo_ine_principal = reparto_mun[0][0] if reparto_mun else None
        municipio_principal = nombre_mun_por_codigo.get(codigo_ine_principal, "?") if reparto_mun else "?"

        filas.append({
            "Codigo_Postal": cp,
            "Municipio": municipio_principal,
            "Codigo_INE_Principal": codigo_ine_principal,
            "Renta_Bruta_Hogar_Media": round(rb_act, 2) if rb_act else None,
            "Evolución_Bruta_%": round(evo_bruta, 2) if evo_bruta is not None else None,
            "Renta_Neta_Hogar_Media": round(rn_act, 2) if rn_act else None,
            "Evolución_Neta_%": round(evo_neta, 2) if evo_neta is not None else None,
            "Pct_Urbanizaciones": round(pct_urba, 2),
            "Nº_Tramos_Urbanizacion": n_urba,
            "Pct_Cobertura_Renta": round(cobertura, 2),
            "Secciones_Consideradas": len(secciones_por_cp.get(cp, [])),
            "Tramos_Totales": total_tramos,
            "Fiabilidad_Renta": "Baja (<50% cobertura)" if cobertura < 50 else "OK"
        })

    df_final = pd.DataFrame(filas)
    df_final["Codigo_Postal"] = df_final["Codigo_Postal"].astype(str).str.zfill(5)
    df_final = df_final.sort_values(["Renta_Bruta_Hogar_Media", "Nº_Tramos_Urbanizacion"], ascending=[False, False])

    df_final.to_excel(SALIDA, index=False, sheet_name="Renta_y_Urba_por_CP")
    print(f"Informe guardado en {SALIDA.name}")


if __name__ == "__main__":
    main()