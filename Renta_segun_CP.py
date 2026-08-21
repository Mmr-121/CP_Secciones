"""
generar_informe_renta_cp.py

----------------------------------------
Genera un Excel con, para cada Código Postal de la provincia de Alicante:
  - Renta neta media ponderada por persona
  - Renta bruta media ponderada por persona
  - % de tramos de tipo "urbanización" (URB/URBAT)
  - Nº de tramos totales y nº de secciones censales (CUSEC) que aportan renta
  - % de cobertura de renta (cuántos tramos del CP tienen dato de renta real,
    frente a secreto estadístico)
----------------------------------------

El INE publica la renta a nivel de SECCIÓN CENSAL (CUSEC), no de Código Postal.
Una misma sección puede repartir sus tramos de calle entre 2+ CP (muy habitual
en zonas urbanas densas).

La renta de cada sección se reparte proporcionalmente: si la sección X
tiene 30 tramos en el CP 03008 y 20 en el 03138, su renta pondera un 60% en la
media del 03008 y un 40% en la del 03138 (peso = nº de tramos, no el 100/0).

Las secciones con secreto estadístico (renta = "." o vacío en el Excel del
INE) se EXCLUYEN tanto del numerador como del denominador de la media
ponderada, para que no infrainflen (deflacten) el resultado.

El % de urbanizaciones NO necesita este reparto: cada tramo de calle ya tiene
un único CP inequívoco en el propio fichero TRAM, la ambigüedad solo existe a
nivel de sección/renta.
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
    """Busca la carpeta de la edición del INE (ej. caj_esp_072026) sin
    depender de su nombre exacto, ya que cambia con cada edición trimestral.
    Acepta tanto la carpeta "caj_esp_*" en la raíz como su variante anidada
    (caj_esp_XXXXXX/caj_esp_XXXXXX, tal y como la distribuye el INE en el zip).
    Si hay más de una edición presente, avisa en vez de adivinar en silencio."""
    candidatas = sorted(BASE_DIR.glob("caj_esp_*"))
    candidatas = [c for c in candidatas if c.is_dir()]
    if not candidatas:
        raise FileNotFoundError(
            f"No se encontró ninguna carpeta 'caj_esp_*' con los datos del INE dentro de {BASE_DIR}. "
            "Descomprime aquí el zip de la última edición (VIAS/TRAM) antes de ejecutar."
        )
    if len(candidatas) > 1:
        print(
            f"[AVISO] Hay más de una carpeta 'caj_esp_*' en {BASE_DIR}: "
            f"{[c.name for c in candidatas]}. Se usará '{candidatas[-1].name}' (la última alfabéticamente). "
            "Si no es la edición correcta, borra o renombra las carpetas antiguas antes de ejecutar."
        )
    carpeta = candidatas[-1]
    anidada = carpeta / carpeta.name
    return anidada if anidada.is_dir() else carpeta


FOLDER_DATOS = _localizar_carpeta_datos()

PROVINCIA = "03"  # Alicante
TIPOS_VIA_URBANIZACION = {"URB", "URBAT"}

SALIDA = BASE_DIR / "Informe_Renta_Urbanizacion_por_CP.xlsx"
SALIDA_RELACION = BASE_DIR / "Relacion_Secciones_CP.xlsx"


def encontrar_fichero(patron: str) -> Path:
    candidatos = sorted(glob.glob(str(FOLDER_DATOS / f"{patron}*")))
    candidatos = [c for c in candidatos if not c.endswith(".xlsx")]
    if not candidatos:
        raise FileNotFoundError(
            f"No se encontró ningún fichero que empiece por '{patron}' en {FOLDER_DATOS}"
        )
    return Path(candidatos[0])


def encontrar_excel_renta() -> Path:
    candidatos = sorted(glob.glob(str(BASE_DIR / "*.xlsx")))
    candidatos = [
        c for c in candidatos
        if "Informe_Renta" not in c and "Relacion_Secciones" not in c
    ]
    if not candidatos:
        raise FileNotFoundError(f"No se encontró el Excel de renta del INE en {BASE_DIR}")
    if len(candidatos) > 1:
        print(
            f"[AVISO] Hay más de un .xlsx candidato a Excel de renta en {BASE_DIR}: "
            f"{[Path(c).name for c in candidatos]}. Se usará '{Path(candidatos[0]).name}'. "
            "Si no es el correcto, borra los ficheros de renta de años anteriores."
        )
    return Path(candidatos[0])


def main():
    fichero_vias = encontrar_fichero("VIAS")
    fichero_tram = encontrar_fichero("TRAM")
    fichero_renta = encontrar_excel_renta()

    print(f"VIAS : {fichero_vias.name}")
    print(f"TRAM : {fichero_tram.name}")
    print(f"Renta: {fichero_renta.name}")

    # VIAS: tipo de vía por (municipio, codigo_via), para detectar urbanizaciones (URB / URBAT)
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
                # Fallback por si la fecha de vigencia cambia en próximas descargas
                m = re.search(r"\d{8}", linea)
                if not m:
                    continue
                idx = m.start()
            tipo = linea[idx + 14: idx + 19].strip()
            tipo_via_dict[(mun, vcode)] = tipo

    # TRAM: por cada tramo sabemos: CUSEC, CP y si su  vía es de tipo urbanización. Con eso construimos:
    #   - cusec_cp_tramos[cusec][cp] = nº de tramos (para ponderar renta)
    #   - total_tramos_cp[cp], urba_tramos_cp[cp] (para % urbanización)

    print("2. Leyendo TRAM (puede tardar unos minutos, es un fichero grande)...")
    cusec_cp_tramos = defaultdict(lambda: defaultdict(int))
    total_tramos_cp = defaultdict(int)
    urba_tramos_cp = defaultdict(int)
    mun_tramos_cp = defaultdict(lambda: defaultdict(int))  # cp -> {codigo_municipio: nº tramos}

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

            tipo = tipo_via_dict.get((mun, vcode))
            if tipo in TIPOS_VIA_URBANIZACION:
                urba_tramos_cp[cpos] += 1

    # Excel de renta del INE -> renta neta/bruta media por CUSEC
    print("3. Leyendo Excel de renta del INE...")
    df_bruto = pd.read_excel(fichero_renta, header=None)

    fila_cabecera = -1
    for i, row in df_bruto.iterrows():
        if row.astype(str).str.contains("renta neta", case=False, na=False).any():
            fila_cabecera = i
            break
    if fila_cabecera == -1:
        raise ValueError("No se encontró la fila de cabecera con 'renta neta' en el Excel del INE.")

    fila_hdr = df_bruto.iloc[fila_cabecera].astype(str).str.lower()
    col_neta_idx = fila_hdr.str.contains("renta neta media por persona").idxmax()
    col_bruta_idx = fila_hdr.str.contains("renta bruta media por persona").idxmax()

    df_renta = df_bruto.iloc[fila_cabecera + 2:, [0, col_neta_idx, col_bruta_idx]].copy()
    df_renta.columns = ["Territorio", "Renta_Neta", "Renta_Bruta"]
    df_renta["CUSEC"] = df_renta["Territorio"].astype(str).str.extract(r"^(\d{10})")
    df_renta = df_renta.dropna(subset=["CUSEC"]).copy()
    df_renta["Renta_Neta"] = pd.to_numeric(df_renta["Renta_Neta"], errors="coerce")
    df_renta["Renta_Bruta"] = pd.to_numeric(df_renta["Renta_Bruta"], errors="coerce")

    renta_por_cusec = df_renta.set_index("CUSEC")[["Renta_Neta", "Renta_Bruta"]].to_dict("index")

    # Nombre de municipio por código INE de municipio (5 dígitos), a partir de la
    # columna "Territorio" del propio Excel de renta (ej. "0300101001 Alcoi/Alcoy sección 01001")
    print("3b. Extrayendo nombres de municipio...")
    nombre_mun_por_codigo = {}
    for territorio in df_renta["Territorio"].astype(str):
        cod_mun = territorio[:5]
        if cod_mun in nombre_mun_por_codigo:
            continue
        resto = territorio[10:].strip()
        match_secc = re.search(r"\s+sección\s+.*$", resto, re.IGNORECASE)
        match_dist = re.search(r"\s+distrito\s+.*$", resto, re.IGNORECASE)
        corte = min([m.start() for m in (match_secc, match_dist) if m], default=None)
        nombre_mun_por_codigo[cod_mun] = resto[:corte].strip() if corte is not None else resto

    # Reparto ponderado de la renta por CP
    #    peso = nº de tramos de esa sección que caen en ese CP
    #    Se excluyen del numerador Y del denominador las secciones sin dato
    #    de renta (secreto estadístico), para no deflactar la media.
    print("4. Repartiendo la renta de cada sección entre los CP según sus tramos...")
    suma_ponderada_neta = defaultdict(float)
    suma_ponderada_bruta = defaultdict(float)
    peso_total_neta = defaultdict(float)
    peso_total_bruta = defaultdict(float)
    secciones_por_cp = defaultdict(set)

    secciones_sin_renta = 0
    for cusec, cp_counts in cusec_cp_tramos.items():
        renta = renta_por_cusec.get(cusec)
        if renta is None:
            secciones_sin_renta += 1
            continue
        r_neta, r_bruta = renta["Renta_Neta"], renta["Renta_Bruta"]
        for cp, peso in cp_counts.items():
            secciones_por_cp[cp].add(cusec)
            if pd.notna(r_neta):
                suma_ponderada_neta[cp] += peso * r_neta
                peso_total_neta[cp] += peso
            if pd.notna(r_bruta):
                suma_ponderada_bruta[cp] += peso * r_bruta
                peso_total_bruta[cp] += peso

    # Tabla final por Código Postal
    print("5. Construyendo tabla final...")
    filas = []
    for cp, total_tramos in total_tramos_cp.items():
        peso_neta = peso_total_neta.get(cp, 0)
        peso_bruta = peso_total_bruta.get(cp, 0)
        renta_neta_media = (suma_ponderada_neta[cp] / peso_neta) if peso_neta > 0 else None
        renta_bruta_media = (suma_ponderada_bruta[cp] / peso_bruta) if peso_bruta > 0 else None
        n_urba = urba_tramos_cp.get(cp, 0)
        pct_urba = (n_urba / total_tramos * 100) if total_tramos else 0
        cobertura = (peso_bruta / total_tramos * 100) if total_tramos else 0

        # Municipio principal (el que más tramos aporta a este CP) + resto si lo comparte
        reparto_mun = sorted(mun_tramos_cp.get(cp, {}).items(), key=lambda kv: kv[1], reverse=True)
        codigo_ine_principal = reparto_mun[0][0] if reparto_mun else None
        municipio_principal = nombre_mun_por_codigo.get(codigo_ine_principal, "?") if reparto_mun else "?"
        otros_municipios = ", ".join(
            f"{nombre_mun_por_codigo.get(cod, cod)} ({tramos})"
            for cod, tramos in reparto_mun[1:]
        ) if len(reparto_mun) > 1 else ""

        filas.append({
            "Codigo_Postal": cp,
            "Municipio": municipio_principal,
            "Codigo_INE_Principal": codigo_ine_principal,
            "Otros_Municipios_en_este_CP": otros_municipios,
            "Renta_Neta_Media": round(renta_neta_media, 2) if renta_neta_media is not None else None,
            "Renta_Bruta_Media": round(renta_bruta_media, 2) if renta_bruta_media is not None else None,
            "Pct_Urbanizaciones": round(pct_urba, 2),
            "Nº_Tramos_Urbanizacion": n_urba,
            "Pct_Cobertura_Renta": round(cobertura, 2),
            "Secciones_Consideradas": len(secciones_por_cp.get(cp, [])),
            "Tramos_Totales": total_tramos,
        })

    df_final = pd.DataFrame(filas)
    df_final["Codigo_Postal"] = df_final["Codigo_Postal"].astype(str).str.zfill(5)

    # Aviso de fiabilidad: si menos del 50% de los tramos del CP tienen renta
    # real (resto es secreto estadístico), la media puede no ser representativa.
    df_final["Fiabilidad_Renta"] = df_final["Pct_Cobertura_Renta"].apply(
        lambda p: "Baja (<50% cobertura)" if p < 50 else "OK"
    )

    # Orden pensado para targeting: primero por renta (de más a menos),
    # y dentro de la misma renta, por volumen de tramos de urbanización.
    df_final = df_final.sort_values(
        ["Renta_Bruta_Media", "Nº_Tramos_Urbanizacion"], ascending=[False, False]
    )

    print(f"6. Guardando '{SALIDA.name}'...")
    df_final.to_excel(SALIDA, index=False, sheet_name="Renta_y_Urba_por_CP")

    print("\nCompletado.")
    print(f"  CP procesados        : {len(df_final)}")
    print(f"  Secciones sin renta  : {secciones_sin_renta} (secreto estadístico / no publicado)")
    print(f"  Fichero de salida    : {SALIDA}")


if __name__ == "__main__":
    main()