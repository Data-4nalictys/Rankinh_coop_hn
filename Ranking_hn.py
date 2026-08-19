"""
actualizar_ranking_hn.py — Lee los archivos del CONSUCOOP y actualiza los tres CSV base.

USO:
    python actualizar_ranking_hn.py --eeff    "ruta\8.-Estados-financieros.xlsx"
    python actualizar_ranking_hn.py --afiliados "ruta\5.-Detalle-de-afiliados-por-Genero.xlsx"
    python actualizar_ranking_hn.py --eeff "..." --afiliados "..."

HOJAS QUE LEE:
    EEFF:      "Balance General"  y  "Estado de Resultado"
    Afiliados: "DETALLE DE AFILIADO POR GENERO"
"""

import argparse, re, unicodedata
from pathlib import Path
import numpy as np
import pandas as pd

# ── CONFIGURACION ────────────────────────────────────────

BASE = Path(r"C:\Users\RONYL\OneDrive - Cooperativa de Ahorro y Credito Sonaguera Limitada\Documentos\GitHub\ranking_coop_hn\data")
RUTA_MAESTRO  = BASE / "cooperativashn.csv"
RUTA_BG       = BASE / "Balance_General.csv"
RUTA_ER       = BASE / "Estado_de_Resultados.csv"
RUTA_AFIL     = BASE / "Afiliados.csv"
RUTA_NO_MATCH = BASE / "cooperativas_sin_match.txt"
RUTA_IND      = BASE / "Indicadores.csv"
HOJA_BG       = "Balance General"
HOJA_ER       = "Estado de Resultado"
HOJA_AFIL     = "DETALLE DE AFILIADO POR GENERO"
MESES_ES = {"enero":1,
            "febrero":2,
            "marzo":3,
            "abril":4,
            "mayo":5,
            "junio":6,
            "julio":7,
            "agosto":8,
            "septiembre":9,
            "octubre":10,
            "noviembre":11,
            "diciembre":12}

# ── COLUMNAS DE SALIDA ────────────────────────────────────
BG_COLS = ["COOPERATIVA","MES#","AÑO","ACTIVOS TOTALES","ACTIVOS NETOS","DISPONIBILIDADES",
           "INVERSIONES","PRESTAMOS","CUENTAS X COBRAR","ACTIVOS EVENTUALES",
           "PROPIEDAD PLANTA EQUIPO","CARGOS DIFERIDOS","ACTIVOS DE INVERSION","ACTIVOS INTANGIBLES",
           "PASIVOS TOTALES","EXIGIBILIDADES INMEDIATAS","EXIGIBILIDADES X DEPOSITOS",
           "Depositos Ahorro hn","DPF hn","OTRAS EXIGIBILIDADES","OBLIGACIONES BANCARIAS",
           "CREDITOS DIFERIDOS","PROVISIONES DE VALUACION","PROVISIONES EVENTUALES","PATRIMONIO",
           "PATRIMONIO PRIMARIO","Aportaciones","Reserva","PATRIMONIO COMPLEMENTARIO"]

ER_COLS = ["COOPERATIVA","MES#","AÑO","INGRESOS","Productos Financieros","Productos x Servicios",
           "Otros Productos","Ingresos Anteriores","EGRESOS","Gastos Financieros",
           "Gastos de Administracion","Otros Gastos","Gastos Anteriores","EXCEDENTES"]

AFIL_COLS = ["Nombre de Cooperativa","Name Mes","#Mes","Año","Afiliado Hombres",
             "Afiliados Mujeres","Personas Jurídicas","Menores Ahorrantes",
             "Total Afiliados ","Cooperativa","Cierre"]

RENAME_BG = {
    "ACTIVOS TOTALES ":"ACTIVOS TOTALES","ACTIVOS TOTALES":"ACTIVOS TOTALES",
    "ACTIVOS NETOS":"ACTIVOS NETOS","DISPONIBILIDADES":"DISPONIBILIDADES","INVERSIONES":"INVERSIONES",
    "PRÉSTAMOS, DESCUENTOS Y NEGOCIACIONES":"PRESTAMOS","PRESTAMOS, DESCUENTOS Y NEGOCIACIONES":"PRESTAMOS",
    "CUENTAS Y DOCUMENTOS POR COBRAR":"CUENTAS X COBRAR","ACTIVOS EVENTUALES":"ACTIVOS EVENTUALES",
    "PROPIEDAD, PLANTA Y EQUIPO":"PROPIEDAD PLANTA EQUIPO","CARGOS DIFERIDOS":"CARGOS DIFERIDOS",
    "ACTIVOS DE INVERSIÓN":"ACTIVOS DE INVERSION","ACTIVOS INTANGIBLES":"ACTIVOS INTANGIBLES",
    "PASIVOS TOTALES":"PASIVOS TOTALES","EXIGIBILIDADES INMEDIATAS":"EXIGIBILIDADES INMEDIATAS",
    "EXIGIBILIDADES POR DEPÓSITOS":"EXIGIBILIDADES X DEPOSITOS",
    "Depositos De Ahorro m/n":"Depositos Ahorro hn","Depositos A Plazo m/n":"DPF hn",
    "OTRAS EXIGIBILIDADES":"OTRAS EXIGIBILIDADES","OBLIGACIONES BANCARIAS":"OBLIGACIONES BANCARIAS",
    "CREDITOS DIFERIDOS":"CREDITOS DIFERIDOS","PROVISIONES DE VALUACION":"PROVISIONES DE VALUACION",
    "PROVISIONES EVENTUALES":"PROVISIONES EVENTUALES","PATRIMONIO":"PATRIMONIO",
    "PATRIMONIO PRIMARIO":"PATRIMONIO PRIMARIO","Aportaciones":"Aportaciones",
    "Reserva Legal":"Reserva","PATRIMONIO COMPLEMENTARIO":"PATRIMONIO COMPLEMENTARIO",
}

RENAME_ER = {
    "INGRESOS":"INGRESOS","PRODUCTOS FINANCIEROS":"Productos Financieros",
    "PRODUCTOS POR SERVICIOS":"Productos x Servicios","OTROS PRODUCTOS":"Otros Productos",
    "INGRESOS DE EJERCICIOS ANTERIORES":"Ingresos Anteriores",
    "GASTOS FINANCIEROS":"Gastos Financieros","GASTOS DE ADMINISTRACION":"Gastos de Administracion",
    "OTROS GASTOS":"Otros Gastos","GASTOS DE EJERCICIOS ANTERIORES":"Gastos Anteriores",
    "EXCEDENTES O PÉRDIDAS DEL PERIODO":"EXCEDENTES","EXCEDENTES O PERDIDAS DEL PERIODO":"EXCEDENTES",
}

# ── INDICADORES FINANCIEROS (CONSUCOOP) ───────────────────
# Metas oficiales tomadas del Excel de CONSUCOOP compartido. El Excel solo
# trae el corte final de cada meta (cumple / no cumple), asi que las 4
# bandas A-E intermedias se estiman proporcionalmente alrededor de esa meta
# (igual que en indicadores.html). direccion: "desc" = mas alto es mejor,
# "asc" = mas bajo es mejor.
IND_COLS = [
    "COOPERATIVA", "MES#", "AÑO",
    "SOLVENCIA", "RANGO SOLVENCIA",
    "MOROSIDAD", "RANGO MOROSIDAD",
    "COBERTURA MORA", "RANGO COBERTURA MORA",
    "LIQUIDEZ", "RANGO LIQUIDEZ",
    "ROA", "RANGO ROA",
    "ROE", "RANGO ROE",
    "AUTOSUFICIENCIA", "RANGO AUTOSUFICIENCIA",
]

METAS_IND = {
    "SOLVENCIA":        (8.5, "desc"),    # 3.1 Indice de Capital Institucional
    "MOROSIDAD":        (5.0, "asc"),     # 2.2 Indice de Morosidad
    "COBERTURA MORA":   (100.0, "desc"),  # 2.1 Suficiencia de Provision para Cartera
    "LIQUIDEZ":         (9.0, "desc"),    # 4.1 Cobertura oblig. depositarias (M.N.)
    "ROA":              (0.5, "desc"),    # 6.1 Rentabilidad s/ Activos Netos Promedio
    "ROE":              (10.0, "desc"),   # referencial, sin parametro CONSUCOOP
    "AUTOSUFICIENCIA":  (110.0, "desc"),  # 5.1 Autosuficiencia Operativa
}

# ── FUNCIONES CORE ────────────────────────────────────────
def normalizar(texto):
    if not isinstance(texto, str): return ""
    res = []
    for l in unicodedata.normalize("NFKD", texto):
        if l in ("ñ","Ñ"): res.append(l); continue
        if unicodedata.category(l) == "Mn": continue
        res.append(l)
    return "".join(res).upper().strip()

def cargar_maestro(ruta):
    df = pd.read_csv(ruta, encoding="utf-8-sig", dtype=str)
    df.columns = df.columns.str.strip()
    return {re.sub(r"\s+"," ",normalizar(str(r["Nombre Completo"]))): str(r["Abreviatura"]).strip()
            for _,r in df.iterrows()}

def resolver(texto, mapa, sin_match):
    if not isinstance(texto, str) or not texto.strip(): return ""
    t = re.sub(r"\s+", " ", normalizar(texto))
    # Quitar texto suelto despues del ultimo parentesis (ej: "( COOHDETUR) Febrero 2023")
    t = re.sub(r"\)\s+[A-Za-z].+$", ")", t).strip()
    # Variantes de limpieza progresiva
    t2 = re.sub(r",?\s*(LIMITADA|R\.L\.?|RL)\s*\.?$", "", t).strip()
    t2 = re.sub(r"\s*\(.*?\)\s*$", "", t2).strip()
    t3 = re.sub(r"\s*\(.*?\)\s*", "", t).strip()
    t3 = re.sub(r",?\s*(LIMITADA|R\.L\.?|RL)\s*\.?$", "", t3).strip()
    for c in (t, t2, t3):
        if c in mapa: return mapa[c]
    for c, v in mapa.items():
        if c in t: return v
    m = re.search(r"COOPERATIVA MIXTA (.*?)(?:,| LIMITADA| R\.L|\(|$)", t)
    if m: return ("MIXTA " + m.group(1)).strip()
    m = re.search(r"COOPERATIVA DE AHORRO Y CREDITO (.*?)(?:,| LIMITADA| R\.L|\(|$)", t)
    if m: return m.group(1).strip()
    sin_match.add(texto.strip()); return t.strip()
def limpiar_num(s):
    s = s.astype(str).str.replace(",","",regex=False).str.replace(r"\..*$","",regex=True)
    return pd.to_numeric(s, errors="coerce").fillna(0).astype(int)

def encontrar_header(ruta, hoja, texto_exacto):
    """Encuentra la fila donde una celda ES EXACTAMENTE el texto buscado."""
    df_raw = pd.read_excel(ruta, sheet_name=hoja, header=None, dtype=str, nrows=15)
    for i, row in df_raw.iterrows():
        if texto_exacto.upper() in row.astype(str).str.strip().str.upper().tolist():
            return i
    return 0

def leer_hoja(ruta, hoja, texto_header):
    hdr = encontrar_header(ruta, hoja, texto_header)
    df  = pd.read_excel(ruta, sheet_name=hoja, header=hdr, dtype=str)
    df  = df.loc[:, ~df.columns.str.startswith("Unnamed")]
    df.columns = df.columns.str.strip()
    return df

def filtrar_ruido(df, col="COOPERATIVA"):
    df = df[df[col].notna()]
    df = df[df[col].astype(str).str.len() > 5]
    df = df[~df[col].astype(str).str.contains(
        r"FACACH|Fuente:|Superintendencia|N°|Balance|Estado|En Lempiras",
        case=False, na=False)]
    return df

def periodos_eeff(ruta):
    if not Path(ruta).exists(): return set()
    df = pd.read_csv(ruta, encoding="utf-8-sig", dtype=str, usecols=["MES#","AÑO"])
    return set(zip(df["MES#"].str.strip(), df["AÑO"].str.strip()))

def periodos_afil(ruta):
    if not Path(ruta).exists(): return set()
    df = pd.read_csv(ruta, encoding="utf-8-sig", dtype=str, usecols=["Cierre"])
    return set(df["Cierre"].str.strip())

# ── PROCESAR EEFF ─────────────────────────────────────────
def procesar_eeff(ruta, mapa, sin_match):

    # Balance General
    print(f"  Leyendo '{HOJA_BG}'...")
    bg = leer_hoja(ruta, HOJA_BG, "COOPERATIVA")
    bg = bg.rename(columns=RENAME_BG)
    bg = filtrar_ruido(bg)
    bg["COOPERATIVA"] = bg["COOPERATIVA"].apply(lambda x: resolver(x, mapa, sin_match))
    bg = bg[bg["COOPERATIVA"].str.strip() != ""]
    bg["MES#"] = bg["MES"].str.strip().str.lower().map(MESES_ES).fillna(0).astype(int).astype(str)
    bg["AÑO"]  = bg["AÑO"].astype(str).str.strip().str.split(".").str[0]
    for c in BG_COLS:
        if c not in ("COOPERATIVA","MES#","AÑO") and c in bg.columns:
            bg[c] = limpiar_num(bg[c])
    bg = bg.reindex(columns=BG_COLS, fill_value=0)

    exist = periodos_eeff(RUTA_BG)
    nuevos = bg[~bg.apply(lambda r: (r["MES#"], r["AÑO"]) in exist, axis=1)]
    if nuevos.empty:
        print("  Balance General: todos los periodos ya existen.")
    else:
        for anio, mes in sorted(nuevos[["AÑO","MES#"]].drop_duplicates().values,
                                key=lambda x: (x[0], x[1].zfill(2))):
            n = len(nuevos[(nuevos["AÑO"]==anio) & (nuevos["MES#"]==mes)])
            print(f"  + BG {mes.zfill(2)}-{anio}: {n} cooperativas")
        base = pd.read_csv(RUTA_BG, encoding="utf-8-sig", dtype=str) if RUTA_BG.exists() else pd.DataFrame(columns=BG_COLS)
        r = pd.concat([base, nuevos], ignore_index=True)
        r.to_csv(RUTA_BG, index=False, encoding="utf-8-sig")
        print(f"  → Balance_General.csv actualizado ({len(r)} filas)")

    # Estado de Resultados
    print(f"\n  Leyendo '{HOJA_ER}'...")
    er = leer_hoja(ruta, HOJA_ER, "COOPERATIVA")
    er = er.rename(columns=RENAME_ER)
    er = filtrar_ruido(er)
    er["COOPERATIVA"] = er["COOPERATIVA"].apply(lambda x: resolver(x, mapa, sin_match))
    er = er[er["COOPERATIVA"].str.strip() != ""]
    er["MES#"] = er["MES"].str.strip().str.lower().map(MESES_ES).fillna(0).astype(int).astype(str)
    er["AÑO"]  = er["AÑO"].astype(str).str.strip().str.split(".").str[0]
    for c in ER_COLS:
        if c not in ("COOPERATIVA","MES#","AÑO","EGRESOS") and c in er.columns:
            er[c] = limpiar_num(er[c])
    er["EGRESOS"] = (er.get("Gastos Financieros",    pd.Series(0,index=er.index)) +
                     er.get("Gastos de Administracion",pd.Series(0,index=er.index)) +
                     er.get("Otros Gastos",           pd.Series(0,index=er.index)) +
                     er.get("Gastos Anteriores",       pd.Series(0,index=er.index)))
    if "EXCEDENTES" not in er.columns or er["EXCEDENTES"].eq(0).all():
        er["EXCEDENTES"] = er["INGRESOS"] - er["EGRESOS"]
    er = er.reindex(columns=ER_COLS, fill_value=0)

    exist = periodos_eeff(RUTA_ER)
    nuevos = er[~er.apply(lambda r: (r["MES#"], r["AÑO"]) in exist, axis=1)]
    if nuevos.empty:
        print("  Estado de Resultados: todos los periodos ya existen.")
    else:
        for anio, mes in sorted(nuevos[["AÑO","MES#"]].drop_duplicates().values,
                                key=lambda x: (x[0], x[1].zfill(2))):
            n = len(nuevos[(nuevos["AÑO"]==anio) & (nuevos["MES#"]==mes)])
            print(f"  + ER {mes.zfill(2)}-{anio}: {n} cooperativas")
        base = pd.read_csv(RUTA_ER, encoding="utf-8-sig", dtype=str) if RUTA_ER.exists() else pd.DataFrame(columns=ER_COLS)
        r = pd.concat([base, nuevos], ignore_index=True)
        r.to_csv(RUTA_ER, index=False, encoding="utf-8-sig")
        print(f"  → Estado_de_Resultados.csv actualizado ({len(r)} filas)")

# ── INDICADORES: FUNCIONES CORE ───────────────────────────
def a_float(col):
    """Convierte una columna string (con comas de miles) a numerico float."""
    return pd.to_numeric(
        col.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce"
    ).fillna(0.0)


def calificar(valor, meta, direccion):
    """Devuelve la letra A-E segun que tan lejos esta el valor de la meta."""
    if pd.isna(valor):
        return ""
    if direccion == "desc":
        ratio = valor / meta if meta else 0
        if ratio >= 1.30: return "A"
        if ratio >= 1.00: return "B"
        if ratio >= 0.85: return "C"
        if ratio >= 0.60: return "D"
        return "E"
    else:
        if valor <= meta * 0.50: return "A"
        if valor <= meta:        return "B"
        if valor <= meta * 1.40: return "C"
        if valor <= meta * 2.00: return "D"
        return "E"


def cargar_balance_indicadores():
    bg = pd.read_csv(RUTA_BG, encoding="utf-8-sig", dtype=str)
    bg.columns = bg.columns.str.strip()

    campos = [
        "ACTIVOS TOTALES", "ACTIVOS NETOS", "PRESTAMOS", "PATRIMONIO",
        "PROVISIONES DE VALUACION", "DISPONIBILIDADES",
        "EXIGIBILIDADES INMEDIATAS", "EXIGIBILIDADES X DEPOSITOS",
        "Vigentes hn", "Atrasados hn", "Vencidos hn", "Demanda Judicial hn",
    ]
    for c in campos:
        if c not in bg.columns:
            print(f"  Advertencia: '{c}' no existe en {RUTA_BG.name} (se asume 0). "
                  f"Revisa BG_COLS si es una columna de mora/depositos.")
            bg[c] = 0

    # Si Atrasados/Vencidos/Demanda Judicial vienen las 3 en blanco (no
    # reportadas todavia), la cartera en riesgo debe quedar SIN DATO, no en
    # 0% -- de lo contrario Morosidad se ve falsamente "sana".
    mora_cols = ["Atrasados hn", "Vencidos hn", "Demanda Judicial hn"]
    sin_mora_reportada = bg[mora_cols].apply(
        lambda col: col.isna() | col.astype(str).str.strip().eq("")
    ).all(axis=1)

    for c in campos:
        bg[c] = a_float(bg[c])

    bg["CARTERA EN RIESGO"] = bg["Atrasados hn"] + bg["Vencidos hn"] + bg["Demanda Judicial hn"]
    bg.loc[sin_mora_reportada, "CARTERA EN RIESGO"] = np.nan
    bg["EXIGIBILIDADES TOTALES"] = bg["EXIGIBILIDADES INMEDIATAS"] + bg["EXIGIBILIDADES X DEPOSITOS"]
    return bg


def cargar_resultados_indicadores():
    er = pd.read_csv(RUTA_ER, encoding="utf-8-sig", dtype=str)
    er.columns = er.columns.str.strip()
    for c in ("INGRESOS", "EGRESOS", "EXCEDENTES"):
        er[c] = a_float(er[c]) if c in er.columns else 0
    return er


def calcular_indicadores(bg, er):
    df = bg.merge(
        er[["COOPERATIVA", "MES#", "AÑO", "INGRESOS", "EGRESOS", "EXCEDENTES"]],
        on=["COOPERATIVA", "MES#", "AÑO"], how="left"
    )

    df["SOLVENCIA"] = (df["PATRIMONIO"] / df["ACTIVOS TOTALES"].replace(0, np.nan)) * 100
    df["MOROSIDAD"] = (df["CARTERA EN RIESGO"] / df["PRESTAMOS"].replace(0, np.nan)) * 100
    df["COBERTURA MORA"] = (df["PROVISIONES DE VALUACION"] / df["CARTERA EN RIESGO"].replace(0, np.nan)) * 100
    df["LIQUIDEZ"] = (df["DISPONIBILIDADES"] / df["EXIGIBILIDADES TOTALES"].replace(0, np.nan)) * 100
    df["ROA"] = (df["EXCEDENTES"] / df["ACTIVOS NETOS"].replace(0, np.nan)) * 100
    df["ROE"] = (df["EXCEDENTES"] / df["PATRIMONIO"].replace(0, np.nan)) * 100
    df["AUTOSUFICIENCIA"] = (df["INGRESOS"] / df["EGRESOS"].replace(0, np.nan)) * 100

    for ind, (meta, direccion) in METAS_IND.items():
        df[f"RANGO {ind}"] = df[ind].apply(lambda v: calificar(v, meta, direccion))

    for ind in METAS_IND:
        df[ind] = df[ind].round(2)

    return df.reindex(columns=IND_COLS)


def generar_indicadores():
    """Recalcula Indicadores.csv completo a partir de Balance_General.csv y
    Estado_de_Resultados.csv ya actualizados. Se corre siempre al final,
    tenga o no la corrida actual archivos --eeff/--afiliados nuevos."""
    print("\nIndicadores Financieros (CONSUCOOP):")
    if not RUTA_BG.exists() or not RUTA_ER.exists():
        print("  Advertencia: faltan Balance_General.csv y/o Estado_de_Resultados.csv, "
              "no se genera Indicadores.csv todavia.")
        return

    bg = cargar_balance_indicadores()
    er = cargar_resultados_indicadores()
    ind = calcular_indicadores(bg, er)
    ind.to_csv(RUTA_IND, index=False, encoding="utf-8-sig")
    print(f"  → {RUTA_IND.name} generado ({len(ind)} filas)")

    sin_mora = ind["MOROSIDAD"].isna().sum()
    if sin_mora:
        print(f"  Advertencia: {sin_mora} fila(s) sin desglose de mora reportado "
              f"(Atrasados/Vencidos/Demanda Judicial) -> MOROSIDAD queda vacio ahi.")


# ── PROCESAR AFILIADOS ────────────────────────────────────
def procesar_afiliados(ruta, mapa, sin_match):
    print(f"  Leyendo '{HOJA_AFIL}'...")
    af = leer_hoja(ruta, HOJA_AFIL, "Nombre de Cooperativa")
    col_nombre = next((c for c in af.columns if "nombre" in c.lower() and "cooperativa" in c.lower()), None)
    col_mes    = next((c for c in af.columns if normalizar(c) == "MES"), None)
    col_anio   = next((c for c in af.columns if normalizar(c) in ("AÑO","ANO")), None)
    if not col_nombre:
        print("  Advertencia: no se encontro 'Nombre de Cooperativa'."); return

    af = filtrar_ruido(af, col=col_nombre)
    af["Cooperativa"] = af[col_nombre].apply(lambda x: resolver(x, mapa, sin_match))
    af = af[af["Cooperativa"].str.strip() != ""]
    af["Cierre"] = af.apply(lambda r:
        f"{str(r.get(col_anio,'')).strip().split('.')[0]}-"
        f"{str(MESES_ES.get(str(r.get(col_mes,'')).strip().lower(),0)).zfill(2)}", axis=1)
    af = af[af["Cierre"].str.match(r"\d{4}-\d{2}")]
    af["#Mes"] = af["Cierre"].str.split("-").str[1].str.lstrip("0")

    rn = {}
    for c in af.columns:
        cn = normalizar(c)
        if c == col_nombre:    rn[c] = "Nombre de Cooperativa"
        elif c == col_mes:     rn[c] = "Name Mes"
        elif c == col_anio:    rn[c] = "Año"
        elif "HOMBRE"  in cn:  rn[c] = "Afiliado Hombres"
        elif "MUJER"   in cn:  rn[c] = "Afiliados Mujeres"
        elif "JURIDIC" in cn:  rn[c] = "Personas Jurídicas"
        elif "MENOR"   in cn:  rn[c] = "Menores Ahorrantes"
        elif "TOTAL"   in cn:  rn[c] = "Total Afiliados "
    af = af.rename(columns=rn).reindex(columns=AFIL_COLS, fill_value="")

    exist  = periodos_afil(RUTA_AFIL)
    nuevos = af[~af["Cierre"].isin(exist)]
    if nuevos.empty:
        print("  Afiliados: todos los periodos ya existen.")
        return
    for p in sorted(nuevos["Cierre"].unique()):
        print(f"  + Afiliados {p}: {len(nuevos[nuevos['Cierre']==p])} cooperativas")
    base = pd.read_csv(RUTA_AFIL, encoding="utf-8-sig", dtype=str) if RUTA_AFIL.exists() else pd.DataFrame(columns=AFIL_COLS)
    r = pd.concat([base, nuevos], ignore_index=True)
    r.to_csv(RUTA_AFIL, index=False, encoding="utf-8-sig")
    print(f"  → Afiliados.csv actualizado ({len(r)} filas)")

# ── MAIN ─────────────────────────────────────────────────
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Actualiza los CSV base con datos del CONSUCOOP")
    parser.add_argument("--eeff",      metavar="RUTA", help="Excel de Estados Financieros")
    parser.add_argument("--afiliados", metavar="RUTA", help="Excel de Afiliados")

    args = parser.parse_args()

    downloads = Path.home() / "Downloads"

    # Buscar EEFF más reciente válido (con o sin número de versión al final, ej. "-5")
    if not args.eeff:

        archivos_eeff = sorted(
            set(downloads.glob("8.-Estados-financieros.xlsx")) |
            set(downloads.glob("8.-Estados-financieros-*.xlsx")),
            key=lambda x: x.stat().st_mtime,
            reverse=True
        )

        if archivos_eeff:
            args.eeff = str(archivos_eeff[0])

    # Buscar Afiliados más reciente (con o sin número de versión al final, ej. "-5")
    if not args.afiliados:

        archivos_afiliados = sorted(
            set(downloads.glob("5.-Detalle-de-afiliados-por-Genero.xlsx")) |
            set(downloads.glob("5.-Detalle-de-afiliados-por-Genero-*.xlsx")),
            key=lambda x: x.stat().st_mtime,
            reverse=True
        )

        if archivos_afiliados:
            args.afiliados = str(archivos_afiliados[0])

    # Validación
    if not args.eeff and not args.afiliados:
        print("⚠ No se encontraron archivos nuevos en Downloads: solo se va a refrescar Indicadores.csv")
        print("  con lo que ya hay en Balance_General.csv / Estado_de_Resultados.csv.\n")



    print("=" * 55)
    mapa = cargar_maestro(RUTA_MAESTRO)
    print(f"Maestro cargado: {len(mapa)} cooperativas")
    print("=" * 55)
    sin_match = set()

    if args.eeff:
        print(f"\nEstados Financieros:\n  {args.eeff}")
        procesar_eeff(Path(args.eeff), mapa, sin_match)
    if args.afiliados:
        print(f"\nAfiliados:\n  {args.afiliados}")
        procesar_afiliados(Path(args.afiliados), mapa, sin_match)

    generar_indicadores()

    print("\n" + "=" * 55)
    if sin_match:
        RUTA_NO_MATCH.write_text("\n".join(sorted(sin_match)), encoding="utf-8")
        print(f"Advertencia: {len(sin_match)} nombre(s) sin match -> {RUTA_NO_MATCH}")
    else:
        print("Todos los nombres coincidieron con el maestro.")
    print("=" * 55 + "\nListo.")