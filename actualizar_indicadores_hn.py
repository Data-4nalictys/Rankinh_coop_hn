"""
actualizar_indicadores_hn.py — Lee Balance_General.csv y Estado_de_Resultados.csv
y (re)genera Indicadores.csv: el resumen de indicadores financieros publicable,
con la calificación A-E de CONSUCOOP para cada cooperativa y periodo.

USO:
    python actualizar_indicadores_hn.py

No necesita un Excel de entrada: se corre DESPUÉS de actualizar_ranking_hn.py,
porque toma los dos CSV que ese script ya deja actualizados.

INDICADORES QUE CALCULA (parámetros oficiales del Excel de CONSUCOOP compartido):
    SOLVENCIA         ~ 3.1 Índice de Capital Institucional        meta >= 8.5%
    MOROSIDAD         = 2.2 Índice de Morosidad                    meta <= 5%
    COBERTURA MORA    = 2.1 Suficiencia de Provisión para Cartera  meta >= 100%
    LIQUIDEZ          ~ 4.1 Cobertura oblig. depositarias (M.N.)   meta >= 9%
    ROA               ~ 6.1 Rentabilidad s/ Activos Netos Promedio meta >= 0.50%
    ROE               referencial, sin meta oficial en la lista de CONSUCOOP
    AUTOSUFICIENCIA   = 5.1 Autosuficiencia Operativa              meta >= 110%

Nota: el Excel de CONSUCOOP solo trae el corte final de cada meta (cumple / no
cumple), no las 4 bandas A-E intermedias. Esas se estiman proporcionalmente
alrededor de la meta (ver RANGOS_ES más abajo) — igual que en indicadores.html.
Si consigues la tabla oficial de rangos por letra, ajusta RANGOS_ES.
"""

from pathlib import Path
import numpy as np
import pandas as pd

# ── CONFIGURACION ────────────────────────────────────────

BASE = Path(r"C:\Users\RONYL\OneDrive - Cooperativa de Ahorro y Credito Sonaguera Limitada\Documentos\GitHub\ranking_coop_hn\data")
RUTA_BG  = BASE / "Balance_General.csv"
RUTA_ER  = BASE / "Estado_de_Resultados.csv"
RUTA_IND = BASE / "Indicadores.csv"

# ── METAS OFICIALES (valor, dirección) ────────────────────
# dirección "desc" = más alto es mejor · "asc" = más bajo es mejor
METAS = {
    "SOLVENCIA":        (8.5, "desc"),
    "MOROSIDAD":        (5.0, "asc"),
    "COBERTURA MORA":   (100.0, "desc"),
    "LIQUIDEZ":         (9.0, "desc"),
    "ROA":              (0.5, "desc"),
    "ROE":              (10.0, "desc"),   # referencial, sin parámetro CONSUCOOP
    "AUTOSUFICIENCIA":  (110.0, "desc"),
}

# ── COLUMNAS DE SALIDA ────────────────────────────────────
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

# ── FUNCIONES CORE ────────────────────────────────────────

def a_num(col):
    """Convierte una columna string (con comas de miles) a numérico."""
    return pd.to_numeric(
        col.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce"
    ).fillna(0.0)


def calificar(valor, meta, direccion):
    """Devuelve la letra A-E según qué tan lejos está el valor de la meta."""
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


def cargar_balance():
    print(f"  Leyendo '{RUTA_BG.name}'...")
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
                  f"Revisa BG_COLS en actualizar_ranking_hn.py si es una columna de mora/depósitos.")
            bg[c] = 0

    # Si Atrasados/Vencidos/Demanda Judicial vienen las 3 en blanco (no reportadas
    # todavía, como pasó en 2026-06), la cartera en riesgo debe quedar SIN DATO,
    # no en 0% — de lo contrario Morosidad se ve falsamente "sana".
    mora_cols = ["Atrasados hn", "Vencidos hn", "Demanda Judicial hn"]
    sin_mora_reportada = bg[mora_cols].apply(
        lambda col: col.isna() | col.astype(str).str.strip().eq("")
    ).all(axis=1)

    for c in campos:
        bg[c] = a_num(bg[c])

    bg["CARTERA EN RIESGO"] = bg["Atrasados hn"] + bg["Vencidos hn"] + bg["Demanda Judicial hn"]
    bg.loc[sin_mora_reportada, "CARTERA EN RIESGO"] = np.nan
    bg["EXIGIBILIDADES TOTALES"] = bg["EXIGIBILIDADES INMEDIATAS"] + bg["EXIGIBILIDADES X DEPOSITOS"]
    return bg


def cargar_resultados():
    print(f"  Leyendo '{RUTA_ER.name}'...")
    er = pd.read_csv(RUTA_ER, encoding="utf-8-sig", dtype=str)
    er.columns = er.columns.str.strip()
    for c in ("INGRESOS", "EGRESOS", "EXCEDENTES"):
        er[c] = a_num(er[c]) if c in er.columns else 0
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

    for ind, (meta, direccion) in METAS.items():
        df[f"RANGO {ind}"] = df[ind].apply(lambda v: calificar(v, meta, direccion))

    for ind in METAS:
        df[ind] = df[ind].round(2)

    return df.reindex(columns=IND_COLS)


# ── MAIN ─────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 55)
    print("Calculando indicadores financieros (CONSUCOOP)")
    print("=" * 55)

    if not RUTA_BG.exists() or not RUTA_ER.exists():
        raise SystemExit(
            "❌ Faltan Balance_General.csv y/o Estado_de_Resultados.csv. "
            "Corre primero actualizar_ranking_hn.py."
        )

    bg = cargar_balance()
    er = cargar_resultados()

    print("\n  Calculando SOLVENCIA, MOROSIDAD, COBERTURA MORA, LIQUIDEZ, ROA, ROE, AUTOSUFICIENCIA...")
    ind = calcular_indicadores(bg, er)

    ind.to_csv(RUTA_IND, index=False, encoding="utf-8-sig")
    print(f"  → {RUTA_IND.name} generado ({len(ind)} filas)")

    periodos_sin_mora = ind[ind["MOROSIDAD"].isna()][["COOPERATIVA", "MES#", "AÑO"]]
    if not periodos_sin_mora.empty:
        print(f"\n  Advertencia: {len(periodos_sin_mora)} fila(s) sin PRESTAMOS o sin desglose de mora "
              f"(Atrasados/Vencidos/Demanda Judicial) → MOROSIDAD queda vacío en vez de 0%.")

    print("\n" + "=" * 55 + "\nListo.")
