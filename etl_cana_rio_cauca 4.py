"""
ETL INTEGRADO - Impacto ambiental hídrico de la caña de azúcar en el Valle del Cauca
==================================================================================
Nueva funcionalidad:
  - Consolidación en un único Dataset unificado bajo la llave estricta ('Zona', 'Año').
  - Enfoque 100% Anual: Eliminación de suposiciones o expansiones mensuales.
  - Agrupación geo-temporal simétrica real tomada de bases originales.
"""

import pandas as pd
import numpy as np
import os
import sys

# ══════════════════════════════════════════════════════════════════
# CONFIGURACIÓN (Ajusta las rutas locales)
# ══════════════════════════════════════════════════════════════════
RUTA_BASE_AGRICOLA = r"C:\Users\USUARIO\Downloads\Base_Agricola20260610.xlsx"
RUTA_BASE_CALIDAD  = r"C:\Users\USUARIO\Downloads\Calidad_del_agua_del_Rio_Cauca_20260602.xlsx"
RUTA_SALIDA        = r"C:\Users\USUARIO\Downloads\resultado2.7_etl.xlsx"

HOJA_AGRICOLA = None  # Ajustado tras tu verificación
HOJA_CALIDAD  = None   # Ajustado tras tu verificación

AÑO_INICIO = 2019
AÑO_FIN    = 2024

# ── Nombres de columnas originales en tus archivos ────────────────
COL_CULTIVO      = "Desagregación cultivo"
COL_DEPARTAMENTO = "Departamento"
COL_MUNICIPIO    = "Municipio"
COL_AGNO          = "Año"
COL_AREA_SEM     = "Área sembrada (ha)"
COL_AREA_COS     = "Área cosechada (ha)"

COL_FECHA       = "FECHA DE MUESTREO"
COL_ESTACION    = "ESTACIONES"
COL_TEMPERATURA = "TEMPERATURA (°C)"
COL_TURBIEDAD   = "TURBIEDAD (UNT)"
COL_SST         = "SOLIDOS SUSPENDIDOS TOTALES (mg SS/l)"
COL_DBO         = "DEMANDA BIOQUIMICA DE OXIGENO (mg O2/l)"
COL_DQO         = "DEMANDA QUIMICA DE OXIGENO (mg O2/l)"
COL_CONDUCTIV   = "CONDUCTIVIDAD ELÉCTRICA (µS/cm)"
COL_NITRATOS    = "NITRATOS (mg N-NO3/l)"
COL_FOSFATOS    = "FOSFATOS (mg PO4/l)"
COL_CAUDAL      = "CAUDAL (m3/s)"

# ══════════════════════════════════════════════════════════════════
# DICCIONARIOS DE REGIONALIZACIÓN (VALLE DEL CAUCA)
# ══════════════════════════════════════════════════════════════════
ZONAS_MUNICIPIOS = {
    'ALCALA': 'Norte', 'ANSERMANUEVO': 'Norte', 'ARGELIA': 'Norte', 'BOLIVAR': 'Norte', 'BUGALAGRANDE': 'Norte',
    'CAICEDONIA': 'Norte', 'CARTAGO': 'Norte', 'EL CAIRO': 'Norte', 'EL AGUILA': 'Norte', 'EL DOVIO': 'Norte',
    'LA UNION': 'Norte', 'LA VICTORIA': 'Norte', 'OBANDO': 'Norte', 'ROLDANILLO': 'Norte', 'TORO': 'Norte',
    'ULLOA': 'Norte', 'VERSALLES': 'Norte', 'ZARZAL': 'Norte', 'SEVILLA': 'Norte', 'RIOFRIO': 'Norte', 'TRUJILLO': 'Norte',
    'BUGA': 'Centro', 'GUADALAJARA DE BUGA': 'Centro', 'ANDALUCIA': 'Centro', 'CALIMA': 'Centro', 'DARIEN': 'Centro',
    'EL CERRITO': 'Centro', 'GINEBRA': 'Centro', 'GUACARI': 'Centro', 'RESTREPO': 'Centro', 'SAN PEDRO': 'Centro',
    'TULUA': 'Centro', 'YOTOCO': 'Centro',
    'SANTIAGO DE CALI': 'Sur', 'BUENAVENTURA': 'Sur', 'CAICEDO': 'Sur', 'CANDELARIA': 'Sur', 'DAGUA': 'Sur', 'FLORIDA': 'Sur',
    'JAMUNDI': 'Sur', 'LA CUMBRE': 'Sur', 'PALMIRA': 'Sur', 'PRADERA': 'Sur', 'YUMBO': 'Sur', 'VIJES': 'Sur'
}

ZONAS_ESTACIONES = {
    "ANTES RIO TIMBA"       : "Sur", "ANTES SUAREZ"          : "Sur", "ANTES RIO OVEJAS"      : "Sur",
    "ANACARO"               : "Sur", "ANTES INTERCEPTOR SUR" : "Sur", "ANTES INTERCEPTOR"     : "Centro",
    "PUENTE HORMIGUERO"     : "Centro", "PUENTE  HORMIGUERO"    : "Centro", "JUANCHITO"             : "Centro",
    "PASO DEL COMERCIO"     : "Centro", "PASO DE LA BOLSA"      : "Centro", "PASO DE  LA BOLSA"     : "Centro",
    "PASO DE LA TORRE"      : "Centro", "PASO DE  LA TORRE"     : "Centro", "PASO DE LA BALSA"      : "Centro",
    "PASO DE  LA BALSA"     : "Centro", "VIJES"                 : "Centro", "VIJES LA TORRE"        : "Centro",
    "YOTOCO"                : "Centro", "MEDIACANOA"            : "Centro", "PUENTE GUAYABAL"       : "Norte",
    "RIOFRIO"               : "Norte", "PUERTO ISAACS"         : "Norte", "LA VICTORIA"           : "Norte",
    "PUANTE LA VIRGINIA"    : "Norte"
}

def verificar_rutas():
    for ruta, nombre in [(RUTA_BASE_AGRICOLA, "Base Agrícola"), (RUTA_BASE_CALIDAD, "Base Calidad Agua")]:
        if not os.path.isfile(ruta):
            sys.exit(f"[ERROR] No se encontró el archivo '{nombre}':\n  {ruta}")

def leer_excel(ruta, hoja=None):
    ext = os.path.splitext(ruta)[1].lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    df = pd.read_excel(ruta, engine=engine, sheet_name=hoja)
    if isinstance(df, dict):
        df = df[list(df.keys())[0]]
    df.columns = df.columns.str.strip()
    return df

def normalizar_texto(s):
    import unicodedata
    if pd.isna(s): return ""
    s = str(s).upper().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def resolver_columna(df, nombre_config, descripcion):
    mapa = {normalizar_texto(c): c for c in df.columns}
    clave = normalizar_texto(nombre_config)
    if clave in mapa: return mapa[clave]
    for norm, real in mapa.items():
        if clave in norm or norm in clave: return real
    return None

# ─────────────────────────────────────────────────────────────────
# TRANSFORMACIÓN AGRÍCOLA (MODIFICADO: 100% ANUAL REAL)
# ─────────────────────────────────────────────────────────────────
def transformar_agricola(df_raw: pd.DataFrame) -> pd.DataFrame:
    print("\n── TRANSFORMACIÓN BASE AGRÍCOLA (ANUAL REAL) ─────────────────")
    col_cultivo  = resolver_columna(df_raw, COL_CULTIVO, "Desagregación del Cultivo")
    col_depto    = resolver_columna(df_raw, COL_DEPARTAMENTO, "Departamento")
    col_mun      = resolver_columna(df_raw, COL_MUNICIPIO, "Municipio")
    col_agno      = resolver_columna(df_raw, COL_AGNO, "Año")
    col_area_sem = resolver_columna(df_raw, COL_AREA_SEM, "Área Sembrada")
    col_area_cos = resolver_columna(df_raw, COL_AREA_COS, "Área Cosechada")

    df_raw["_mun_norm"] = df_raw[col_mun].apply(normalizar_texto)
    df_raw["_depto_norm"] = df_raw[col_depto].apply(normalizar_texto)
    df_raw["_cultivo_norm"] = df_raw[col_cultivo].apply(normalizar_texto)

    mask = (df_raw["_cultivo_norm"] == normalizar_texto("Caña de azúcar")) & \
           (df_raw["_depto_norm"] == normalizar_texto("Valle del Cauca"))
    df = df_raw[mask].copy()

    for col in [col_area_sem, col_area_cos]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df[col_agno] = pd.to_numeric(df[col_agno], errors="coerce").astype("Int64")

    # Filtrar por el rango de años requerido antes de agrupar
    df = df[df[col_agno].between(AÑO_INICIO, AÑO_FIN)].copy()

    df_agg = df.groupby(["_mun_norm", col_agno], as_index=False).agg({
        col_area_sem: "sum", col_area_cos: "sum"
    })

    df_agg["Zona"] = df_agg["_mun_norm"].map(ZONAS_MUNICIPIOS)
    df_agg.rename(columns={"_mun_norm": "Municipio", col_agno: "Año", col_area_sem: "Area_Sembrada_ha", col_area_cos: "Area_Cosechada_ha"}, inplace=True)
    
    # Se eliminó por completo el bucle 'for mes in range(1, 13)' para mantener la naturaleza puramente anual
    df_final = df_agg[df_agg["Zona"].notna()].copy()
    
    return df_final[["Zona", "Municipio", "Año", "Area_Sembrada_ha", "Area_Cosechada_ha"]]

# ─────────────────────────────────────────────────────────────────
# TRANSFORMACIÓN CALIDAD DEL AGUA (MODIFICADO: AGRUPACIÓN ANUAL)
# ─────────────────────────────────────────────────────────────────
def transformar_calidad(df_raw: pd.DataFrame) -> pd.DataFrame:
    print("\n── TRANSFORMACIÓN BASE CALIDAD DEL AGUA (ANUAL) ──────────────")
    col_fecha    = resolver_columna(df_raw, COL_FECHA, "Fecha de Muestreo")
    col_estacion = resolver_columna(df_raw, COL_ESTACION, "Estaciones")
    
    mapeo_vars = {
        "Temperatura (°C)": COL_TEMPERATURA, "Turbiedad (UNT)": COL_TURBIEDAD,
        "SST (mg SS/l)": COL_SST, "DBO (mg O2/l)": COL_DBO, "DQO (mg O2/l)": COL_DQO,
        "Conductividad (µS/cm)": COL_CONDUCTIV, "Nitratos (mg N-NO3/l)": COL_NITRATOS,
        "Fosfatos (mg PO4/l)": COL_FOSFATOS, "Caudal (m3/s)": COL_CAUDAL
    }
    
    df = pd.DataFrame()
    df["Estación_Original"] = df_raw[col_estacion]
    df["_est_norm"] = df["Estación_Original"].apply(normalizar_texto)
    
    df["_fecha_dt"] = pd.to_datetime(df_raw[col_fecha], dayfirst=True, errors='coerce')
    df["Año"] = df["_fecha_dt"].dt.year
    
    df["Zona"] = df["_est_norm"].map(ZONAS_ESTACIONES)
    df = df[df["Zona"].notna()].copy()
    df = df[df["Año"].between(AÑO_INICIO, AÑO_FIN)].copy()
    df["Año"] = df["Año"].astype("Int64")

    for dest, orig in mapeo_vars.items():
        col_orig = resolver_columna(df_raw, orig, dest)
        if col_orig:
            df[dest] = pd.to_numeric(df_raw.loc[df.index, col_orig], errors="coerce")
            
    # Algoritmos de imputación científica adaptados a nivel anual por Estación
    for v in ["Temperatura (°C)"]:
        if v in df.columns:
            df[v] = df.groupby(["_est_norm", "Año"])[v].transform(lambda x: x.fillna(x.median()))
            
    for v in ["SST (mg SS/l)", "DBO (mg O2/l)", "DQO (mg O2/l)", "Conductividad (µS/cm)", "Nitratos (mg N-NO3/l)", "Fosfatos (mg PO4/l)"]:
        if v in df.columns:
            df[v] = df.groupby("_est_norm")[v].transform(lambda x: x.interpolate(method='linear', limit_direction='both'))
            df[v] = df.groupby(["_est_norm", "Año"])[v].transform(lambda x: x.fillna(x.median()))

    df.drop(columns=["_est_norm", "_fecha_dt"], inplace=True)
    df.rename(columns={"Estación_Original": "Estación"}, inplace=True)
    
    cols_finales = ["Zona", "Estación", "Año"] + [k for k in mapeo_vars.keys() if k in df.columns]
    return df[cols_finales]

# ─────────────────────────────────────────────────────────────────
# CONSOLIDACIÓN INTEGRADA (MODIFICADO: DATASET ÚNICO POR ZONA-AÑO)
# ─────────────────────────────────────────────────────────────────
def consolidar_dataset(df_agr: pd.DataFrame, df_cal: pd.DataFrame) -> pd.DataFrame:
    print("\n── CREANDO DATASET CONSOLIDADO ANUAL ÚNICO ───────────────────")
    
    # 1. Componente Agrícola: Agrupar municipios y sumar áreas exclusivamente por Zona y Año
    df_agr_zona = df_agr.groupby(["Zona", "Año"], as_index=False).agg({
        "Area_Sembrada_ha": "sum",
        "Area_Cosechada_ha": "sum"
    })
    
    # 2. Componente de Calidad: Promediar los muestreos de las estaciones por Zona y Año
    vars_calidad = [c for c in df_cal.columns if c not in ["Zona", "Estación", "Año"]]
    df_cal_zona = df_cal.groupby(["Zona", "Año"], as_index=False)[vars_calidad].mean()
    
    # 3. Cruzar ambos componentes mediante la Llave Compuesta ('Zona', 'Año')
    df_merged = pd.merge(df_agr_zona, df_cal_zona, on=["Zona", "Año"], how="inner")
    
    df_merged = df_merged.sort_values(["Zona", "Año"]).reset_index(drop=True)
    print(f"  ✓ Integración finalizada. Filas (Registros Zona-Año) consolidadas: {len(df_merged)}")
    return df_merged

# ─────────────────────────────────────────────────────────────────
# DIAGNÓSTICO DE COMPLETITUD
# ─────────────────────────────────────────────────────────────────
def diagnostico_completitud(df: pd.DataFrame) -> pd.DataFrame:
    vars_analisis = [c for c in df.columns if c not in ("Estación", "Zona", "Año")]
    registros_totales = len(df)
    filas = []
    for var in vars_analisis:
        n_faltantes  = df[var].isna().sum()
        pct_faltante = round(n_faltantes / registros_totales * 100, 1) if registros_totales else 0
        rec = "Imputado" if pct_faltante == 0 else "Espacio preservado (Protección MNAR)"
        filas.append({
            "Variable": var, "Registros Totales": registros_totales,
            "Faltantes Post-ETL": n_faltantes, "% Faltante Final": pct_faltante, "Nota Técnico": rec
        })
    return pd.DataFrame(filas)

# ─────────────────────────────────────────────────────────────────
# CARGA
# ─────────────────────────────────────────────────────────────────
def cargar(df_agr: pd.DataFrame, df_cal: pd.DataFrame, df_consolidado: pd.DataFrame, df_diag: pd.DataFrame, ruta_salida: str):
    print(f"\n── EXPORTANDO ARTIFACTOS A EXCEL ───────────────────────────\n Ruta: {ruta_salida}")
    os.makedirs(os.path.dirname(os.path.abspath(ruta_salida)), exist_ok=True)

    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        if not df_consolidado.empty:
            df_consolidado.to_excel(writer, sheet_name="Dataset_Consolidado", index=False)
            print("  ✓ Hoja MAESTRA: 'Dataset_Consolidado' (Anual por Zona) creada.")
        if not df_agr.empty:
            df_agr.to_excel(writer, sheet_name="Agricola_Transformada", index=False)
        if not df_cal.empty:
            df_cal.to_excel(writer, sheet_name="Calidad_Agua_Transformada", index=False)
        if not df_diag.empty:
            df_diag.to_excel(writer, sheet_name="Diagnostico_Completitud", index=False)
            
    print("\n[PROCESO COMPLETADO EXITOSAMENTE]")

def main():
    verificar_rutas()
    df_agr_raw = leer_excel(RUTA_BASE_AGRICOLA, HOJA_AGRICOLA)
    df_cal_raw = leer_excel(RUTA_BASE_CALIDAD, HOJA_CALIDAD)

    df_agr = transformar_agricola(df_agr_raw)
    df_cal = transformar_calidad(df_cal_raw)
    
    # Ejecución de la fase de consolidación 100% Anual
    df_consolidado = consolidar_dataset(df_agr, df_cal)
    
    df_diag = diagnostico_completitud(df_cal)
    cargar(df_agr, df_cal, df_consolidado, df_diag, RUTA_SALIDA)

if __name__ == "__main__":
    main()