# Código Pilar Martínez y Nicolás Tumaián
# Pilar - 268474 , Nicolás 251664
# Tesis para Ingeniería en electrónica y en telecominicaciones - Universidad ORT

import os
import pandas as pd

# -----------------------------------------------------
# Configuración general
# -----------------------------------------------------
INPUT_FILE = os.path.join("output", "alarmas_agrupadas.xlsx")
SHEET = "alarmas_clasificadas"
OUTPUT_ANALISIS = os.path.join("output", "analisis_alarmas.xlsx")

COLUMNA_ESTACION = "Estacion"
COLUMNA_GRUPO = "grupo"
COLUMNA_ESTADO = "Estado"  # "Activa" / "Clear"

# -----------------------------------------------------
# Cargar datos
# -----------------------------------------------------
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET)
print(f"✅ Archivo cargado: {df.shape[0]} filas, {df.shape[1]} columnas")

for col in [COLUMNA_ESTACION, COLUMNA_GRUPO, COLUMNA_ESTADO]:
    if col not in df.columns:
        raise ValueError(f"No encuentro la columna '{col}' en el archivo.")

# -----------------------------------------------------
# Agregados por estación y grupo
# -----------------------------------------------------
tabla_est_grupo = (
    df.groupby([COLUMNA_ESTACION, COLUMNA_GRUPO])
      .size()
      .unstack(fill_value=0)
      .sort_index()
)
tabla_est_grupo["__TOTAL_ESTACION__"] = tabla_est_grupo.sum(axis=1)

# Totales por grupo
totales_por_grupo = (
    df.groupby(COLUMNA_GRUPO)
      .size()
      .rename("TOTAL_POR_GRUPO")
      .to_frame()
      .sort_values("TOTAL_POR_GRUPO", ascending=False)
)

# Top 3 grupos por estación
top3_por_estacion = (
    df.groupby([COLUMNA_ESTACION, COLUMNA_GRUPO])
      .size()
      .rename("cantidad")
      .reset_index()
      .sort_values([COLUMNA_ESTACION, "cantidad"], ascending=[True, False])
      .groupby(COLUMNA_ESTACION, as_index=False)
      .head(3)
)

# -----------------------------------------------------
# Totales por estación y estado (Activa / Clear)
# -----------------------------------------------------
tabla_est_estado = (
    df.groupby([COLUMNA_ESTACION, COLUMNA_ESTADO])
      .size()
      .unstack(fill_value=0)
      .sort_index()
)

# Asegurar columnas aunque no existan
if "Activa" not in tabla_est_estado.columns:
    tabla_est_estado["Activa"] = 0
if "Clear" not in tabla_est_estado.columns:
    tabla_est_estado["Clear"] = 0

# Forzar a numérico por las dudas (evita dtypes raros)
tabla_est_estado["Activa"] = pd.to_numeric(tabla_est_estado["Activa"], errors="coerce").fillna(0)
tabla_est_estado["Clear"]  = pd.to_numeric(tabla_est_estado["Clear"], errors="coerce").fillna(0)

# Total bruto (informativo)
tabla_est_estado["TOTAL"] = tabla_est_estado["Activa"] + tabla_est_estado["Clear"]

# -----------------------------------------------------
# 📌 Indicador correcto: RATIO CLEAR = Clear / Activa
#    Usamos NaN clásicos (float) para evitar dtype "object"
# -----------------------------------------------------
activa_float = tabla_est_estado["Activa"].astype("float64")
clear_float  = tabla_est_estado["Clear"].astype("float64")

denominador = activa_float.replace(0, float("nan"))
ratio = clear_float / denominador

tabla_est_estado["ratio_clear"] = ratio.astype("float64").round(4).fillna(0.0)

# -----------------------------------------------------
# Porcentajes dentro de cada estación (para otra hoja)
# -----------------------------------------------------
tabla_est_estado_pct = (
    tabla_est_estado.div(tabla_est_estado["TOTAL"], axis=0)
                     .round(4)
)

# -----------------------------------------------------
# Totales generales por estación (independiente del estado)
# -----------------------------------------------------
totales_por_estacion_general = (
    df.groupby(COLUMNA_ESTACION)
      .size()
      .rename("TOTAL_POR_ESTACION")
      .to_frame()
      .sort_values("TOTAL_POR_ESTACION", ascending=False)
)

# -----------------------------------------------------
# Exportar resultados
# -----------------------------------------------------
os.makedirs("output", exist_ok=True)

with pd.ExcelWriter(OUTPUT_ANALISIS, engine="openpyxl") as writer:
    # Resultados previos
    tabla_est_grupo.to_excel(writer, sheet_name="por_estacion_grupo")
    totales_por_grupo.to_excel(writer, sheet_name="totales_grupo")
    top3_por_estacion.to_excel(writer, index=False, sheet_name="top3_por_estacion")

    # Nuevos resultados
    tabla_est_estado.to_excel(writer, sheet_name="por_estacion_estado")
    tabla_est_estado_pct.to_excel(writer, sheet_name="por_estacion_estado_pct")
    totales_por_estacion_general.to_excel(writer, sheet_name="totales_estacion_general")

print(f"📊 Análisis generado en: {OUTPUT_ANALISIS}")
print("👉 Hojas creadas:")
print("   - por_estacion_grupo")
print("   - totales_grupo")
print("   - top3_por_estacion")
print("   - por_estacion_estado (★ incluye ratio_clear)")
print("   - por_estacion_estado_pct")
print("   - totales_estacion_general")
