# Código Pilar Martínez y Nicolás Tumaián
# Pilar - 268474 , Nicolás 251664
# Tesis para Ingeniería en electrónica y en telecominicaciones - Universidad ORT

import os
import pandas as pd

# -----------------------------------------------------
# CONFIG
# -----------------------------------------------------
BASE_DIR = os.path.dirname(os.path.dirname(__file__))

INPUT_AGRUPADAS = os.path.join(BASE_DIR, "output", "alarmas_agrupadas.xlsx")
INPUT_ANALISIS = os.path.join(BASE_DIR, "output", "analisis_alarmas.xlsx")

SHEET_ALARMAS = "alarmas_clasificadas"
SHEET_ESTADO = "por_estacion_estado"

OUTPUT_DIR = os.path.join(BASE_DIR, "output")
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "analisis_avanzado.xlsx")

RATIO_MIN = 0.90
RATIO_MAX = 1.10
COL_ALARMA = "Alarma"   # nombre de la columna con el texto original
COL_GRUPO = "grupo"     # nombre de la columna de tipo de falla

os.makedirs(OUTPUT_DIR, exist_ok=True)

# -----------------------------------------------------
# 1) Cargar dataset base
# -----------------------------------------------------
df = pd.read_excel(INPUT_AGRUPADAS, sheet_name=SHEET_ALARMAS)

df["HoraRecepcionAlarma"] = pd.to_datetime(df["HoraRecepcionAlarma"], errors="coerce")
df = df.dropna(subset=["HoraRecepcionAlarma"])

print(f"✅ Dataset base: {df.shape[0]} filas")

# -----------------------------------------------------
# 2) Cargar ratio_clear por estación (para análisis, no filtramos)
# -----------------------------------------------------
tabla_estado = pd.read_excel(INPUT_ANALISIS, sheet_name=SHEET_ESTADO)
tabla_estado = tabla_estado[["Estacion", "ratio_clear"]]

estaciones_sospechosas = tabla_estado[
    ~tabla_estado["ratio_clear"].between(RATIO_MIN, RATIO_MAX)
].copy()

print(f"➡ Estaciones totales: {tabla_estado.shape[0]}")
print(f"   Estaciones sospechosas (fuera de {RATIO_MIN}-{RATIO_MAX}): {estaciones_sospechosas.shape[0]}")

# -----------------------------------------------------
# 2b) Detalle por entrada en estaciones sospechosas
# -----------------------------------------------------
print("\n🔍 Analizando entradas en estaciones sospechosas...")

df_sospechosas = df[df["Estacion"].isin(estaciones_sospechosas["Estacion"])].copy()
df_sospechosas["Estado_lower"] = df_sospechosas["Estado"].astype(str).str.lower()

detalle_entrada = (
    df_sospechosas
    .groupby(["Estacion", "Entrada", "Estado_lower"])
    .size()
    .unstack(fill_value=0)
)

# Normalizamos nombres de columnas y aseguramos activas/clears
detalle_entrada = detalle_entrada.rename(columns={"activa": "activas", "clear": "clears"})
for c in ["activas", "clears"]:
    if c not in detalle_entrada.columns:
        detalle_entrada[c] = 0

# Calculamos ratio_clear por entrada (solo donde activas > 0)
detalle_entrada["ratio_clear"] = (
    detalle_entrada["clears"] / detalle_entrada["activas"].replace(0, pd.NA)
)

# Comentario correcto según combinación de activas/clears
def comentar_fila(row):
    a = row["activas"]
    c = row["clears"]
    r = row["ratio_clear"]

    if a == 0 and c == 0:
        return "sin eventos"
    if a == 0 and c > 0:
        return "solo clears (sin activas)"
    if a > 0 and c == 0:
        return "solo activas (sin clears)"
    # acá a > 0 y c > 0
    if pd.isna(r):
        return "ratio no calculable"
    if r < RATIO_MIN or r > RATIO_MAX:
        return "desbalance ⚠"
    return "OK"

detalle_entrada["comentario"] = detalle_entrada.apply(comentar_fila, axis=1)
detalle_entrada = detalle_entrada.reset_index()

print("➡ Detalle por entrada generado:", detalle_entrada.shape[0], "filas")

# -----------------------------------------------------
# 3) Preparar dataset para MTTR 
# -----------------------------------------------------
df_mttr = df.copy()

df_mttr["Alarma_norm_match"] = (
    df_mttr[COL_ALARMA]
    .astype(str)
    .str.strip()
    .str.lower()
)

df_mttr = df_mttr.sort_values(
    ["Estacion", "Alarma_norm_match", "HoraRecepcionAlarma"]
)

# -----------------------------------------------------
# 4) Encontrar pares Activa -> Clear consecutivos
# -----------------------------------------------------
df_mttr["Estado_lower"] = df_mttr["Estado"].astype(str).str.lower()

df_mttr["Estado_next"] = (
    df_mttr
    .groupby(["Estacion", "Alarma_norm_match"])["Estado_lower"]
    .shift(-1)
)

df_mttr["HoraRecepcionAlarma_next"] = (
    df_mttr
    .groupby(["Estacion", "Alarma_norm_match"])["HoraRecepcionAlarma"]
    .shift(-1)
)

mask_pares = (
    (df_mttr["Estado_lower"] == "activa") &
    (df_mttr["Estado_next"] == "clear")
)

recuperacion = df_mttr.loc[mask_pares].copy()

print(f"\n➡ Pares Activa→Clear encontrados: {recuperacion.shape[0]} filas")

# -----------------------------------------------------
# 5) Calcular tiempo de recuperación (minutos)
# -----------------------------------------------------
recuperacion["tiempo_min"] = (
    (recuperacion["HoraRecepcionAlarma_next"] - recuperacion["HoraRecepcionAlarma"])
    .dt.total_seconds() / 60
)

antes_filtro = recuperacion.shape[0]
recuperacion = recuperacion[recuperacion["tiempo_min"] >= 0]

print(f"➡ Después de filtrar tiempos negativos: {antes_filtro} → {recuperacion.shape[0]} filas")

# -----------------------------------------------------
# 6) MTTR por estación
# -----------------------------------------------------
mttr_estacion = (
    recuperacion.groupby("Estacion")["tiempo_min"]
    .mean()
    .reset_index(name="MTTR_min")
    .sort_values("MTTR_min")
)

print(f"➡ MTTR por estación: {mttr_estacion.shape[0]} filas")

# -----------------------------------------------------
# 7) MTTR por estación y grupo
# -----------------------------------------------------
if COL_GRUPO in recuperacion.columns:
    mttr_estacion_grupo = (
        recuperacion.groupby(["Estacion", COL_GRUPO])["tiempo_min"]
        .mean()
        .reset_index(name="MTTR_min")
    )
    mttr_estacion_grupo = mttr_estacion_grupo.sort_values(["Estacion", "MTTR_min"])

    print(f"➡ MTTR por estación y grupo: {mttr_estacion_grupo.shape[0]} filas")

    pivot_num = mttr_estacion_grupo.pivot(
        index="Estacion",
        columns=COL_GRUPO,
        values="MTTR_min"
    )

    pivot_fmt = pivot_num.round(2).astype(object)
    pivot_fmt = pivot_fmt.where(~pivot_fmt.isna(), "Sin tiempo")

else:
    print(f"⚠ No se encontró columna '{COL_GRUPO}' en 'recuperacion'.")
    mttr_estacion_grupo = pd.DataFrame(columns=["Estacion", COL_GRUPO, "MTTR_min"])
    pivot_fmt = pd.DataFrame()

# -----------------------------------------------------
# 8) MTTR por tipo de alarma
# -----------------------------------------------------
cols_rec = recuperacion.columns.tolist()
if COL_ALARMA in cols_rec:
    col_alarma_group = COL_ALARMA
else:
    col_alarma_group = next((c for c in cols_rec if c.lower().startswith("alarma")), None)

if col_alarma_group is not None:
    mttr_alarma = (
        recuperacion.groupby(col_alarma_group)["tiempo_min"]
        .mean()
        .reset_index(name="MTTR_min")
        .sort_values("MTTR_min")
    )
    print(f"➡ MTTR por alarma: {mttr_alarma.shape[0]} filas")
else:
    print("⚠ No se encontró columna de alarma en 'recuperacion' para MTTR por tipo.")
    mttr_alarma = pd.DataFrame(columns=[COL_ALARMA, "MTTR_min"])

# -----------------------------------------------------
# 9) Preparar dataframes de export con timestamps como texto ISO
# -----------------------------------------------------
df_mttr_export = df_mttr.copy()
recuperacion_export = recuperacion.copy()

for df_tmp, cols in [
    (df_mttr_export, ["HoraRecepcionAlarma", "HoraRecepcionAlarma_next"]),
    (recuperacion_export, ["HoraRecepcionAlarma", "HoraRecepcionAlarma_next"]),
]:
    for col in cols:
        if col in df_tmp.columns:
            df_tmp[col] = pd.to_datetime(df_tmp[col], errors="coerce").dt.strftime(
                "%Y-%m-%d %H:%M:%S"
            )

# -----------------------------------------------------
# 10) Exportar TODO a un solo Excel
# -----------------------------------------------------
with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
    tabla_estado.to_excel(writer, sheet_name="ratio_clear_estacion", index=False)
    estaciones_sospechosas.to_excel(writer, sheet_name="estaciones_sospechosas", index=False)
    detalle_entrada.to_excel(writer, sheet_name="detalle_por_entrada", index=False)
    df_mttr_export.to_excel(writer, sheet_name="dataset_ordenado", index=False)
    recuperacion_export.to_excel(writer, sheet_name="pares_activa_clear", index=False)
    mttr_estacion.to_excel(writer, sheet_name="mttr_por_estacion", index=False)
    mttr_estacion_grupo.to_excel(writer, sheet_name="mttr_estacion_grupo", index=False)
    pivot_fmt.to_excel(writer, sheet_name="mttr_estacion_grupo_pivot")
    mttr_alarma.to_excel(writer, sheet_name="mttr_por_alarma", index=False)

print("\n✅ Análisis avanzado generado en:")
print("   ", OUTPUT_XLSX)














