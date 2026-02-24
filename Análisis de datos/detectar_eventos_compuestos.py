# Código Pilar Martínez y Nicolás Tumaián
# Pilar - 268474 , Nicolás 251664
# Tesis para Ingeniería en electrónica y en telecominicaciones - Universidad ORT

import os
import pandas as pd

# -----------------------------------------------------
# CONFIG GENERAL
# -----------------------------------------------------
BASE_DIR = os.path.dirname(os.path.dirname(__file__))

INPUT_AGRUPADAS = os.path.join(BASE_DIR, "output", "alarmas_agrupadas.xlsx")
SHEET_ALARMAS = "alarmas_clasificadas"

OUTPUT_DIR = os.path.join(BASE_DIR, "output")
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "eventos_compuestos.xlsx")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# -----------------------------------------------------
# REGLAS DE EVENTOS COMPUESTOS 
# -----------------------------------------------------
REGLAS_EVENTOS = [
    {
        "nombre": "corte_red_grupo_ok",
        "descripcion": "Corte de red donde el grupo entra y sostiene la carga",
        "ventana_min": 5,
        "requiere": ["Batería en descarga", "Grupo en marcha"],
        "requiere_al_menos_uno": [
            "Alta tensión de red", "Baja tensión de red",
            "Tensión anormal", "Falla de red"
        ],
        "prohibe": [
            "Dispositivo sin conectividad", "Baja tensión de baterías",
            "Bajo tiempo respaldo de baterías", "Falla de grupo motor",
            "Falla de predisposición de grupo motor",
        ],
    },
    {
        "nombre": "corte_red_grupo_no_arranca",
        "descripcion": "Corte de red donde el grupo no entra y el sitio cae por baterías",
        "ventana_min": 5.1,
        "requiere": ["Batería en descarga", "Baja tensión de baterías",
                      "Bajo tiempo respaldo de baterías"],
        "requiere_al_menos_uno": [
            "Alta tensión de red", "Baja tensión de red",
            "Tensión anormal", "Falla de red"
        ],
        "prohibe": ["Grupo en marcha"],
    },
    {
        "nombre": "falla_rectificador_cadena_dc",
        "descripcion": "Falla de rectificador / alimentación DC que lleva a uso de baterías",
        "ventana_min": 2,
        "requiere": ["Falla rectificador", "Falla alimentación rectificador", "Batería en descarga"],
        "requiere_al_menos_uno": ["Baja tensión de baterías", "Bajo tiempo respaldo de baterías"],
        "prohibe": [],
    },
    {
        "nombre": "problema_ac_sobretemperatura",
        "descripcion": "Alta temperatura relacionada a falla de acondicionador",
        "ventana_min": 20,
        "requiere": ["Falla acondicionador", "Falla alimentación acondicionador"],
        "requiere_al_menos_uno": ["Alta temperatura"],
        "prohibe": [],
    },
    {
        "nombre": "problema_ac_frio_excesivo",
        "descripcion": "Baja temperatura relacionada a falla de acondicionador",
        "ventana_min": 20,
        "requiere": ["Falla acondicionador", "Falla alimentación acondicionador"],
        "requiere_al_menos_uno": ["Baja temperatura"],
        "prohibe": [],
    },
    {
        "nombre": "problema_tablero",
        "descripcion": "Falla de tablero asociada a fallas de alimentación",
        "ventana_min": 2,
        "requiere": ["Falla tablero"],
        "requiere_al_menos_uno": [
            "Falla rectificador", "Falla UPS",
            "Falla alimentación acondicionador", "Falla alimentación rectificador",
        ],
        "prohibe": [],
    },
    {
        "nombre": "problema_balizaje_por_energia",
        "descripcion": "Falla de balizaje asociada a problemas de energía",
        "ventana_min": 5,
        "requiere": ["Falla balizaje"],
        "requiere_al_menos_uno": [
            "Falla UPS", "Falla tablero",
            "Alta tensión de red", "Baja tensión de red", "Tensión anormal","Baja tensión de baterías",
        ],
        "prohibe": [],
    },
    {
        "nombre": "presencia_en_evento_critico",
        "descripcion": "Alerta de presencia en contexto de fallo crítico",
        "ventana_min": 120,
        "requiere": ["Alerta presencia"],
        "requiere_al_menos_uno": [
            "Falla tablero", "Falla rectificador", "Falla UPS",
            "Grupo en marcha", "Batería en descarga","Falla de red",
        ],
        "prohibe": [],
    },
        {
        "nombre": "bateria_sin_grupo_sin_alarma_bateria_mala",
        "descripcion": "Batería en descarga sin grupo, sin baja tensión ni bajo tiempo respaldo (batería no muestra síntomas de estar mal en la ventana)",
        "ventana_min": 90,  
        "requiere": ["Batería en descarga"],
        "requiere_al_menos_uno": ["Alta tensión de red", "Baja tensión de red", "Tensión anormal", "Falla de red"],
        "prohibe": [
            "Grupo en marcha",
            "Baja tensión de baterías",
            "Bajo tiempo respaldo de baterías",
            "Dispositivo sin conectividad",
            "Falla de grupo motor",
            "Falla de predisposición de grupo motor",
        ],
    },
        {
        "nombre": "bateria_muy_mal_en_10min",
        "descripcion": "Batería en descarga y aparecen síntomas de batería mala muy rápido (<=10 min)",
        "ventana_min": 10,
        "requiere": ["Batería en descarga"],
        "requiere_al_menos_uno": ["Baja tensión de baterías", "Bajo tiempo respaldo de baterías"],
        "prohibe": [],
    },
{
        "nombre": "bateria_ok_al_menos_60min",
        "descripcion": "Batería en descarga sin síntomas de batería mala ni caída (>=60 min)",
        "ventana_min": 60,
        "requiere": ["Batería en descarga"],
        "requiere_al_menos_uno": [],  
        "prohibe": ["Baja tensión de baterías", "Bajo tiempo respaldo de baterías", "Dispositivo sin conectividad"],
    },
{
        "nombre": "bateria_muy_bien_al_menos_180min",
        "descripcion": "Batería en descarga sin síntomas de batería mala ni caída (>=180 min)",
        "ventana_min": 180,
        "requiere": ["Batería en descarga"],
        "requiere_al_menos_uno": [],
        "prohibe": ["Baja tensión de baterías", "Bajo tiempo respaldo de baterías", "Dispositivo sin conectividad"],
    },
]

# -----------------------------------------------------
# AUX
# -----------------------------------------------------
def regla_cumple(grupos_en_ventana: set, regla: dict) -> bool:
    if any(g not in grupos_en_ventana for g in regla.get("requiere", [])):
        return False
    if regla.get("requiere_al_menos_uno"):
        if not any(g in grupos_en_ventana for g in regla["requiere_al_menos_uno"]):
            return False
    if any(g in grupos_en_ventana for g in regla.get("prohibe", [])):
        return False
    return True


# -----------------------------------------------------
# CARGA DE DATOS
# -----------------------------------------------------
print("📥 Cargando alarmas agrupadas...")
df = pd.read_excel(INPUT_AGRUPADAS, sheet_name=SHEET_ALARMAS)

df["HoraRecepcionAlarma"] = pd.to_datetime(df["HoraRecepcionAlarma"], errors="coerce")
df = df.dropna(subset=["HoraRecepcionAlarma"])

df = df.sort_values(["Estacion", "HoraRecepcionAlarma"]).reset_index(drop=True)

print(f"✅ Dataset cargado: {df.shape[0]} filas, {df['Estacion'].nunique()} estaciones.\n")

# -----------------------------------------------------
# DETECCIÓN DE EVENTOS COMPUESTOS
# -----------------------------------------------------
eventos = []

for estacion, df_est in df.groupby("Estacion"):
    df_est = df_est.reset_index(drop=True)
    n = len(df_est)

    tiempos = df_est["HoraRecepcionAlarma"]

    for regla in REGLAS_EVENTOS:
        ventana = pd.Timedelta(minutes=regla["ventana_min"])

        end_idx = 0
        for start_idx in range(n):
            t0 = tiempos.iloc[start_idx]
            tmax = t0 + ventana

            while end_idx < n and tiempos.iloc[end_idx] <= tmax:
                end_idx += 1

            sub = df_est.iloc[start_idx:end_idx]
            if sub.empty:
                continue

            grupos_set = set(sub["grupo"].dropna().unique())

            if regla_cumple(grupos_set, regla):
                eventos.append({
                    "Estacion": estacion,
                    "nombre_evento": regla["nombre"],
                    "descripcion_evento": regla["descripcion"],
                    "inicio_evento": sub["HoraRecepcionAlarma"].min(),
                    "fin_evento": sub["HoraRecepcionAlarma"].max(),
                    "ventana_min": regla["ventana_min"],
                    "grupos_en_ventana": " | ".join(sorted(grupos_set)),
                    "cantidad_alarmas_en_ventana": len(sub),
                    "cantidad_grupos_distintos": len(grupos_set),
                })

    print(f"⬇ Estación {estacion}: {n} alarmas procesadas.")

# Convertimos a DataFrame
df_eventos = pd.DataFrame(eventos)
if df_eventos.empty:
    print("\n⚠ No se detectaron eventos compuestos.")
    exit()

df_eventos["inicio_evento"] = df_eventos["inicio_evento"].dt.strftime("%Y-%m-%d %H:%M:%S")
df_eventos["fin_evento"] = df_eventos["fin_evento"].dt.strftime("%Y-%m-%d %H:%M:%S")

df_eventos = df_eventos.sort_values(["Estacion", "inicio_evento"]).reset_index(drop=True)

# -----------------------------------------------------
# RESÚMENES ADICIONALES
# -----------------------------------------------------

# 1) Cantidad de eventos por estación
resumen_estacion = (
    df_eventos.groupby("Estacion")["nombre_evento"]
    .count()
    .reset_index(name="cantidad_eventos")
    .sort_values("cantidad_eventos", ascending=False)
)

# 2) Cantidad por tipo de evento
resumen_tipo = (
    df_eventos.groupby("nombre_evento")["Estacion"]
    .count()
    .reset_index(name="cantidad_apariciones")
    .sort_values("cantidad_apariciones", ascending=False)
)

# 3) Matriz Estación × Tipo
resumen_matrix = pd.crosstab(df_eventos["Estacion"], df_eventos["nombre_evento"])

# -----------------------------------------------------
# EXPORTAR
# -----------------------------------------------------
with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
    df_eventos.to_excel(writer, sheet_name="eventos_compuestos", index=False)
    resumen_estacion.to_excel(writer, sheet_name="resumen_por_estacion", index=False)
    resumen_tipo.to_excel(writer, sheet_name="resumen_por_evento", index=False)
    resumen_matrix.to_excel(writer, sheet_name="matriz_estacion_evento")

print("\n✅ Detección de eventos compuestos completada.")
print(f"   Archivo generado: {OUTPUT_XLSX}")
print(f"   Total eventos detectados: {df_eventos.shape[0]}")
