# Código Pilar Martínez y Nicolás Tumaián
# Pilar - 268474 , Nicolás 251664
# Tesis para Ingeniería en electrónica y en telecominicaciones - Universidad ORT

import os
import re
import unicodedata
import pandas as pd
from rapidfuzz import process, fuzz

# ----------------------------
# Parámetros de entrada/salida
# ----------------------------
INPUT_FILE = os.path.join("data", "alarmas_raw.xlsx")  # Cambia a .csv si hace falta
INPUT_SHEET = 0  # si es Excel y la hoja 0 tiene los datos
COLUMNA_TEXTO = "Alarma"  # nombre exacto de la columna con el texto de la alarma
OUTPUT_FILE = os.path.join("output", "alarmas_agrupadas.xlsx")
MAP_RULES = os.path.join("rules", "reglas_mapeo.csv")

# -------------------------------------------------
# Helpers para normalizar y limpiar texto de alarmas
# -------------------------------------------------
def normalizar(txt: str) -> str:
    if pd.isna(txt):
        return ""
    # a minúsculas
    t = str(txt).lower().strip()
    # quitar acentos
    t = unicodedata.normalize('NFKD', t).encode('ascii', 'ignore').decode('ascii')
    # reemplazos comunes
    t = re.sub(r"\b(\d+)\b", "", t)          # sacar números sueltos (ej: "1", "2")
    t = re.sub(r"\s+", " ", t)               # espacios múltiples
    t = re.sub(r"[-_/]", " ", t)             # separadores a espacio
    return t.strip()

# -----------------------------------
# Cargar catálogo de reglas de mapeo
# -----------------------------------
def cargar_reglas(path_csv: str):
    df_map = pd.read_csv(path_csv)
    # normalizar columna 'patron'
    df_map['patron_norm'] = df_map['patron'].apply(normalizar)
    # el conjunto de grupos válidos (normalizados para match difuso)
    grupos = sorted(set(df_map['grupo'].dropna().astype(str)))
    return df_map, grupos

# ----------------------------------------------------
# Match exacto por substring + fallback con fuzzy match
# ----------------------------------------------------
def asignar_grupo(texto_norm: str, df_map: pd.DataFrame, grupos: list, score_min=85):
    # 1) match por substring de los patrones (más controlable si los patrones son buenos)
    for _, row in df_map.iterrows():
        patron = row['patron_norm']
        grupo = row['grupo']
        if patron and patron in texto_norm:
            return grupo, "substring"

    # 2) fuzzy match contra lista de grupos (por si el texto 'sugiere' algún grupo)
    if texto_norm:
        candidato, score, _ = process.extractOne(
            texto_norm, grupos, scorer=fuzz.token_set_ratio
        )
        if score >= score_min:
            return candidato, f"fuzzy({score})"

    return "sin_clasificar", "none"

def leer_datos(path):
    if path.lower().endswith(".csv"):
        return pd.read_csv(path)
    else:
        return pd.read_excel(path, sheet_name=INPUT_SHEET)

def main():
    os.makedirs("output", exist_ok=True)

    # Cargar dataset
    df = leer_datos(INPUT_FILE)
    if COLUMNA_TEXTO not in df.columns:
        raise ValueError(f"No encuentro la columna '{COLUMNA_TEXTO}' en el archivo de entrada.")

    # Normalizar columna
    df['alarma_norm'] = df[COLUMNA_TEXTO].apply(normalizar)

    # Cargar reglas y grupos
    df_map, grupos = cargar_reglas(MAP_RULES)

    # Asignar grupo
    resultados = df['alarma_norm'].apply(
        lambda t: asignar_grupo(t, df_map, grupos)
    )
    df['grupo'], df['metodo_match'] = zip(*resultados)

    # (Opcional) estadísticas rápidas
    stats = df['grupo'].value_counts(dropna=False).rename_axis('grupo').reset_index(name='cantidad')

    # Guardar
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="alarmas_clasificadas")
        stats.to_excel(writer, index=False, sheet_name="resumen_grupos")

    print(f"Listo ✅ -> {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
