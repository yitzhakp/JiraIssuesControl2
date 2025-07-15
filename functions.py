from io import BytesIO
import pandas as pd
import json

def load_json(filename):
    try:
        with open(filename, 'r') as file:
            data = json.load(file)
        return data
    except FileNotFoundError:
        print("Error: 'data.json' not found.")
    except json.JSONDecodeError:
        print("Error: Invalid JSON format in 'data.json'.")


def evaluar_dia(horas):
    if horas == 0:
        return "❌ No estimó"
    elif horas < 8:
        return "⚠️ Incumple estimativo"
    elif horas == 8:
        return "✅ Cumple estimativo"
    else:
        return "🚀 Excede estimativo"
    
def evaluar_semana(row):
    if row['Time Spent (hours)'] == 0:
        return "❌ No estimó en la semana"
    elif row['Time Spent (hours)'] < row['Horas esperadas']:
        return "⚠️ Incumple estimativo semanal"
    elif row['Time Spent (hours)'] == row['Horas esperadas']:
        return "✅ Cumple estimativo semanal"
    else:
        return "🚀 Excede estimativo semanal"
    
def to_excel(df, nombre_hoja='Sheet1'):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=nombre_hoja)
        worksheet = writer.sheets[nombre_hoja]
        for idx, col in enumerate(df.columns):
            column_len = max(min(df[col].astype(str).map(len).max(), 50), len(col))
            worksheet.set_column(idx, idx, column_len + 2)
    processed_data = output.getvalue()
    return processed_data

def analizar_comentario(comentario):
    comentario = str(comentario).lower()
    coincidencias = set()

    clasificaciones = load_json('./clasificaciones.json')
    for categoria, palabras in clasificaciones.items():
        for palabra in palabras:
            if palabra in comentario:
                coincidencias.add(categoria)

    # Elegimos la primera categoría encontrada para clasificar
    clasificacion = list(coincidencias)[0] if coincidencias else "No clasificado"

    # Si hay más de una categoría, marcar como supervisado
    supervisado = "🚨" if len(coincidencias) > 1 else "✅"

    return pd.Series([clasificacion, supervisado])