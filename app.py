import streamlit as st
import pandas as pd
import re
from datetime import timedelta
from io import BytesIO
import xlsxwriter

from functions import evaluar_dia, evaluar_semana, to_excel, load_json, analizar_comentario


st.set_page_config(page_title="Reporte y Disponibilidad", layout="wide")

# Sidebar
st.sidebar.title("📁 Formularios")
seccion = st.sidebar.selectbox(
    "Selecciona una sección:",
    [
        "📊 Reporte de estimaciones por usuario",
        "🧾 Consulta Disponibilidad",
        "📌 Reporte de gestión"
    ]
)

# -------------------------------
# SECCIÓN 1: REPORTE DE ESTIMACIONES POR USUARIO
# -------------------------------
if seccion == "📊 Reporte de estimaciones por usuario":
    st.markdown("<h1 style='color:#0030f6'>📊 Reporte de estimaciones por usuario</h1>", unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Sube tu archivo Excel (ej: worklogs_2025-03-29_2025-04-29)",
        type=["xlsx"],
        key="reporte"
    )

    if uploaded_file is not None:
        try:
            filename = uploaded_file.name
            match = re.search(r'worklogs?_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})', filename)
            if not match:
                st.error("⚠️ El nombre del archivo no contiene las fechas esperadas.")
            else:
                start_str, end_str = match.groups()
                start_date = pd.to_datetime(start_str).date()
                end_date = pd.to_datetime(end_str).date()

                business_days = pd.bdate_range(start=start_date, end=end_date).date

                df = pd.read_excel(uploaded_file)
                required_columns = ['Issue Key', 'Time Spent', 'Time Spent (seconds)', 'Author', 'Start Date', 'Project Key']
                df = df[required_columns]
                df['Start Date'] = pd.to_datetime(df['Start Date']).dt.date
                df['Time Spent (hours)'] = df['Time Spent (seconds)'] / 3600

                df_grouped = df.groupby(['Author', 'Start Date'], as_index=False)['Time Spent (hours)'].sum()
                authors = df_grouped['Author'].unique()
                complete_index = pd.MultiIndex.from_product([authors, business_days], names=['Author', 'Start Date'])
                df_complete = pd.DataFrame(index=complete_index).reset_index()
                df_final = pd.merge(df_complete, df_grouped, on=['Author', 'Start Date'], how='left')
                df_final['Time Spent (hours)'] = df_final['Time Spent (hours)'].fillna(0)

                

                df_final['Evaluación'] = df_final['Time Spent (hours)'].apply(evaluar_dia)

                st.markdown("<h4 style='color:#f15a30'>Filtrar por autor</h4>", unsafe_allow_html=True)
                selected_author = st.selectbox("Selecciona un autor", options=["Todos"] + list(authors), key="autor")

                if selected_author != "Todos":
                    df_filtered = df_final[df_final['Author'] == selected_author]
                else:
                    df_filtered = df_final

                st.markdown("<h4 style='color:#f15a30'>🗓️ Estimaciones diarias</h4>", unsafe_allow_html=True)
                st.dataframe(df_filtered.sort_values(by=['Author', 'Start Date']))

                df_final['Start Date'] = pd.to_datetime(df_final['Start Date'])
                fechas_ordenadas = sorted(business_days)
                semanas = []
                i = 0
                while i < len(fechas_ordenadas):
                    lunes = fechas_ordenadas[i]
                    viernes = lunes
                    while i < len(fechas_ordenadas) and fechas_ordenadas[i].weekday() <= 4 and (fechas_ordenadas[i] - lunes).days < 5:
                        viernes = fechas_ordenadas[i]
                        i += 1
                    semanas.append((lunes, viernes))

                semana_map = {}
                for idx, (inicio, fin) in enumerate(semanas, start=1):
                    etiqueta = f'W{idx}'
                    for f in pd.bdate_range(start=inicio, end=fin):
                        semana_map[f.date()] = (etiqueta, (inicio, fin))

                df_final['Semana Etiqueta'] = df_final['Start Date'].dt.date.map(lambda d: semana_map[d][0] if d in semana_map else None)

                dias_laborales_por_semana = (
                    pd.Series([semana_map[d][0] for d in business_days])
                    .value_counts()
                    .to_dict()
                )

                df_semanal = df_final.groupby(['Author', 'Semana Etiqueta'], as_index=False)['Time Spent (hours)'].sum()
                df_semanal['Días laborales'] = df_semanal['Semana Etiqueta'].map(dias_laborales_por_semana)
                df_semanal['Horas esperadas'] = df_semanal['Días laborales'] * 8

                

                df_semanal['Evaluación Semanal'] = df_semanal.apply(evaluar_semana, axis=1)

                st.markdown("<h4 style='color:#f15a30'>📆 Estimaciones semanales</h4>", unsafe_allow_html=True)
                if selected_author != "Todos":
                    df_semanal = df_semanal[df_semanal['Author'] == selected_author]

                st.dataframe(df_semanal.sort_values(by=['Author', 'Semana Etiqueta']))

        except Exception as e:
            st.error(f"❌ Error al procesar el archivo: {e}")


# -------------------------------
# SECCIÓN 2: CONSULTA DISPONIBILIDAD
# -------------------------------
elif seccion == "🧾 Consulta Disponibilidad":
    
    st.markdown("<h1 style='color:#007200'>🧾 Consulta Disponibilidad</h1>", unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Sube hasta 6 archivos Excel para verificar disponibilidad",
        type=["xlsx"],
        accept_multiple_files=True,
        key="disponibilidad"
    )

    if uploaded_files:
        if len(uploaded_files) > 6:
            st.error("⚠️ Solo se permiten hasta 6 archivos.")
        else:
            try:
                dataframes = []
                for file in uploaded_files:
                    df_temp = pd.read_excel(file)
                    base_name = file.name.split('.')[0]
                    match = re.search(r'Tracking_([A-Za-z]+)(\d{4})', base_name)
                    if match:
                        mes = match.group(1).capitalize()
                        anio = match.group(2)
                        periodo = f"{mes} {anio}"
                    else:
                        periodo = "Desconocido"
                    df_temp['Periodo'] = periodo
                    dataframes.append(df_temp)

                df_merged = pd.concat(dataframes, ignore_index=True)

                if 'Time Spent' in df_merged.columns and 'Time spent' in df_merged.columns:
                    df_merged['Time Spent Final'] = df_merged['Time Spent'].combine_first(df_merged['Time spent'])
                elif 'Time Spent' in df_merged.columns:
                    df_merged['Time Spent Final'] = df_merged['Time Spent']
                elif 'Time spent' in df_merged.columns:
                    df_merged['Time Spent Final'] = df_merged['Time spent']
                else:
                    df_merged['Time Spent Final'] = None

                palabras_clave = [
                    "ruta de aprendizaje", "curso", "espera de asignaciones",
                    "sin asignaciones", "disponibilidad", "capacitación"
                ]
                pattern = '|'.join([re.escape(p) for p in palabras_clave])
                mask = df_merged['Comment'].astype(str).str.lower().str.contains(pattern)
                df_disponibilidad = df_merged[mask]

                columnas_a_mostrar = ['Author', 'Comment', 'Time Spent Final', 'Periodo']
                df_disponibilidad = df_disponibilidad[columnas_a_mostrar]
                df_disponibilidad = df_disponibilidad.rename(columns={'Time Spent Final': 'Time Spent'})

                st.markdown("### 👤 Registros con comentarios de disponibilidad")
                autores_unicos = sorted(df_disponibilidad['Author'].dropna().unique())
                autores_seleccionados = st.multiselect(
                    "Filtrar por autor(es) (opcional):",
                    options=autores_unicos,
                    key="filtro_autor_disponibilidad"
                )

                if autores_seleccionados:
                    df_disponibilidad_filtrado = df_disponibilidad[df_disponibilidad['Author'].isin(autores_seleccionados)]
                else:
                    df_disponibilidad_filtrado = df_disponibilidad

                st.dataframe(df_disponibilidad_filtrado)

                excel_bytes_detalle = to_excel(df_disponibilidad_filtrado, "Detalle")
                st.download_button(
                    label="📥 Descargar registros como Excel",
                    data=excel_bytes_detalle,
                    file_name="disponibilidad_detallada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"❌ Error al procesar los archivos: {e}")

# -------------------------------
# SECCIÓN 3: REPORTE DE GESTIÓN
# -------------------------------

elif seccion == "📌 Reporte de gestión":
    
    st.markdown("<h1 style='color:#8700a6'>📌 Reporte de gestión</h1>", unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Sube el archivo Tracking a analizar (Excel)",
        type=["xlsx"],
        key="reporte_gestion"
    )

    clasificaciones = load_json('./clasificaciones.json')

    if uploaded_file:
        try:  
            df = pd.read_excel(uploaded_file)
            key_columns = {'Comment', 'Issue Summary'}
            
            if key_columns.issubset(set(df.columns)):
                # Reemplazar comentarios vacíos por el Issue Summary
                df['Comment'] = df['Comment'].fillna('').astype(str)
                df['Issue Summary'] = df['Issue Summary'].fillna('').astype(str)
                df['Comment'] = df.apply(
                    lambda row: row['Comment'] if row['Comment'].strip() else row['Issue Summary'],
                    axis=1)

                df[['Clasificacion', 'Supervisado']] = df['Comment'].apply(analizar_comentario)
                st.markdown("### 🧮 Resultados clasificados")
                
                df = df.sort_values(by='Clasificacion')
                df["final_tag"] = df["tag"].fillna(df["Clasificacion"])
                excel_data = to_excel(df)
                st.download_button(
                    label="📥 Descargar registros como Excel",
                    data=excel_data,
                    file_name="reporte_clasificacion.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.dataframe(df)
            else:
                st.error(f"❌ El archivo debe contener las columnas {key_columns}.")
            
        except Exception as e:
            st.error(f"❌ Error al procesar el archivo: {e}")
