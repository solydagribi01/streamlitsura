import streamlit as st

import pandas as pd

import matplotlib.pyplot as plt

import seaborn as sns

import numpy as np

import io



st.set_page_config(page_title="Análisis y Evaluación de Avisos", layout="wide")

st.title("Aplicación de Análisis y Evaluación")



# --- Subida del archivo ---

st.sidebar.header("1. Subir archivo Excel")

file = st.sidebar.file_uploader("Selecciona tu archivo .xlsx", type=["xlsx"])



if file:

    @st.cache_data

    def load_and_merge_data(uploaded_file):

        xls = pd.ExcelFile(uploaded_file)

        iw29 = pd.read_excel(xls, sheet_name=0)

        iw39 = pd.read_excel(xls, sheet_name=1)

        ih08 = pd.read_excel(xls, sheet_name=2)

        iw65 = pd.read_excel(xls, sheet_name=3)

        zpm015 = pd.read_excel(xls, sheet_name=4)



        for df in (iw29, iw39, ih08, iw65, zpm015):

            df.columns = df.columns.str.strip()



        equipo_original = iw29[["Aviso", "Equipo", "Duración de parada", "Descripción"]].copy()

        iw39_subset = iw39[["Aviso", "Total general (real)"]]



        tmp1 = pd.merge(iw29, iw39_subset, on="Aviso", how="left")

        tmp2 = pd.merge(tmp1, iw65, on="Aviso", how="left")

        tmp2.drop(columns=["Equipo"], errors='ignore', inplace=True)

        tmp2 = pd.merge(tmp2, equipo_original, on="Aviso", how="left")



        tmp3 = pd.merge(tmp2, ih08[[

            "Equipo", "Inic.garantía prov.", "Fin garantía prov.", "Texto", "Indicador ABC", "Denominación de objeto técnico"

        ]], on="Equipo", how="left")



        tmp4 = pd.merge(tmp3, zpm015[["Equipo", "TIPO DE SERVICIO"]], on="Equipo", how="left")



        tmp4.rename(columns={

            "Texto": "Texto_equipo",

            "Total general (real)": "Costes tot.reales"

        }, inplace=True)



        columnas_finales = [

            "Aviso",

            "Orden",

            "Fecha de aviso",

            "Código postal",

            "Status del sistema",

            "Descripción",

            "Ubicación técnica",

            "Indicador",

            "Equipo",

            "Denominación de objeto técnico",

            "Denominación ejecutante",

            "Duración de parada",

            "Centro de coste",

            "Costes tot.reales",

            "Inic.garantía prov.",

            "Fin garantía prov.",

            "Texto_equipo",

            "Indicador ABC",

            "Texto código acción",

            "Texto de acción",

            "Texto grupo acción",

            "TIPO DE SERVICIO"

        ]



        columnas_finales = [col for col in columnas_finales if col in tmp4.columns]

        return tmp4[columnas_finales]



    df = load_and_merge_data(file)

    st.success("Archivo procesado correctamente. ✅")



    # --- Filtros ---

    st.sidebar.header("2. Filtros")

    proveedor_opciones = df["Proveedor"].dropna().unique() if "Proveedor" in df.columns else []

    equipo_opciones = df["Equipo"].dropna().unique() if "Equipo" in df.columns else []



    proveedor = st.sidebar.multiselect("Filtrar por Proveedor", proveedor_opciones)

    equipo = st.sidebar.multiselect("Filtrar por Equipo", equipo_opciones)



    df_filtrado = df.copy()

    if proveedor:

        df_filtrado = df_filtrado[df_filtrado["Proveedor"].isin(proveedor)]

    if equipo:

        df_filtrado = df_filtrado[df_filtrado["Equipo"].isin(equipo)]



    # --- Funciones de indicadores ---

    def calcular_mttr(df):

        return df["Duración de parada"].mean() if "Duración de parada" in df.columns else np.nan



    def calcular_mtbf(df):

        equipos = df["Equipo"].nunique() if "Equipo" in df.columns else 0

        return df["Duración de parada"].sum() / equipos if equipos else np.nan



    def calcular_disponibilidad(mttr, mtbf):

        return mtbf / (mtbf + mttr) if (mttr and mtbf and (mttr + mtbf) != 0) else np.nan



    # --- Menú de navegación ---

    opcion = st.sidebar.radio("3. Selecciona una opción:", ["Análisis", "Evaluación"])



    # --- ANÁLISIS ---

    if opcion == "Análisis":

        st.header("🔍 Análisis de Costos y Equipos")



        if "Denominación ejecutante" in df_filtrado.columns and "Costes tot.reales" in df_filtrado.columns:

            costos_por_ejecutante = df_filtrado.groupby("Denominación ejecutante")["Costes tot.reales"].sum().sort_values()

            fig, ax = plt.subplots(figsize=(10, 5))

            sns.barplot(x=costos_por_ejecutante.values, y=costos_por_ejecutante.index, palette="Blues_r", ax=ax)

            ax.set_xlabel("Costo Total ($)")

            ax.set_ylabel("Ejecutante")

            ax.set_title("Costo Total por Ejecutante")

            st.pyplot(fig)

        else:

            st.warning("No se encontró la columna 'Denominación ejecutante' o 'Costes tot.reales'.")



        # --- Indicadores ---

        mttr = calcular_mttr(df_filtrado)

        mtbf = calcular_mtbf(df_filtrado)

        disponibilidad = calcular_disponibilidad(mttr, mtbf)



        st.subheader("📊 Indicadores")

        col1, col2, col3 = st.columns(3)

        col1.metric("MTTR (Media de tiempo de reparación)", f"{mttr:.2f}" if not np.isnan(mttr) else "N/A")

        col2.metric("MTBF (Media de tiempo entre fallas)", f"{mtbf:.2f}" if not np.isnan(mtbf) else "N/A")

        col3.metric("Disponibilidad", f"{disponibilidad:.2%}" if not np.isnan(disponibilidad) else "N/A")



        st.dataframe(df_filtrado.head(20))



    # --- EVALUACIÓN ---

    elif opcion == "Evaluación":

        st.header("✅ Evaluación Cualitativa")



        preguntas = [

            ("Calidad", "¿Las soluciones propuestas son coherentes con el diagnóstico y causa raíz del problema?", -1, 2),

            ("Calidad", "¿El trabajo entregado tiene materiales nuevos, originales y de marcas reconocidas?", -1, 2),

            ("Calidad", "¿Cuenta con acabados homogéneos, limpios y pulidos?", -1, 2),

            ("Calidad", "¿El trabajo entregado corresponde completamente con lo contratado?", -1, 2),

            ("Calidad", "¿La facturación refleja correctamente lo ejecutado y acordado?", -1, 2)

        ]



        aviso_seleccionado = st.selectbox("Selecciona un aviso para evaluar", df_filtrado["Aviso"].unique())

        aviso_data = df_filtrado[df_filtrado["Aviso"] == aviso_seleccionado].iloc[0]



        st.write("### Detalles del aviso")

        st.write({

            "Equipo": aviso_data.get("Equipo"),

            "Descripción": aviso_data.get("Descripción"),

            "Duración de parada": aviso_data.get("Duración de parada"),

            "Costes tot.reales": aviso_data.get("Costes tot.reales")

        })



        st.write("### Evaluación cualitativa con preguntas")



        respuestas = []

        for area, pregunta, min_val, max_val in preguntas:

            valor = st.slider(pregunta, min_val, max_val, 0)

            respuestas.append(valor)



        promedio = np.mean(respuestas)

        st.success(f"Puntaje promedio: {promedio:.2f} / 2")



        # --- Mostrar resultados en tabla ---

        columnas = [f"P{i+1}" for i in range(len(respuestas))]

        evaluacion_df = pd.DataFrame({

            "Aviso": [aviso_seleccionado],

            **{col: [val] for col, val in zip(columnas, respuestas)},

            "Promedio": [promedio]

        })



        st.write("### Resultado de evaluación")

        st.dataframe(evaluacion_df)



    # --- Descarga de archivo procesado ---

    output = io.BytesIO()

    df_filtrado.to_excel(output, index=False, engine='xlsxwriter')

    st.sidebar.download_button(

        label="📥 Descargar Excel filtrado",

        data=output.getvalue(),

        file_name="avisos_filtrados.xlsx",

        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    )



else:

    st.info("Por favor sube un archivo Excel para comenzar.")
