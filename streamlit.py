import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import io

# Configuración de la página de Streamlit
st.set_page_config(page_title="Análisis y Evaluación de Avisos", layout="wide")
st.title("Aplicación de Análisis y Evaluación")

# --- Sección de subida del archivo ---
st.sidebar.header("1. Subir archivo Excel")
file = st.sidebar.file_uploader("Selecciona tu archivo .xlsx", type=["xlsx"])

# Se ejecuta solo si se ha subido un archivo
if file:
    # Decorador para cachear los datos y evitar recargas en cada interacción
    @st.cache_data
    def load_and_merge_data(uploaded_file):
        """
        Carga las hojas de un archivo Excel y las fusiona en un único DataFrame.
        Realiza la limpieza de nombres de columnas y conversión de tipos.
        """
        xls = pd.ExcelFile(uploaded_file)
        
        # Cargar cada hoja en un DataFrame
        iw29 = pd.read_excel(xls, sheet_name=0)
        iw39 = pd.read_excel(xls, sheet_name=1)
        ih08 = pd.read_excel(xls, sheet_name=2)
        iw65 = pd.read_excel(xls, sheet_name=3)
        zpm015 = pd.read_excel(xls, sheet_name=4)

        # Limpiar espacios en blanco de los nombres de las columnas en todos los DataFrames
        for df_sheet in (iw29, iw39, ih08, iw65, zpm015):
            df_sheet.columns = df_sheet.columns.str.strip()

        # Iniciar el DataFrame principal con los datos de iw29
        df_merged = iw29.copy()

        # Preparar y fusionar datos de iw39 (Costes tot.reales)
        # Renombrar la columna antes de la fusión para evitar conflictos
        iw39_subset = iw39[["Aviso", "Total general (real)"]].rename(columns={"Total general (real)": "Costes tot.reales"})
        df_merged = pd.merge(df_merged, iw39_subset, on="Aviso", how="left")

        # Fusionar datos de iw65 (columnas de acción)
        # Usar sufijos para manejar posibles columnas duplicadas, manteniendo las de df_merged por defecto
        df_merged = pd.merge(df_merged, iw65, on="Aviso", how="left", suffixes=('', '_iw65'))
        
        # Si 'Equipo' fue duplicado por iw65, preferimos la columna original de iw29 (sin sufijo)
        if 'Equipo_iw65' in df_merged.columns:
            df_merged.drop(columns=['Equipo_iw65'], inplace=True)

        # Preparar y fusionar datos de ih08 (detalles del equipo)
        ih08_cols_to_merge = ih08[[
            "Equipo", "Inic.garantía prov.", "Fin garantía prov.", "Texto", "Indicador ABC", "Denominación de objeto técnico"
        ]].rename(columns={"Texto": "Texto_equipo"})
        
        # Fusionar ih08. Si hay columnas con el mismo nombre que ya existen en df_merged
        # (ej. 'Denominación de objeto técnico' si viniera de iw29), la de ih08 recibirá sufijo '_ih08'.
        df_merged = pd.merge(df_merged, ih08_cols_to_merge, on="Equipo", how="left", suffixes=('', '_ih08'))
        
        # Si 'Denominación de objeto técnico' fue duplicada por ih08,
        # preferimos explícitamente la versión de ih08 y eliminamos la duplicada.
        if 'Denominación de objeto técnico_ih08' in df_merged.columns:
            df_merged['Denominación de objeto técnico'] = df_merged['Denominación de objeto técnico_ih08']
            df_merged.drop(columns=['Denominación de objeto técnico_ih08'], inplace=True)

        # Fusionar datos de zpm015 (Tipo de servicio)
        df_merged = pd.merge(df_merged, zpm015[["Equipo", "TIPO DE SERVICIO"]], on="Equipo", how="left")

        # Convertir las columnas a tipo numérico DESPUÉS de todas las fusiones
        # Esto asegura que las columnas existan y sean Series antes de la conversión.
        df_merged["Duración de parada"] = pd.to_numeric(df_merged["Duración de parada"], errors='coerce')
        df_merged["Costes tot.reales"] = pd.to_numeric(df_merged["Costes tot.reales"], errors='coerce')

        # Definir el orden final y las columnas deseadas.
        # Es importante que estos nombres coincidan con los nombres finales de las columnas después de las fusiones y renombramientos.
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

        # Filtrar solo las columnas que realmente existen en el DataFrame fusionado final
        final_columns_exist = [col for col in columnas_finales if col in df_merged.columns]
        return df_merged[final_columns_exist]

    # Cargar y fusionar los datos
    df = load_and_merge_data(file)
    st.success("Archivo procesado correctamente. ✅")

    # --- Sección de Filtros ---
    st.sidebar.header("2. Filtros")
    # El filtro por 'Proveedor' se ha eliminado ya que la columna no se carga/fusiona
    # Puedes añadirlo si esa columna se incorpora en el futuro.
    
    # Opciones de filtro para 'Equipo'
    equipo_opciones = df["Equipo"].dropna().unique() if "Equipo" in df.columns else []
    equipo = st.sidebar.multiselect("Filtrar por Equipo", equipo_opciones)

    # Aplicar filtros
    df_filtrado = df.copy()
    # if proveedor:
    #     df_filtrado = df_filtrado[df_filtrado["Proveedor"].isin(proveedor)]
    if equipo:
        df_filtrado = df_filtrado[df_filtrado["Equipo"].isin(equipo)]

    # --- Funciones para el cálculo de indicadores ---
    def calcular_mttr(df_input):
        """Calcula el Mean Time To Repair (MTTR) de un DataFrame."""
        # Se asegura que la columna 'Duración de parada' exista y sea numérica
        if "Duración de parada" in df_input.columns:
            return df_input["Duración de parada"].mean()
        return np.nan

    def calcular_mtbf(df_input):
        """Calcula el Mean Time Between Failures (MTBF) de un DataFrame."""
        # Nota: Esta es una simplificación. Un cálculo más preciso requeriría
        # el tiempo total de operación y el número de fallas reales.
        if "Duración de parada" in df_input.columns and "Equipo" in df_input.columns:
            equipos_unicos = df_input["Equipo"].nunique()
            if equipos_unicos > 0:
                # Filtrar NaNs en 'Duración de parada' antes de sumar para MTBF
                return df_input["Duración de parada"].sum() / equipos_unicos
        return np.nan

    def calcular_disponibilidad(mttr, mtbf):
        """Calcula la disponibilidad basándose en MTTR y MTBF."""
        if not np.isnan(mttr) and not np.isnan(mtbf) and (mttr + mtbf) != 0:
            return mtbf / (mtbf + mttr)
        return np.nan

    # --- Menú de navegación principal ---
    opcion = st.sidebar.radio("3. Selecciona una opción:", ["Análisis", "Evaluación"])

    # --- Sección de ANÁLISIS ---
    if opcion == "Análisis":
        st.header("🔍 Análisis de Costos y Equipos")

        # Gráfico de costo total por ejecutante
        if "Denominación ejecutante" in df_filtrado.columns and "Costes tot.reales" in df_filtrado.columns:
            # Asegurarse de eliminar NaNs antes de agrupar para evitar errores
            costos_por_ejecutante = df_filtrado.dropna(subset=["Denominación ejecutante", "Costes tot.reales"]).groupby("Denominación ejecutante")["Costes tot.reales"].sum().sort_values()
            
            if not costos_por_ejecutante.empty:
                fig, ax = plt.subplots(figsize=(10, 5))
                sns.barplot(x=costos_por_ejecutante.values, y=costos_por_ejecutante.index, palette="Blues_r", ax=ax)
                ax.set_xlabel("Costo Total ($)")
                ax.set_ylabel("Ejecutante")
                ax.set_title("Costo Total por Ejecutante")
                st.pyplot(fig)
            else:
                st.warning("No hay datos para mostrar el costo total por ejecutante después del filtrado.")
        else:
            st.warning("No se encontraron las columnas 'Denominación ejecutante' o 'Costes tot.reales' en los datos filtrados.")

        # Mostrar indicadores clave
        st.subheader("📊 Indicadores")
        mttr = calcular_mttr(df_filtrado)
        mtbf = calcular_mtbf(df_filtrado)
        disponibilidad = calcular_disponibilidad(mttr, mtbf)

        col1, col2, col3 = st.columns(3)
        col1.metric("MTTR (Media de tiempo de reparación)", f"{mttr:.2f}" if not np.isnan(mttr) else "N/A")
        col2.metric("MTBF (Media de tiempo entre fallas)", f"{mtbf:.2f}" if not np.isnan(mtbf) else "N/A")
        col3.metric("Disponibilidad", f"{disponibilidad:.2%}" if not np.isnan(disponibilidad) else "N/A")

        # Mostrar una vista previa de los datos filtrados
        st.subheader("Datos filtrados (primeras 20 filas)")
        st.dataframe(df_filtrado.head(20))

    # --- Sección de EVALUACIÓN ---
    elif opcion == "Evaluación":
        st.header("✅ Evaluación Cualitativa")

        # Definición de preguntas para la evaluación
        preguntas = [
            ("Calidad", "¿Las soluciones propuestas son coherentes con el diagnóstico y causa raíz del problema?", -1, 2),
            ("Calidad", "¿El trabajo entregado tiene materiales nuevos, originales y de marcas reconocidas?", -1, 2),
            ("Calidad", "¿Cuenta con acabados homogéneos, limpios y pulidos?", -1, 2),
            ("Calidad", "¿El trabajo entregado corresponde completamente con lo contratado?", -1, 2),
            ("Calidad", "¿La facturación refleja correctamente lo ejecutado y acordado?", -1, 2)
        ]

        # Selección de un aviso para evaluar
        if not df_filtrado["Aviso"].empty:
            aviso_seleccionado = st.selectbox("Selecciona un aviso para evaluar", df_filtrado["Aviso"].unique())
            aviso_data = df_filtrado[df_filtrado["Aviso"] == aviso_seleccionado].iloc[0]

            st.write("### Detalles del aviso")
            # Mostrar detalles relevantes del aviso seleccionado
            st.write({
                "Equipo": aviso_data.get("Equipo", "N/A"),
                "Descripción": aviso_data.get("Descripción", "N/A"),
                "Duración de parada": aviso_data.get("Duración de parada", "N/A"),
                "Costes tot.reales": aviso_data.get("Costes tot.reales", "N/A")
            })

            st.write("### Evaluación cualitativa con preguntas")

            # Recopilar respuestas a las preguntas de evaluación
            respuestas = []
            for area, pregunta, min_val, max_val in preguntas:
                # El valor por defecto del slider se establece a 0, asumiendo neutralidad inicial
                valor = st.slider(pregunta, min_val, max_val, 0, key=f"slider_{aviso_seleccionado}_{pregunta}")
                respuestas.append(valor)

            # Calcular el promedio de las respuestas
            if respuestas:
                promedio = np.mean(respuestas)
                st.success(f"Puntaje promedio: {promedio:.2f} / 2")
            else:
                promedio = np.nan
                st.warning("No hay preguntas de evaluación definidas.")

            # Mostrar resultados en una tabla
            columnas_respuestas = [f"P{i+1}" for i in range(len(respuestas))]
            evaluacion_data = {
                "Aviso": [aviso_seleccionado],
                **{col: [val] for col, val in zip(columnas_respuestas, respuestas)},
                "Promedio": [promedio]
            }
            evaluacion_df = pd.DataFrame(evaluacion_data)

            st.write("### Resultado de evaluación")
            st.dataframe(evaluacion_df)
        else:
            st.warning("No hay avisos para evaluar. Por favor, asegúrate de que tus datos contienen la columna 'Aviso'.")

    # --- Sección de descarga del archivo procesado ---
    output = io.BytesIO()
    # Asegurarse de que el motor 'xlsxwriter' esté instalado si se usa
    try:
        df_filtrado.to_excel(output, index=False, engine='xlsxwriter')
    except ImportError:
        st.error("El motor 'xlsxwriter' no está instalado. Por favor, instálalo con 'pip install xlsxwriter'.")
        df_filtrado.to_excel(output, index=False) # Intentar sin motor específico si xlsxwriter no está
    
    # Botón de descarga
    st.sidebar.download_button(
        label="📥 Descargar Excel filtrado",
        data=output.getvalue(),
        file_name="avisos_filtrados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Mensaje inicial si no se ha subido ningún archivo
else:
    st.info("Por favor sube un archivo Excel para comenzar.")
