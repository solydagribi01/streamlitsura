
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

st.set_page_config(page_title="Análisis de Propuesta", layout="wide")

st.title("App de Análisis y Evaluación")

st.sidebar.header("Opciones")
opcion = st.sidebar.radio("Selecciona una opción:", ["Análisis", "Evaluación"])

uploaded_file = st.sidebar.file_uploader("Carga el archivo Excel", type=["xlsx"])

if uploaded_file is not None:
    # Carga de datos
    st.success("Archivo cargado correctamente.")
    try:
        # Aquí insertamos el código original (ya limpiado)
        
        
        import pandas as pd
        import matplotlib.pyplot as plt
        import seaborn as sns
        import ipywidgets as widgets
        from IPython.display import display, clear_output, HTML
        import re
        import matplotlib.ticker as mticker
        import numpy as np
        import os
        
        
        # ------------------------ 1. Montar Drive (opcional) ----------------------
        try:
            from google.colab import drive  # type: ignore
            drive.mount('/content/drive')
            print("✅ Google Drive montado.")
        except ModuleNotFoundError:
            print("▶️ Ejecutando fuera de Colab – se omite el montaje de Drive.")
        
        # ------------------------ 2. Ruta del archivo ----------------------------
        FILEPATH = '/content/DATA2.XLSX'
        OUTPUT_PATH = '/content/avisos_filtrados.xlsx'
        
        # ------------------------ 3. Función de carga & unión --------------------
        def load_and_merge_data(path: str) -> pd.DataFrame:
            # Cargar hojas
            iw29 = pd.read_excel(path, sheet_name=0)  # Hoja 1
            iw39 = pd.read_excel(path, sheet_name=1)  # Hoja 2
            ih08 = pd.read_excel(path, sheet_name=2)  # Hoja 3
            iw65 = pd.read_excel(path, sheet_name=3)  # Hoja 4
            zpm015 = pd.read_excel(path, sheet_name=4)  # Hoja 5
        
            # Limpiar encabezados
            for df in (iw29, iw39, ih08, iw65, zpm015):
                df.columns = df.columns.str.strip()
        
            # Guardar "Equipo" original desde IW29 para evitar pérdida
            equipo_original = iw29[["Aviso", "Equipo", "Duración de parada","Descripción", ]].copy()
        
            # Extraer solo columnas necesarias de iw39 para el merge (incluyendo 'Total general (real)')
            iw39_subset = iw39[["Aviso", "Total general (real)"]]
        
            # Unir por 'Aviso'
            tmp1 = pd.merge(iw29, iw39_subset, on="Aviso", how="left")
            tmp2 = pd.merge(tmp1, iw65, on="Aviso", how="left")
        
            # Restaurar el valor original de "Equipo" de IW29 después del merge
            tmp2.drop(columns=["Equipo"], errors='ignore', inplace=True)
            tmp2 = pd.merge(tmp2, equipo_original, on="Aviso", how="left")
        
            # Unir por 'Equipo' con IH08
            tmp3 = pd.merge(tmp2, ih08[[
                "Equipo", "Inic.garantía prov.", "Fin garantía prov.", "Texto", "Indicador ABC", "Denominación de objeto técnico"
            ]], on="Equipo", how="left")
        
            # Unir por 'Equipo' con ZPM015
            tmp4 = pd.merge(tmp3, zpm015[["Equipo", "TIPO DE SERVICIO"]], on="Equipo", how="left")
        
            # Renombrar columnas
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
        
        # ------------------------ 4. Ejecutar ------------------------------------
        
        try:
            df = load_and_merge_data(FILEPATH)
        except FileNotFoundError:
            raise SystemExit(f"❌ No se encontró el archivo en {FILEPATH}. Verifica la ruta.")
        
        # ➖ Eliminar registros cuyo 'Status del sistema' contenga "PTBO"
        df = df[~df["Status del sistema"].str.contains("PTBO", case=False, na=False)]
        
        # ➕ Dejar solo una fila con coste por cada aviso
        df['Costes tot.reales'] = df.groupby('Aviso')['Costes tot.reales'].transform(
            lambda x: [x.iloc[0]] + [0]*(len(x)-1)
        )
        
        print(f"✅ Datos filtrados. Filas: {len(df)} – Columnas: {len(df.columns)}")
        display(df.head())
        
        # Guardar resultado
        df.to_excel(OUTPUT_PATH, index=False)
        print(f"💾 Archivo final guardado en {OUTPUT_PATH}")
        
        
        import pandas as pd
        import matplotlib.pyplot as plt
        import seaborn as sns
        import ipywidgets as widgets
        from IPython.display import display, clear_output, HTML
        import re
        import matplotlib.ticker as mticker
        import numpy as np
        import os
        
        # Set a nice style for plots
        sns.set_style('whitegrid')
        
        # --- Data Loading and Preprocessing ---
        file_path = "/content/avisos_filtrados.xlsx"
        
        try:
            df = pd.read_excel(file_path)
        except FileNotFoundError:
            print(f"Error: The file '{file_path}' was not found. Please check the path.")
            exit()
        
        # Normalize column names more robustly
        # Ensure these original column names match your Excel file exactly
        ORIGINAL_EJECUTANTE_COL_NAME = "Denominación ejecutante"
        ORIGINAL_CP_COL_NAME = "Código postal"
        ORIGINAL_OBJETO_TECNICO_COL_NAME = "Denominación de objeto técnico"
        ORIGINAL_TEXTO_CODIGO_ACCION_COL_NAME = "Texto código acción"
        ORIGINAL_TEXTO_ACCION_COL_NAME = "Texto de acción"
        ORIGINAL_TIPO_SERVICIO_COL_NAME = "Tipo de servicio"
        ORIGINAL_COSTOS_COL_NAME = "Costes tot.reales"
        ORIGINAL_DESCRIPTION_COL_NAME = "Descripción" # This is 'Descripción' as seen in your error output
        ORIGINAL_FECHA_AVISO_COL_NAME = "Fecha de aviso"
        ORIGINAL_TEXTO_POSICION_COL_NAME = "Texto de Posición" # This is the missing column
        ORIGINAL_TEXTO_EQUIPO_COL_NAME = "Texto_equipo"
        ORIGINAL_DURACION_PARADA_COL_NAME = "Duración de parada"
        ORIGINAL_EQUIPO_COL_NAME = "Equipo"
        ORIGINAL_AVISO_COL_NAME = "Aviso"
        
        column_mapping = {
            ORIGINAL_EJECUTANTE_COL_NAME: "denominacion_ejecutante",
            ORIGINAL_CP_COL_NAME: "codigo_postal",
            ORIGINAL_OBJETO_TECNICO_COL_NAME: "denominacion_de_objeto_tecnico",
            ORIGINAL_TEXTO_CODIGO_ACCION_COL_NAME: "texto_codigo_accion",
            ORIGINAL_TEXTO_ACCION_COL_NAME: "texto_de_accion",
            ORIGINAL_TIPO_SERVICIO_COL_NAME: "tipo_de_servicio",
            ORIGINAL_COSTOS_COL_NAME: "costes_totreales",
            ORIGINAL_DESCRIPTION_COL_NAME: "descripcion",
            ORIGINAL_FECHA_AVISO_COL_NAME: "fecha_de_aviso",
            ORIGINAL_TEXTO_POSICION_COL_NAME: "texto_de_posicion",
            ORIGINAL_TEXTO_EQUIPO_COL_NAME: "texto_equipo",
            ORIGINAL_DURACION_PARADA_COL_NAME: "duracion_de_parada",
            ORIGINAL_EQUIPO_COL_NAME: "equipo",
            ORIGINAL_AVISO_COL_NAME: "aviso"
        }
        
        # Normalize column names by stripping, lowercasing, and replacing spaces/accents
        normalized_df_columns = []
        for col in df.columns:
            found_match = False
            for original, normalized in column_mapping.items():
                if col.strip().lower() == original.strip().lower():
                    normalized_df_columns.append(normalized)
                    found_match = True
                    break
            if not found_match:
                # Fallback for columns not explicitly mapped
                normalized_df_columns.append(
                    col.lower()
                    .strip()
                    .replace(" ", "_")
                    .replace(".", "")
                    .replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
                )
        df.columns = normalized_df_columns
        
        print("Normalized DataFrame columns:", df.columns.tolist())
        
        # Assign relevant columns to new, simplified names for easier access (from first code)
        df['PROVEEDOR'] = df['denominacion_ejecutante']
        df['COSTO'] = df['costes_totreales']
        df['TIEMPO PARADA'] = pd.to_numeric(df['duracion_de_parada'], errors='coerce')
        df['EQUIPO'] = pd.to_numeric(df['equipo'], errors='coerce')
        df['AVISO'] = pd.to_numeric(df['aviso'], errors='coerce')
        df['TIPO DE SERVICIO'] = df['tipo_de_servicio']
        
        # Ensure 'costes_totreales' is numeric
        df['costes_totreales'] = pd.to_numeric(df['costes_totreales'], errors='coerce')
        
        # --- HORARIO Mapping (from first code) ---
        horarios_dict = {
            "HORARIO_99": (17, 364.91), "HORARIO_98": (14.5, 312.78), "HORARIO_97": (9.818181818, 286.715),
            "HORARIO_96": (14.5, 312.78), "HORARIO_95": (4, 208.52), "HORARIO_93": (13.45454545, 286.715),
            "HORARIO_92": (6, 338.845), "HORARIO_91": (9.25, 312.78), "HORARIO_90": (11, 260.65),
            "HORARIO_9": (16, 312.78), "HORARIO_89": (9.5, 260.65), "HORARIO_88": (14, 260.65),
            "HORARIO_87": (9.333333333, 312.78), "HORARIO_86": (9.666666667, 312.78), "HORARIO_85": (12, 312.78),
            "HORARIO_84": (9.5, 312.78), "HORARIO_83": (8.416666667, 312.78), "HORARIO_82": (6, 312.78),
            "HORARIO_81": (10, 312.78), "HORARIO_80": (8.5, 312.78), "HORARIO_8": (11.6, 260.65),
            "HORARIO_79": (14, 312.78), "HORARIO_78": (12, 312.78), "HORARIO_77": (3, 312.78),
            "HORARIO_76": (16, 312.78), "HORARIO_75": (12.16666667, 312.78), "HORARIO_74": (11.33333333, 312.78),
            "HORARIO_73": (12.66666667, 312.78), "HORARIO_72": (11.83333333, 312.78), "HORARIO_71": (11, 312.78),
            "HORARIO_70": (15.16666667, 312.78), "HORARIO_7": (15.33333333, 312.78), "HORARIO_69": (9.166666667, 312.78),
            "HORARIO_68": (4, 312.78), "HORARIO_67": (10, 260.65), "HORARIO_66": (4, 260.65),
            "HORARIO_65": (16.76923077, 338.845), "HORARIO_64": (17.15384615, 338.845), "HORARIO_63": (22.5, 312.78),
            "HORARIO_62": (12.25, 312.78), "HORARIO_61": (4, 312.78), "HORARIO_60": (13, 312.78),
            "HORARIO_6": (18.46153846, 338.845), "HORARIO_59": (12.66666667, 312.78), "HORARIO_58": (12.33333333, 312.78),
            "HORARIO_57": (13.53846154, 338.845), "HORARIO_56": (12.16666667, 312.78), "HORARIO_55": (6.333333333, 312.78),
            "HORARIO_54": (7.230769231, 338.845), "HORARIO_53": (5.5, 312.78), "HORARIO_52": (4, 312.78),
            "HORARIO_51": (14, 338.845), "HORARIO_50": (15, 312.78), "HORARIO_5": (17, 312.78),
            "HORARIO_49": (15.27272727, 286.715), "HORARIO_48": (14.76923077, 338.845), "HORARIO_47": (14.5, 312.78),
            "HORARIO_46": (14.33333333, 312.78), "HORARIO_45": (14.16666667, 312.78), "HORARIO_44": (13.83333333, 312.78),
            "HORARIO_43": (13.5, 312.78), "HORARIO_42": (13.91666667, 312.78), "HORARIO_41": (15, 364.91),
            "HORARIO_40": (15.81818182, 286.715), "HORARIO_4": (16.16666667, 312.78), "HORARIO_39": (15.27272727, 286.715),
            "HORARIO_38": (13.84615385, 338.845), "HORARIO_37": (15.09090909, 286.715), "HORARIO_36": (14, 364.91),
            "HORARIO_35": (14.30769231, 338.845), "HORARIO_34": (14.90909091, 286.715), "HORARIO_33": (13.55, 312.78),
            "HORARIO_32": (14, 338.845), "HORARIO_31": (14.72727273, 286.715), "HORARIO_30": (13.08333333, 312.78),
            "HORARIO_3": (16, 312.78), "HORARIO_29": (14, 286.715), "HORARIO_28": (13, 364.91),
            "HORARIO_27": (14, 286.715), "HORARIO_26": (12.58333333, 312.78), "HORARIO_25": (12, 312.78),
            "HORARIO_24": (13.27272727, 286.715), "HORARIO_23": (11.83333333, 312.78), "HORARIO_22": (11.91666667, 312.78),
            "HORARIO_21": (13.09090909, 286.715), "HORARIO_20": (5, 312.78), "HORARIO_2": (23.5, 364.91),
            "HORARIO_19": (12.18181818, 286.715), "HORARIO_18": (5, 312.78), "HORARIO_17": (9.75, 312.78),
            "HORARIO_16": (10.36363636, 286.715), "HORARIO_15": (10.18181818, 286.715), "HORARIO_14": (8.5, 312.78),
            "HORARIO_134": (12, 364.91), "HORARIO_133": (12, 260.65), "HORARIO_132": (13, 312.78),
            "HORARIO_131": (10, 312.78), "HORARIO_130": (11, 260.65), "HORARIO_13": (9.454545455, 286.715),
            "HORARIO_129": (9.384615385, 338.845), "HORARIO_128": (12.33333333, 312.78), "HORARIO_127": (9.666666667, 312.78),
            "HORARIO_126": (10.83333333, 312.78), "HORARIO_125": (4, 312.78), "HORARIO_124": (13.66666667, 312.78),
            "HORARIO_123": (16.61538462, 338.845), "HORARIO_122": (11, 260.65), "HORARIO_121": (11.66666667, 312.78),
            "HORARIO_120": (8.25, 312.78), "HORARIO_12": (9.272727273, 286.715), "HORARIO_119": (11.23076923, 338.845),
            "HORARIO_118": (11.27272727, 286.715), "HORARIO_117": (11.41666667, 312.78), "HORARIO_116": (11, 312.78),
            "HORARIO_115": (9.25, 312.78), "HORARIO_114": (23.07692308, 338.845), "HORARIO_113": (20, 338.845),
            "HORARIO_112": (10.61538462, 338.845), "HORARIO_111": (9.454545455, 286.715), "HORARIO_110": (6.833333333, 312.78),
            "HORARIO_11": (8, 312.78), "HORARIO_109": (12.90909091, 286.715), "HORARIO_108": (10.54545455, 286.715),
            "HORARIO_107": (12.61538462, 338.845), "HORARIO_106": (14.76923077, 338.845), "HORARIO_105": (12, 156.39),
            "HORARIO_104": (7.666666667, 312.78), "HORARIO_103": (3, 260.65), "HORARIO_102": (10.16666667, 312.78),
            "HORARIO_101": (12, 260.65), "HORARIO_100": (11.16666667, 312.78), "HORARIO_10": (6, 312.78),
            "HORARIO_1": (24, 364.91),
        }
        df['HORARIO'] = df['texto_equipo'].str.strip().str.upper()
        df['HORA/ DIA'] = df['HORARIO'].map(lambda x: horarios_dict.get(x, (None, None))[0])
        df['DIAS/ AÑO'] = df['HORARIO'].map(lambda x: horarios_dict.get(x, (None, None))[1])
        df['DIAS/ AÑO'] = pd.to_numeric(df['DIAS/ AÑO'], errors='coerce')
        df['HORA/ DIA'] = pd.to_numeric(df['HORA/ DIA'], errors='coerce')
        
        
        # --- Initial Filtering from first code ---
        initial_rows = len(df)
        df = df.dropna(subset=['EQUIPO'])
        rows_after_filter_equipo = len(df)
        print(f"Filtrado: {initial_rows - rows_after_filter_equipo} avisos sin equipo fueron excluidos.")
        
        
        # --- Additional Preprocessing for Second Code's requirements ---
        df["fecha_de_aviso"] = pd.to_datetime(df["fecha_de_aviso"], errors="coerce")
        df["año"] = df["fecha_de_aviso"].dt.year
        df["mes"] = df["fecha_de_aviso"].dt.strftime("%B")
        
        def extract_description_category(description):
            if pd.isna(description):
                return "Otros"
            match = re.match(r'^([A-Z]{2})/', str(description).strip())
            if match:
                return match.group(1)
            return "Otros"
        
        df["description_category"] = df['descripcion'].apply(extract_description_category)
        print("Description categories created:", df["description_category"].unique())
        
        # --- DEFINICIÓN DE PREGUNTAS PARA EVALUACIÓN (from first code) ---
        preguntas = [
            ("Calidad", "¿Las soluciones propuestas son coherentes con el diagnóstico y causa raíz del problema?", "2,1,0,-1"),
            ("Calidad", "¿El trabajo entregado tiene materiales nuevos, originales y de marcas reconocidas?", "2,1,0,-1"),
            ("Calidad", "¿Cuenta con acabados homogéneos, limpios y pulidos?", "2,1,0,-1"),
            ("Calidad", "¿El trabajo entregado corresponde completamente con lo contratado?", "2,1,0,-1"),
            ("Calidad", "¿La facturación refleja correctamente lo ejecutado y acordado?", "2,1,0,-1"),
            ("Oportunidad", "¿La entrega de cotizaciones fue oportuna, según el contrato?", "2,1,0,-1"),
            ("Oportunidad", "¿El reporte del servicio fue entregado oportunamente, según el contrato?", "2,1,0,-1"),
            ("Oportunidad", "¿Cumple las fechas y horas programadas para los trabajos, según el contrato?", "2,1,0,-1"),
            ("Oportunidad", "¿Responde de forma efectiva ante eventualidades emergentes, según el contrato?", "2,1,0,-1"),
            ("Oportunidad", "¿Soluciona rápidamente reclamos o inquietudes por garantía, según el contrato?", "2,1,0,-1"),
            ("Oportunidad", "¿Dispone de los repuestos requeridos en los tiempos necesarios, según el contrato?", "2,1,0,-1"),
            ("Oportunidad", "¿Entrega las facturas en los tiempos convenidos, según el contrato?", "2,1,0,-1"),
            ("Precio", "¿Los precios ofrecidos para equipos son competitivos respecto al mercado?", "2,1,0,-1"),
            ("Precio", "¿Los precios ofrecidos para repuestos son competitivos respecto al mercado?", "2,1,0,-1"),
            ("Precio", "¿Los precios ofrecidos para mantenimientos son competitivos respecto al mercado?", "2,1,0,-1"),
            ("Precio", "¿Los precios ofrecidos para insumos son competitivos respecto al mercado?", "2,1,0,-1"),
            ("Precio", "Facilita llegar a una negociación (precios)", "2,1,0,-1"), # New question
            ("Precio", "Pone en consideración contratos y trabajos adjudicados en el último periodo de tiempo", "2,1,0,-1"), # New question
            ("Postventa", "¿Tiene disposición y actitud de servicio frente a solicitudes?", "2,1,0,-1"),
            ("Postventa", "¿Conoce necesidades y ofrece alternativas adecuadas?", "2,1,0,-1"),
            ("Postventa", "¿Realiza seguimiento a los resultados de los trabajos?", "2,1,0,-1"),
            ("Postventa", "¿Ofrece capacitaciones para el manejo de los equipos?", "2,1,0,-1"),
            ("Postventa", "¿Los métodos de capacitación ofrecidos son efectivos y adecuados?", "2,1,0,-1"),
            ("Desempeño técnico", "Disponibilidad promedio (%)", "auto"),
            ("Desempeño técnico", "MTTR promedio (hrs)", "auto"),
            ("Desempeño técnico", "MTBF promedio (hrs)", "auto"),
            ("Desempeño técnico", "Rendimiento promedio equipos", "auto"),
        ]
        
        # --- FUNCIONES DE CÁLCULO DE INDICADORES (from first code) ---
        def calcular_indicadores_servicio(df_sub):
            if df_sub.empty:
                return pd.Series(dtype=object), pd.Series(dtype=object), pd.Series(dtype=object), pd.Series(dtype=object), pd.Series(dtype=object), pd.Series(dtype=object)
        
            cnt = df_sub['TIPO DE SERVICIO'].value_counts()
            cost = df_sub.groupby('TIPO DE SERVICIO')['COSTO'].sum()
            mttr = df_sub.groupby('TIPO DE SERVICIO')['TIEMPO PARADA'].mean()
        
            ttot = df_sub.groupby('TIPO DE SERVICIO').apply(
                lambda g: (g['DIAS/ AÑO'].mean() * g['HORA/ DIA'].mean()) if not g['DIAS/ AÑO'].isnull().all() and not g['HORA/ DIA'].isnull().all() else np.nan
            )
        
            down = df_sub.groupby('TIPO DE SERVICIO')['TIEMPO PARADA'].sum()
            fails = df_sub.groupby('TIPO DE SERVICIO')['AVISO'].count()
            mtbf = (ttot - down) / fails.replace(0, np.nan)
        
            disp = (mtbf / (mtbf + mttr)).replace([np.inf, -np.inf], np.nan) * 100
            rend = disp.apply(lambda v: 'Alto' if v >= 90 else ('Medio' if v >= 75 else 'Bajo') if not pd.isna(v) else np.nan)
            return cnt, cost, mttr, mtbf, disp, rend
        
        def calcular_indicadores_equipo(df_sub):
            if df_sub.empty:
                return pd.Series(dtype=object), pd.Series(dtype=object), pd.Series(dtype=object), pd.Series(dtype=object), pd.Series(dtype=object), pd.Series(dtype=object), pd.Series(dtype=object)
        
            cnt_equipo = df_sub['EQUIPO'].value_counts()
            cost_equipo = df_sub.groupby('EQUIPO')['COSTO'].sum()
            mttr_equipo = df_sub.groupby('EQUIPO')['TIEMPO PARADA'].mean()
        
            # Get description category for each unique equipment
            # We need to ensure that 'descripcion' and 'description_category' columns exist and are not NaN for the equipment
            # Group by 'EQUIPO' and take the first valid 'description_category'
            desc_category_equipo = df_sub.groupby('EQUIPO')['description_category'].first().fillna('Desconocido')
        
        
            ttot_equipo = df_sub.groupby('EQUIPO').apply(
                lambda g: (g['DIAS/ AÑO'].mean() * g['HORA/ DIA'].mean()) if not g['DIAS/ AÑO'].isnull().all() and not g['HORA/ DIA'].isnull().all() else np.nan
            )
        
            down_equipo = df_sub.groupby('EQUIPO')['TIEMPO PARADA'].sum()
            fails_equipo = df_sub.groupby('EQUIPO')['AVISO'].count()
            mtbf_equipo = (ttot_equipo - down_equipo) / fails_equipo.replace(0, np.nan)
        
            disp_equipo = (mtbf_equipo / (mtbf_equipo + mttr_equipo)).replace([np.inf, -np.inf], np.nan) * 100
            rend_equipo = disp_equipo.apply(lambda v: 'Alto' if v >= 90 else ('Medio' if v >= 75 else 'Bajo') if not pd.isna(v) else np.nan)
            return cnt_equipo, cost_equipo, mttr_equipo, mtbf_equipo, disp_equipo, rend_equipo, desc_category_equipo
        
        
        # --- EVALUATION WIDGETS AND LOGIC (from first code) ---
        class EvaluacionProveedoresApp:
            def __init__(self, df):
                self.df = df
                self.prov_sel = widgets.Dropdown(
                    options=["Todos"] + sorted(self.df['PROVEEDOR'].dropna().unique().tolist()),
                    description='Proveedor:',
                    layout=widgets.Layout(width='300px')
                )
                self.out = widgets.Output()
                self.generar_resumen_btn = widgets.Button(description="Generar Resumen de Evaluación y Exportar a Excel")
                self.generar_resumen_btn.on_click(self.generar_resumen_evaluacion)
                self.mostrar_rangos_btn = widgets.Button(description="Mostrar Rangos de Respuesta")
                self.mostrar_rangos_btn.on_click(self.mostrar_rangos_respuesta)
        
                self.summary_servicio_global_for_export = None
                self.resumen_equipo_global_for_export = None
        
                self.all_evaluation_widgets_map = {}
                self.all_provider_service_types = []
                self.current_provider_metrics = {}
                self.current_page = 0
                self.evaluation_form_container = widgets.VBox([])
                self.service_type_mapping_display = widgets.Output()
        
        
                self.prov_sel.observe(self.on_proveedor_change, names='value')
        
                self.ui = widgets.VBox([self.prov_sel, self.out])
        
            def crear_widgets_evaluacion(self, mttr, mtbf, disp, rend, service_types_to_display):
                service_type_labels_on_page = []
                for st in service_types_to_display:
                    # Handle cases where st might not be in all_provider_service_types
                    if st in self.all_provider_service_types:
                        global_idx = self.all_provider_service_types.index(st) + 1
                        service_type_labels_on_page.append(f"Servicio {global_idx} ({st})")
                    else:
                        service_type_labels_on_page.append(f"Servicio Desconocido ({st})")
        
        
                with self.service_type_mapping_display:
                    clear_output()
                    if service_types_to_display:
                        print("--- Mapeo de Tipos de Servicio en esta página ---")
                        for st in service_types_to_display:
                            print(f"Servicio {self.all_provider_service_types.index(st) + 1}: {st}")
        
                header_widgets = [widgets.Label("Pregunta", layout=widgets.Layout(width='40%'))]
                for st_label in service_type_labels_on_page:
                    header_widgets.append(widgets.Label(st_label, layout=widgets.Layout(width='10%')))
                evaluation_rows = [widgets.HBox(header_widgets)]
        
                for cat, texto, escala in preguntas:
                    row_widgets = [widgets.Label(f"[{cat}] {texto}", layout=widgets.Layout(width='40%'))]
                    for st_original in service_types_to_display:
                        if escala == "auto":
                            val = 0
                            if 'Disponibilidad' in texto and st_original in disp and not pd.isna(disp[st_original]):
                                mean_disp = disp[st_original]
                                val = 2 if mean_disp >= 98 else (1 if mean_disp >= 75 else 0)
                            elif 'MTTR' in texto and st_original in mttr and not pd.isna(mttr[st_original]):
                                mean_mttr = mttr[st_original]
                                val = 2 if mean_mttr <= 5 else (1 if mean_mttr <= 20 else 0)
                            elif 'MTBF' in texto and st_original in mtbf and not pd.isna(mtbf[st_original]):
                                mean_mtbf = mtbf[st_original]
                                val = 2 if mean_mtbf > 1000 else (1 if mean_mtbf >= 100 else 0)
                            elif 'Rendimiento' in texto and st_original in rend and not pd.isna(rend[st_original]):
                                perf_class = rend[st_original]
                                if perf_class == 'Alto':
                                    val = 2
                                elif perf_class == 'Medio':
                                    val = 1
                                elif perf_class == 'Bajo':
                                    val = 0
                            widget = widgets.Label(str(val), layout=widgets.Layout(width='10%'))
                            self.all_evaluation_widgets_map[(cat, texto, st_original)] = widget
                        else:
                            opts = [('Sobresaliente (2)', 2), ('Bueno (1)', 1), ('Indiferente (0)', 0), ('Malo (-1)', -1)]
                            if (cat, texto, st_original) in self.all_evaluation_widgets_map:
                                widget = self.all_evaluation_widgets_map[(cat, texto, st_original)]
                            else:
                                widget = widgets.Dropdown(options=opts, value=0, layout=widgets.Layout(width='10%'))
                                self.all_evaluation_widgets_map[(cat, texto, st_original)] = widget
                        row_widgets.append(widget)
                    evaluation_rows.append(widgets.HBox(row_widgets))
                return widgets.VBox(evaluation_rows)
        
            def update_evaluation_display(self):
                cnt, cost, mttr, mtbf, disp, rend = (
                    self.current_provider_metrics.get('cnt', pd.Series()),
                    self.current_provider_metrics.get('cost', pd.Series()),
                    self.current_provider_metrics.get('mttr', pd.Series()),
                    self.current_provider_metrics.get('mtbf', pd.Series()),
                    self.current_provider_metrics.get('disp', pd.Series()),
                    self.current_provider_metrics.get('rend', pd.Series())
                )
        
                start_index = self.current_page * 5
                end_index = min(start_index + 5, len(self.all_provider_service_types))
                service_types_on_page = self.all_provider_service_types[start_index:end_index]
        
                with self.out:
                    clear_output(wait=True)
                    if not service_types_on_page:
                        print("No hay tipos de servicio para mostrar en esta página.")
                        self.evaluation_form_container.children = []
                        display(self.prov_sel)
                        display(self.generar_resumen_btn)
                        return
        
                    form = self.crear_widgets_evaluacion(mttr, mtbf, disp, rend, service_types_on_page)
        
                    prev_button = widgets.Button(description="Anterior")
                    next_button = widgets.Button(description="Siguiente")
        
                    prev_button.on_click(lambda b: self.navigate_evaluation_pages(-1))
                    next_button.on_click(lambda b: self.navigate_evaluation_pages(1))
        
                    prev_button.disabled = self.current_page == 0
                    next_button.disabled = end_index >= len(self.all_provider_service_types)
        
                    nav_buttons_box = widgets.HBox([prev_button, next_button])
        
                    self.evaluation_form_container.children = [self.service_type_mapping_display, form, nav_buttons_box]
                    display(self.prov_sel, self.evaluation_form_container, self.generar_resumen_btn)
        
                    # Plotting only if there's data to plot
                    if not rend.empty:
                        self.graficar_rendimiento(rend)
                    else:
                        print("\nNo hay datos de rendimiento para graficar.")
        
                    if not mttr.empty or not mtbf.empty or not disp.empty:
                        self.graficar_resumen(mttr, mtbf, disp)
                    else:
                        print("\nNo hay datos de MTTR, MTBF o Disponibilidad para graficar.")
        
        
            def navigate_evaluation_pages(self, direction):
                self.current_page += direction
                self.update_evaluation_display()
        
            def generar_resumen_evaluacion(self, btn):
                with self.out:
                    clear_output()
                    print("Generando resumen de evaluación...")
        
                    if not self.all_evaluation_widgets_map:
                        print("No hay evaluaciones para resumir. Selecciona un proveedor y completa las evaluaciones.")
                        display(self.prov_sel)
                        return
        
                    unique_service_types = sorted(list(set([key[2] for key in self.all_evaluation_widgets_map.keys()])))
        
                    all_categories = sorted(list(set([p[0] for p in preguntas])))
                    category_service_scores = {cat: {st: 0 for st in unique_service_types} for cat in all_categories}
        
                    service_type_totals = {st: 0 for st in unique_service_types}
        
                    for (cat, q_text, st_original), widget in self.all_evaluation_widgets_map.items():
                        try:
                            if isinstance(widget, widgets.Dropdown):
                                score = int(widget.value)
                            elif isinstance(widget, widgets.Label):
                                score = int(widget.value)
                            else:
                                score = 0
        
                            if cat not in category_service_scores: # Defensive check
                                category_service_scores[cat] = {st: 0 for st in unique_service_types}
        
                            category_service_scores[cat][st_original] += score
                            service_type_totals[st_original] += score
                        except ValueError:
                            print(f"Advertencia: No se pudo obtener la puntuación para la pregunta '{q_text}' del servicio '{st_original}'. Asegúrate de que el valor sea numérico.")
                        except KeyError:
                            print(f"Advertencia: Categoría '{cat}' o tipo de servicio '{st_original}' no encontrado en la estructura de puntuación.")
        
                    summary_df_calificacion = pd.DataFrame.from_dict(category_service_scores, orient='index')
                    summary_df_calificacion.index.name = 'Categoría'
                    summary_df_calificacion.loc['Total por Calificación de Servicio'] = pd.Series(service_type_totals)
        
                    # Ensure all_provider_service_types is populated before mapping
                    if not self.all_provider_service_types:
                        self.all_provider_service_types = sorted(self.df['TIPO DE SERVICIO'].dropna().unique().tolist())
        
        
                    service_type_display_names_cal = {st: f"Servicio {self.all_provider_service_types.index(st) + 1} ({st})"
                                                      for st in unique_service_types if st in self.all_provider_service_types}
                    # Add any unique_service_types not found in all_provider_service_types (shouldn't happen if logic is correct)
                    for st in unique_service_types:
                        if st not in self.all_provider_service_types:
                             service_type_display_names_cal[st] = f"Servicio Nuevo ({st})"
        
                    summary_df_calificacion = summary_df_calificacion.rename(columns=service_type_display_names_cal)
                    summary_df_calificacion.columns.name = 'Tipo de Servicio'
        
                    prov_name = self.prov_sel.value.replace(" ", "_").replace("/", "-")
                    output_filename = f"resumen_evaluacion_{prov_name}.xlsx"
        
                    try:
                        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                            if self.summary_servicio_global_for_export is not None and not self.summary_servicio_global_for_export.empty:
                                self.summary_servicio_global_for_export.to_excel(writer, sheet_name='Resumen_Servicio', index=False)
                            else:
                                print("Advertencia: No se encontró la tabla de resumen por Tipo de Servicio para exportar o está vacía.")
        
                            if self.resumen_equipo_global_for_export is not None and not self.resumen_equipo_global_for_export.empty:
                                self.resumen_equipo_global_for_export.to_excel(writer, sheet_name='Resumen_Equipo', index=False)
                            else:
                                print("Advertencia: No se encontró la tabla de resumen por Equipo para exportar o está vacía.")
        
                            if not summary_df_calificacion.empty:
                                summary_df_calificacion.to_excel(writer, sheet_name='Resumen_Calificacion')
                            else:
                                 print("Advertencia: No se encontró la tabla de resumen de Calificación para exportar o está vacía.")
        
        
                        print(f"Resúmenes exportados a: {output_filename}")
                        print("\nPara continuar, selecciona un proveedor o navega por las opciones.")
        
                    except Exception as e:
                        print(f"Error al exportar a Excel: {e}")
                        print("Asegúrate de tener la librería 'xlsxwriter' instalada (pip install xlsxwriter) y de que el archivo no esté abierto en otra aplicación.")
        
                    # Re-display the UI after generating summary
                    display(self.prov_sel, self.evaluation_form_container, self.generar_resumen_btn)
                    # Re-plot if data is available
                    if 'rend' in self.current_provider_metrics and not self.current_provider_metrics['rend'].empty:
                        self.graficar_rendimiento(self.current_provider_metrics['rend'])
                    if ('mttr' in self.current_provider_metrics and not self.current_provider_metrics['mttr'].empty) or \
                       ('mtbf' in self.current_provider_metrics and not self.current_provider_metrics['mtbf'].empty) or \
                       ('disp' in self.current_provider_metrics and not self.current_provider_metrics['disp'].empty):
                        self.graficar_resumen(self.current_provider_metrics['mttr'], self.current_provider_metrics['mtbf'], self.current_provider_metrics['disp'])
        
        
            def on_proveedor_change(self, change):
                if change['type'] == 'change' and change['name'] == 'value':
                    with self.out:
                        clear_output()
                        prov = change['new']
                        if not prov:
                            print("Selecciona un proveedor")
                            return
        
                        if prov == "Todos":
                            sub = self.df.copy()
                        else:
                            sub = self.df[self.df['PROVEEDOR'] == prov].copy() # Ensure we're working with a copy
        
                        if sub.empty:
                            print(f"No hay datos disponibles para el proveedor '{prov}'.")
                            self.current_provider_metrics = {}
                            self.all_provider_service_types = []
                            self.all_evaluation_widgets_map = {}
                            self.summary_servicio_global_for_export = None
                            self.resumen_equipo_global_for_export = None
                            self.update_evaluation_display()
                            return
        
                        cnt_s, cost_s, mttr_s, mtbf_s, disp_s, rend_s = calcular_indicadores_servicio(sub)
                        self.current_provider_metrics = {'cnt': cnt_s, 'cost': cost_s, 'mttr': mttr_s, 'mtbf': mtbf_s, 'disp': disp_s, 'rend': rend_s}
        
                        self.all_provider_service_types = sorted(sub['TIPO DE SERVICIO'].dropna().unique().tolist())
                        self.all_evaluation_widgets_map = {}
                        self.current_page = 0
        
                        # Prepare data for export, ensuring to handle potentially empty Series
                        self.summary_servicio_global_for_export = pd.DataFrame({
                            'Cantidad de Avisos': cnt_s,
                            'Costo Total': cost_s,
                            'Disponibilidad (%)': disp_s.round(2) if not disp_s.empty else np.nan,
                            'MTTR (hrs)': mttr_s.round(2) if not mttr_s.empty else np.nan,
                            'MTBF (hrs)': mtbf_s.round(2) if not mtbf_s.empty else np.nan,
                            'Rendimiento': rend_s
                        }).reset_index().rename(columns={'index': 'TIPO DE SERVICIO'}).sort_values('Disponibilidad (%)', ascending=False)
                        # Fill N/A in numeric columns that might have appeared due to empty series
                        for col in ['Disponibilidad (%)', 'MTTR (hrs)', 'MTBF (hrs)']:
                            if col in self.summary_servicio_global_for_export.columns:
                                self.summary_servicio_global_for_export[col] = self.summary_servicio_global_for_export[col].fillna(0)
        
        
                        cnt_e, cost_e, mttr_e, mtbf_e, disp_e, rend_e, desc_category_e = calcular_indicadores_equipo(sub) # Added desc_category_e
                        self.resumen_equipo_global_for_export = pd.DataFrame({
                            'Avisos': cnt_e,
                            'Costo total': cost_e,
                            'MTTR': mttr_e.round(2) if not mttr_e.empty else np.nan,
                            'MTBF': mtbf_e.round(2) if not mtbf_e.empty else np.nan,
                            'Disponibilidad (%)': disp_e.round(2) if not disp_e.empty else np.nan,
                            'Rendimiento': rend_e,
                            'Categoría de Descripción': desc_category_e # Added this line
                        }).reset_index().rename(columns={'index': 'Denominacion'})
                        for col in ['Disponibilidad (%)', 'MTTR', 'MTBF']:
                            if col in self.resumen_equipo_global_for_export.columns:
                                self.resumen_equipo_global_for_export[col] = self.resumen_equipo_global_for_export[col].fillna(0)
        
        
                        self.update_evaluation_display()
        
        
            def graficar_rendimiento(self, rend):
                if rend.empty:
                    print("No hay datos de rendimiento para graficar.")
                    return
        
                # Filter out NaN values before counting and plotting
                rend_filtered = rend.dropna()
                if rend_filtered.empty:
                    print("No hay datos de rendimiento válidos para graficar después de eliminar NaNs.")
                    return
        
                rend_counts = rend_filtered.value_counts().reindex(['Alto', 'Medio', 'Bajo'], fill_value=0)
        
                # Only plot if there's actual data to display in the pie chart
                if rend_counts.sum() > 0:
                    colores = ['#66bb6a', '#ffee58', '#ef5350']
                    # Filter colors to match available data points
                    active_colors = [color for i, color in enumerate(colores) if rend_counts.iloc[i] > 0]
                    labels = [f'{idx} ({val}%)' for idx, val in rend_counts.items() if val > 0] # Include percentage in labels
                    sizes = [val for val in rend_counts.values if val > 0]
        
                    if not sizes: # If all values are zero after reindexing
                        print("No hay datos de rendimiento para mostrar en el gráfico de pastel.")
                        return
        
                    plt.figure(figsize=(6, 6))
                    wedges, texts, autotexts = plt.pie(sizes, autopct='%1.1f%%', colors=active_colors, startangle=90, counterclock=False, pctdistance=0.85)
                    plt.title('Distribución Rendimiento por Tipo de Servicio')
                    plt.ylabel('')
                    plt.axis('equal') # Equal aspect ratio ensures that pie is drawn as a circle.
        
                    # Add labels if you want them on the outside
                    # plt.legend(wedges, labels, title="Rendimiento", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
        
                    plt.tight_layout()
                    plt.show()
                else:
                    print("No hay datos válidos de rendimiento para graficar.")
        
        
            def graficar_resumen(self, mttr, mtbf, disp):
                # Filter out empty Series before plotting
                plots_to_make = 0
                if not mttr.dropna().empty: plots_to_make += 1
                if not mtbf.dropna().empty: plots_to_make += 1
                if not disp.dropna().empty: plots_to_make += 1
        
                if plots_to_make == 0:
                    print("No hay datos de MTTR, MTBF o Disponibilidad para graficar.")
                    return
        
                fig, axs = plt.subplots(1, plots_to_make, figsize=(5 * plots_to_make, 4))
                # Ensure axs is iterable even if only one plot is made
                if plots_to_make == 1:
                    axs = [axs]
        
                plot_idx = 0
                if not mttr.dropna().empty:
                    sns.histplot(mttr.dropna(), bins=10, kde=True, ax=axs[plot_idx], color='skyblue')
                    axs[plot_idx].set_title('MTTR Promedio por Servicio (hrs)')
                    plot_idx += 1
                if not mtbf.dropna().empty:
                    sns.histplot(mtbf.dropna(), bins=10, kde=True, ax=axs[plot_idx], color='lightgreen')
                    axs[plot_idx].set_title('MTBF Promedio por Servicio (hrs)')
                    plot_idx += 1
                if not disp.dropna().empty:
                    sns.histplot(disp.dropna(), bins=10, kde=True, ax=axs[plot_idx], color='salmon')
                    axs[plot_idx].set_title('Disponibilidad Promedio por Servicio (%)')
                    plot_idx += 1
        
                plt.tight_layout()
                plt.show()
        
            def get_ui(self):
                return self.ui
        
        # --- COST AND NOTICE ANALYSIS WIDGETS AND LOGIC (from second code) ---
        class Pagination:
            def __init__(self, total_items, items_per_page=20):
                self.total_items = total_items
                self.items_per_page = items_per_page
                self.page = 0
                self.max_page = max(0, (total_items - 1) // items_per_page)
        
            def get_slice(self):
                start = self.page * self.items_per_page
                end = start + self.items_per_page
                return slice(start, end)
        
            def next(self):
                if self.page < self.max_page:
                    self.page += 1
        
            def prev(self):
                if self.page > 0:
                    self.page -= 1
        
        class CostosAvisosApp:
            def __init__(self, df):
                self.df = df
                self.EJECUTANTE_COL_NAME_NORMALIZED = "denominacion_ejecutante"
                self.COL_COSTOS_NORMALIZED = "costes_totreales"
                self.CP_COL_NAME_NORMALIZED = "codigo_postal"
                self.DESCRIPTION_COL_NAME_NORMALIZED = "descripcion"
        
                # Ensure filters use only valid options from the DataFrame
                self.ejecutantes = sorted(self.df[self.EJECUTANTE_COL_NAME_NORMALIZED].dropna().unique().tolist())
                self.cps = sorted(self.df[self.CP_COL_NAME_NORMALIZED].dropna().unique().tolist())
                self.años = sorted(self.df["año"].dropna().unique().tolist())
                self.meses = sorted(self.df["mes"].dropna().unique().tolist(), key=lambda x: pd.to_datetime(x, format="%B").month)
        
                self.ejecutante_w = widgets.SelectMultiple(options=self.ejecutantes, value=tuple(self.ejecutantes), description="Ejecutante")
                self.cp_w = widgets.SelectMultiple(options=self.cps, value=tuple(self.cps), description="Código postal")
                self.año_w = widgets.Dropdown(options=["Todos"] + self.años, description="Año")
                self.mes_w = widgets.Dropdown(options=["Todos"] + self.meses, description="Mes")
        
                self.opciones_menu = {
                    "Costos por ejecutante": (self.EJECUTANTE_COL_NAME_NORMALIZED, self.COL_COSTOS_NORMALIZED, "costos"),
                    "Avisos por ejecutante": (self.EJECUTANTE_COL_NAME_NORMALIZED, None, "avisos"),
                    "Costos por objeto técnico": ("denominacion_de_objeto_tecnico", self.COL_COSTOS_NORMALIZED, "costos"),
                    "Avisos por objeto técnico": ("denominacion_de_objeto_tecnico", None, "avisos"),
                    "Costos por texto código acción": ("texto_codigo_accion", self.COL_COSTOS_NORMALIZED, "costos"),
                    "Avisos por texto código acción": ("texto_codigo_accion", None, "avisos"),
                    "Costos por texto de acción": ("texto_de_accion", self.COL_COSTOS_NORMALIZED, "costos"),
                    "Avisos por texto de acción": ("texto_de_accion", None, "avisos"),
                    "Costos por tipo de servicio": ("tipo_de_servicio", self.COL_COSTOS_NORMALIZED, "costos"),
                    "Avisos por tipo de servicio": ("tipo_de_servicio", None, "avisos"),
                    "Costos por categoría de descripción": ("description_category", self.COL_COSTOS_NORMALIZED, "costos"),
                    "Avisos por categoría de descripción": ("description_category", None, "avisos"),
                }
        
                self.paginacion_actual = None
                self.agrupacion_actual = None
                self.opcion_actual = None
        
                self.btn_prev = widgets.Button(description="← Página anterior")
                self.btn_next = widgets.Button(description="Página siguiente →")
                self.menu_dropdown = widgets.Dropdown(
                    options=list(self.opciones_menu.keys()),
                    description="Ver:",
                    layout=widgets.Layout(width="400px")
                )
                self.output_area = widgets.Output()
        
                # Observe changes in filters to re-render the content
                self.ejecutante_w.observe(self.on_filter_change, names='value')
                self.cp_w.observe(self.on_filter_change, names='value')
                self.año_w.observe(self.on_filter_change, names='value')
                self.mes_w.observe(self.on_filter_change, names='value')
        
                self.btn_prev.on_click(self.on_prev_clicked)
                self.btn_next.on_click(self.on_next_clicked)
                self.menu_dropdown.observe(self.on_menu_change, names='value')
        
                self.filtros_ui = widgets.VBox([self.ejecutante_w, self.cp_w, self.año_w, self.mes_w])
                self.botones_ui = widgets.HBox([self.btn_prev, self.btn_next])
                self.ui = widgets.VBox([self.filtros_ui, self.menu_dropdown, self.botones_ui, self.output_area])
        
                # Initial display
                self.on_filter_change() # Trigger initial display of data after filters are set
        
            def mostrar_rangos_respuesta(self, btn):
                with self.out:
                    clear_output()
                    print("--- Rangos de Respuesta para cada Pregunta ---")
                    print("\n**Escala General:**")
                    print("2: Sobresaliente")
                    print("1: Bueno")
                    print("0: Indiferente")
                    print("-1: Malo")
                    print("\n--- Preguntas y sus rangos ---")
        
                    rangos_detallados = {
                        "Calidad": {
                            "¿Las soluciones propuestas son coherentes con el diagnóstico y causa raíz del problema?": {
                                2: "Total coherencia con el diagnóstico y causas identificadas",
                                1: "Coherencia razonable, con pequeños ajustes necesarios",
                                0: "Cumple con lo básico, pero con limitaciones relevantes",
                                -1: "No guarda coherencia o es deficiente respecto al diagnóstico"
                            },
                            "¿El trabajo entregado tiene materiales nuevos, originales y de marcas reconocidas?": {
                                2: "Todos los materiales son nuevos, originales y de marcas reconocidas",
                                1: "La mayoría de los materiales cumplen esas condiciones",
                                0: "Algunos materiales no son nuevos o no están certificados",
                                -1: "Materiales genéricos, usados o sin respaldo de marca"
                            },
                            "¿Cuenta con acabados homogéneos, limpios y pulidos?": {
                                2: "Acabados uniformes, bien presentados y profesionales",
                                1: "En general, los acabados son aceptables y limpios",
                                0: "Presenta inconsistencias notorias en algunos acabados",
                                -1: "Acabados descuidados, sucios o sin terminación adecuada"
                            },
                            "¿El trabajo entregado corresponde completamente con lo contratado?": {
                                2: "Cumple en su totalidad con lo contratado y acordado",
                                1: "Cumple en gran parte con lo contratado, con mínimos desvíos",
                                0: "Cumple con los requisitos mínimos establecidos",
                                -1: "No corresponde con lo contratado o presenta deficiencias importantes"
                            },
                            "¿La facturación refleja correctamente lo ejecutado y acordado?": {
                                2: "Facturación precisa, sin errores y con toda la información requerida",
                                1: "Facturación con pequeños errores que no afectan el control",
                                0: "Facturación con errores importantes (por ejemplo, precios)",
                                -1: "Facturación incorrecta, incompleta o que requiere ser repetida"
                            }
                        },
                        "Oportunidad": {
                            "¿La entrega de cotizaciones fue oportuna, según el contrato?": {
                                2: "Siempre entrega cotizaciones en los tiempos establecidos",
                                1: "Generalmente cumple con los plazos establecidos",
                                0: "A veces entrega fuera del tiempo estipulado",
                                -1: "Frecuentemente incumple los tiempos o no entrega"
                            },
                            "¿El reporte del servicio fue entregado oportunamente, según el contrato?": {
                                2: "Siempre entrega los reportes a tiempo, según lo acordado",
                                1: "Entrega los reportes con mínimos retrasos",
                                0: "Entrega con demoras ocasionales",
                                -1: "Entrega tardía constante o no entrega"
                            },
                            "¿Cumple las fechas y horas programadas para los trabajos, según el contrato?": {
                                2: "Puntualidad absoluta en fechas y horarios de ejecución",
                                1: "Puntualidad general con excepciones menores",
                                0: "Cumplimiento parcial o con retrasos frecuentes",
                                -1: "Incumplimiento reiterado de horarios o fechas"
                            },
                            "¿Responde de forma efectiva ante eventualidades emergentes, según el contrato?": {
                                2: "Respuesta inmediata y eficaz ante cualquier eventualidad",
                                1: "Respuesta adecuada en la mayoría de los casos",
                                0: "Respuesta tardía o poco efectiva en varias situaciones",
                                -1: "No responde adecuadamente o ignora emergencias"
                            },
                            "¿Soluciona rápidamente reclamos o inquietudes por garantía, según el contrato?": {
                                2: "Soluciona siempre con rapidez y eficacia",
                                1: "Responde satisfactoriamente en la mayoría de los casos",
                                0: "Respuesta variable, con demoras ocasionales",
                                -1: "Soluciones lentas o sin resolver adecuadamente"
                            },
                            "¿Dispone de los repuestos requeridos en los tiempos necesarios, según el contrato?": {
                                2: "Siempre cuenta con repuestos disponibles en el tiempo requerido",
                                1: "Generalmente cumple con la disponibilidad de repuestos",
                                0: "Disponibilidad intermitente o con retrasos",
                                -1: "No garantiza disponibilidad o presenta retrasos constantes"
                            },
                            "¿Entrega las facturas en los tiempos convenidos, según el contrato?": {
                                2: "Entrega siempre puntual de facturas",
                                1: "Entrega generalmente puntual con pocas excepciones",
                                0: "Entrega ocasionalmente fuera del tiempo acordado",
                                -1: "Entrega tarde con frecuencia o no entrega"
                            }
                        },
                        "Precio": {
                            "¿Los precios ofrecidos para equipos son competitivos respecto al mercado?": {
                                2: "Muy por debajo del precio promedio de mercado",
                                1: "Por debajo del promedio de mercado",
                                0: "Igual al promedio de mercado",
                                -1: "Por encima del promedio de mercado"
                            },
                            "¿Los precios ofrecidos para repuestos son competitivos respecto al mercado?": {
                                2: "Muy por debajo del precio promedio de mercado",
                                1: "Por debajo del promedio de mercado",
                                0: "Igual al promedio de mercado",
                                -1: "Por encima del promedio de mercado"
                            },
                            "Facilita llegar a una negociación (precios)": {
                                2: "Siempre está dispuesto a negociar de manera flexible",
                                1: "En general muestra disposición al diálogo",
                                0: "Ocasionalmente permite negociar",
                                -1: "Poco o nada dispuesto a negociar"
                            },
                            "Pone en consideración contratos y trabajos adjudicados en el último periodo de tiempo": {
                                2: "Siempre toma en cuenta la relación comercial previa",
                                1: "Generalmente considera trabajos anteriores",
                                0: "Solo ocasionalmente lo toma en cuenta",
                                -1: "No muestra continuidad ni reconocimiento de antecedentes"
                            },
                            "¿Los precios ofrecidos para mantenimientos son competitivos respecto al mercado?": {
                                2: "Muy por debajo del precio promedio de mercado",
                                1: "Por debajo del promedio de mercado",
                                0: "Igual al promedio de mercado",
                                -1: "Por encima del promedio de mercado"
                            },
                            "¿Los precios ofrecidos para insumos son competitivos respecto al mercado?": {
                                2: "Muy por debajo del precio promedio de mercado",
                                1: "Por debajo del promedio de mercado",
                                0: "Igual al promedio de mercado",
                                -1: "Por encima del promedio de mercado"
                            }
                        },
                        "Postventa": {
                            "¿Tiene disposición y actitud de servicio frente a solicitudes?": {
                                2: "Atención proactiva y excelente actitud de servicio",
                                1: "Buena actitud y disposición general",
                                0: "Actitud pasiva o limitada ante las solicitudes",
                                -1: "Falta de disposición o actitudes negativas"
                            },
                            "¿Conoce necesidades y ofrece alternativas adecuadas?": {
                                2: "Conocimiento profundo del cliente y propuestas adecuadas",
                                1: "Buen conocimiento y alternativas en general adecuadas",
                                0: "Soluciones parcialmente adecuadas",
                                -1: "No se adapta a las necesidades o propone soluciones inadecuadas"
                            },
                            "¿Realiza seguimiento a los resultados de los trabajos?": {
                                2: "Hace seguimiento sistemático y detallado",
                                1: "Realiza seguimiento general adecuado",
                                0: "Seguimiento ocasional o no documentado",
                                -1: "No realiza seguimiento posterior"
                            },
                            "¿Ofrece capacitaciones para el manejo de los equipos?": {
                                2: "Capacitaciones constantes y bien estructuradas",
                                1: "Capacitaciones ocasionales pero útiles",
                                0: "Capacitaciones mínimas o informales",
                                -1: "No ofrece capacitaciones"
                            },
                            "¿Los métodos de capacitación ofrecidos son efectivos y adecuados?": {
                                2: "Métodos claros, efectivos y adaptados al usuario",
                                1: "Métodos generalmente útiles y comprensibles",
                                0: "Métodos poco claros o limitados",
                                -1: "Métodos ineficaces o mal estructurados"
                            }
                        },
                        "Desempeño técnico": {
                            "Disponibilidad promedio (%)": {
                                2: "Disponibilidad >= 98%",
                                1: "75% <= Disponibilidad < 98%",
                                0: "Disponibilidad < 75%"
                            },
                            "MTTR promedio (hrs)": {
                                2: "MTTR <= 5 hrs",
                                1: "5 hrs < MTTR <= 20 hrs",
                                0: "MTTR > 20 hrs"
                            },
                            "MTBF promedio (hrs)": {
                                2: "MTBF > 1000 hrs",
                                1: "100 hrs <= MTBF <= 1000 hrs",
                                0: "MTBF < 100 hrs"
                            },
                            "Rendimiento promedio equipos": {
                                2: "Rendimiento 'Alto' (Disponibilidad >= 90%)",
                                1: "Rendimiento 'Medio' (75% <= Disponibilidad < 90%)",
                                0: "Rendimiento 'Bajo' (Disponibilidad < 75%)"
                            }
                        }
                    }
        
                    for cat, texto, escala in preguntas:
                        print(f"\n[{cat}] {texto}:")
                        if escala == "auto":
                            # For 'auto' scales, get the specific ranges from the detailed dictionary
                            if cat in rangos_detallados and texto in rangos_detallados[cat]:
                                for val, desc in rangos_detallados[cat][texto].items():
                                    print(f"   - {val}: {desc}")
                            else:
                                print("   (Rangos definidos automáticamente por el sistema)")
                        else:
                            # For fixed scales, use the detailed dictionary
                            if cat in rangos_detallados and texto in rangos_detallados[cat]:
                                for val, desc in rangos_detallados[cat][texto].items():
                                    print(f"   - {val}: {desc}")
                            else:
                                print(f"   Rangos: {escala}") # Fallback if not found in detailed map, should not happen with full mapping
        
                    print("\nPara continuar, selecciona un proveedor o navega por las opciones.")
                    # Ensure evaluation_form_container and generar_resumen_btn are displayed after showing ranges
                    display(self.prov_sel, self.mostrar_rangos_btn, self.evaluation_form_container, self.generar_resumen_btn)
        
            def get_ui(self):
            # Modifica esta línea para incluir el botón de rangos
                 self.ui = widgets.VBox([self.prov_sel, self.mostrar_rangos_btn, self.out])
                return self.ui
        
            def formato_coste(self, valor):
                if pd.isna(valor):
                    return "$ N/A"
                return "$ {:,.0f}COP".format(valor).replace(",", ".") # Removed division by 1 for direct formatting
        
            def custom_cost_formatter(self, x, pos):
                if x >= 1e9:
                    return f"${x*1e-9:.0f}B"
                elif x >= 1e6:
                    return f"${x*1e-6:.0f}M"
                elif x >= 1e3:
                    return f"${x*1e-3:.0f}K"
                else:
                    return f"${x:.0f}"
        
            def plot_points_with_labels(self, data, title, xlabel="", ylabel="", color="skyblue", rotation=45):
                if data.empty:
                    with self.output_area:
                        print(f"No hay datos para graficar: {title}")
                    return
        
                plt.figure(figsize=(12, 6))
                x_labels = data.index.astype(str) # Ensure x_labels are strings
                x_positions = np.arange(len(data)) # Use numpy for positions
        
                plt.scatter(x_positions, data.values, color=color, zorder=2)
        
                # Add labels on points
                for i, txt in enumerate(data.values):
                    if "avisos" in title.lower():
                        label_text = f"{int(txt):,}".replace(",", ".")
                    else: # For costs
                        label_text = self.custom_cost_formatter(txt, None)
                    plt.text(x_positions[i], txt, label_text, ha='center', va='bottom', fontsize=9, color='black')
        
                plt.title(title)
                plt.xlabel(xlabel)
                plt.ylabel(ylabel)
                plt.xticks(x_positions, x_labels, rotation=rotation, ha="right")
        
                if "costos" in title.lower():
                    plt.gca().yaxis.set_major_formatter(mticker.FuncFormatter(self.custom_cost_formatter))
        
                plt.grid(axis='y', linestyle='--', alpha=0.7)
                plt.tight_layout()
                plt.show()
        
            def filtrar_datos(self):
                # Ensure that filters are applied to the correct column names (normalized)
                filtered_df = self.df.copy()
        
                if self.ejecutante_w.value:
                    filtered_df = filtered_df[filtered_df[self.EJECUTANTE_COL_NAME_NORMALIZED].isin(list(self.ejecutante_w.value))]
                if self.cp_w.value:
                    filtered_df = filtered_df[filtered_df[self.CP_COL_NAME_NORMALIZED].isin(list(self.cp_w.value))]
        
                if self.año_w.value != "Todos":
                    filtered_df = filtered_df[filtered_df["año"] == self.año_w.value]
                if self.mes_w.value != "Todos":
                    filtered_df = filtered_df[filtered_df["mes"] == self.mes_w.value]
                return filtered_df
        
            def mostrar_pagina(self, grouped_data, opcion):
                with self.output_area:
                    clear_output(wait=True)
                    if self.paginacion_actual is None or grouped_data.empty:
                        print("No hay datos para mostrar.")
                        return
        
                    current_slice = self.paginacion_actual.get_slice()
                    data_to_display = grouped_data[current_slice]
        
                    col, col_cost, tipo = self.opciones_menu[opcion]
                    title = f"{opcion} - Página {self.paginacion_actual.page + 1} de {self.paginacion_actual.max_page + 1}"
        
                    if data_to_display.empty:
                        print("No hay datos para mostrar en esta página con los filtros actuales.")
                        return
        
                    # Display table
                    if tipo == "costos":
                        html_table = data_to_display.apply(self.formato_coste).to_frame().to_html(classes='table table-striped')
                        display(HTML(f'<style> .table-striped tbody tr:nth-of-type(odd) {{ background-color: rgba(0,0,0,.05); }} </style>{html_table}'))
                    else:
                        html_table = data_to_display.to_frame().to_html(classes='table table-striped')
                        display(HTML(f'<style> .table-striped tbody tr:nth-of-type(odd) {{ background-color: rgba(0,0,0,.05); }} </style>{html_table}'))
        
                    # Display plot
                    self.plot_points_with_labels(data_to_display, title, xlabel=col, ylabel="Costo Total (COP)" if tipo == "costos" else "Número de Avisos")
        
            def on_filter_change(self, change=None):
                # Reset to the first page when filters change
                self.mostrar_datos(self.menu_dropdown.value, pagina=0)
        
            def mostrar_datos(self, opcion, pagina=0):
                self.opcion_actual = opcion
                dff = self.filtrar_datos()
        
                if dff.empty:
                    with self.output_area:
                        clear_output()
                        print("No hay datos que coincidan con los filtros seleccionados.")
                        self.paginacion_actual = None
                        self.agrupacion_actual = pd.Series() # Ensure it's an empty Series for consistency
                    return
        
                col, col_cost, tipo = self.opciones_menu[opcion]
        
                if col not in dff.columns:
                    with self.output_area:
                        clear_output()
                        print(f"Error: La columna '{col}' no se encontró en los datos filtrados. "
                              f"Columnas disponibles: {dff.columns.tolist()}")
                    self.paginacion_actual = None
                    self.agrupacion_actual = pd.Series()
                    return
        
                if tipo == "costos":
                    if col_cost not in dff.columns:
                        with self.output_area:
                            clear_output()
                            print(f"Error: La columna de costos '{col_cost}' no se encontró en los datos filtrados.")
                        self.paginacion_actual = None
                        self.agrupacion_actual = pd.Series()
                        return
                    self.agrupacion_actual = dff.groupby(col)[col_cost].sum().sort_values(ascending=False)
                else:
                    self.agrupacion_actual = dff[col].value_counts().sort_values(ascending=False)
        
                # Handle cases where grouped data might be empty after grouping/value_counts
                if self.agrupacion_actual.empty:
                    with self.output_area:
                        clear_output()
                        print(f"No hay datos para mostrar para la opción '{opcion}' con los filtros actuales.")
                        self.paginacion_actual = None
                    return
        
                self.paginacion_actual = Pagination(len(self.agrupacion_actual))
                self.paginacion_actual.page = pagina
        
                self.mostrar_pagina(self.agrupacion_actual, opcion)
        
            def on_prev_clicked(self, b):
                if self.paginacion_actual and self.agrupacion_actual is not None:
                    self.paginacion_actual.prev()
                    self.mostrar_pagina(self.agrupacion_actual, self.opcion_actual)
        
            def on_next_clicked(self, b):
                if self.paginacion_actual and self.agrupacion_actual is not None:
                    self.paginacion_actual.next()
                    self.mostrar_pagina(self.agrupacion_actual, self.opcion_actual)
        
            def on_menu_change(self, change):
                if change['type'] == 'change' and change['name'] == 'value':
                    self.mostrar_datos(change['new'])
        
            def get_ui(self):
                return self.ui
        
        # --- Main Application Controller ---
        class AppController:
            def __init__(self, df):
                self.df = df
                self.main_output = widgets.Output()
                self.evaluacion_app = EvaluacionProveedoresApp(self.df)
                self.costos_avisos_app = CostosAvisosApp(self.df)
        
                self.main_menu_button = widgets.Button(description="Volver al Menú Principal", layout=widgets.Layout(width='200px'))
                self.main_menu_button.on_click(self._display_main_menu)
        
            def _display_main_menu(self, b=None):
                with self.main_output:
                    clear_output()
                    print("Selecciona el tipo de análisis:")
                    btn_evaluacion = widgets.Button(description="Evaluación de Proveedores", layout=widgets.Layout(width='300px'))
                    btn_costos_avisos = widgets.Button(description="Análisis de Costos y Avisos", layout=widgets.Layout(width='300px'))
        
                    btn_evaluacion.on_click(self._start_evaluacion)
                    btn_costos_avisos.on_click(self._start_costos_avisos)
        
                    display(widgets.VBox([btn_evaluacion, btn_costos_avisos]))
        
            def _start_evaluacion(self, b):
                with self.main_output:
                    clear_output()
                    display(self.evaluacion_app.get_ui())
                    display(self.main_menu_button)
        
            def _start_costos_avisos(self, b):
                with self.main_output:
                    clear_output()
                    display(self.costos_avisos_app.get_ui())
                    display(self.main_menu_button)
        
            def run(self):
                display(self.main_output)
                self._display_main_menu()
        
        # --- Run the application ---
        app_controller = AppController(df)
        app_controller.run()
    except Exception as err:
        st.error(f"Ocurrió un error al procesar el archivo: {err}")
else:
    st.info("Por favor carga un archivo Excel para comenzar.")
