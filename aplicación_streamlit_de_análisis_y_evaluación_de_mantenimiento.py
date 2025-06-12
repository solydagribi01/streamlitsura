import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import matplotlib.pyplot as plt
import seaborn as sns
import re

# Suprimir advertencias de Matplotlib para una salida más limpia en Streamlit
# La opción 'deprecation.showPyplotGlobalUse' ha sido eliminada en versiones recientes de Streamlit,
# por lo tanto, la hemos removido para evitar errores.
# st.set_option('deprecation.showPyplotGlobalUse', False)

# --- Configuración de la aplicación ---
st.set_page_config(page_title="Herramienta de Análisis y Evaluación de Mantenimiento", layout="wide")
st.title("🔧 Herramienta Integral de Mantenimiento")

# --- Almacenamiento Global de Datos (para el estado de la sesión) ---
if 'df_processed' not in st.session_state:
    st.session_state.df_processed = None

# --- Helper para normalizar cadenas (usado para nombres de columnas) ---
def normalize_string(s):
    """Normaliza una cadena a minúsculas, con guiones bajos y sin caracteres especiales."""
    if pd.isna(s):
        return None
    s = str(s).strip().lower()
    # Reemplaza espacios y caracteres especiales con guiones bajos
    s = re.sub(r'[\s\.\(\)/%-]+', '_', s)
    # Elimina cualquier carácter no alfanumérico restante (excepto guiones bajos)
    s = re.sub(r'[^\w]+', '', s)
    # Maneja caracteres especiales del español
    s = s.replace('á', 'a').replace('é', 'e').replace('í', 'i').replace('ó', 'o').replace('ú', 'u').replace('ñ', 'n')
    # Elimina guiones bajos iniciales/finales que puedan resultar de los reemplazos
    s = s.strip('_')
    return s

# --- Sección de Carga y Procesamiento de Datos ---
with st.container():
    st.header("1. Procesamiento y Limpieza de Datos de Avisos")
    st.markdown("Carga un archivo `.xlsx` con las hojas IW29, IW39, IH08, IW65 y ZPM015 para su procesamiento inicial.")

    @st.cache_data
    def load_and_merge_data(file) -> pd.DataFrame:
        """
        Carga datos de las hojas especificadas en un archivo Excel y los fusiona.
        Realiza limpieza y cambio de nombre de columnas.
        """
        xls = pd.ExcelFile(file)
        iw29 = pd.read_excel(xls, sheet_name=0)
        iw39 = pd.read_excel(xls, sheet_name=1)
        ih08 = pd.read_excel(xls, sheet_name=2)
        iw65 = pd.read_excel(xls, sheet_name=3)
        zpm015 = pd.read_excel(xls, sheet_name=4)

        # Creando copias profundas para asegurar que las modificaciones no afecten los objetos originales
        dfs = [iw29.copy(), iw39.copy(), ih08.copy(), iw65.copy(), zpm015.copy()]
        iw29_clean, iw39_clean, ih08_clean, iw65_clean, zpm015_clean = dfs

        # Limpiar y normalizar encabezados de todas las copias de las hojas inmediatamente
        for df_temp in dfs:
            df_temp.columns = [normalize_string(col) for col in df_temp.columns]

        # DEBUG: Mostrar columnas de cada hoja después de la carga inicial y normalización
        st.info(f"Columnas de IW29 (normalizadas): {iw29_clean.columns.tolist()}")
        st.info(f"Columnas de IW39 (normalizadas): {iw39_clean.columns.tolist()}")
        st.info(f"Columnas de IH08 (normalizadas): {ih08_clean.columns.tolist()}")
        st.info(f"Columnas de IW65 (normalizadas): {iw65_clean.columns.tolist()}")
        st.info(f"Columnas de ZPM015 (normalizadas): {zpm015_clean.columns.tolist()}")
        
        # Asegurar que las columnas clave de fusión son de tipo cadena
        for df_temp in [iw29_clean, iw39_clean, ih08_clean, iw65_clean, zpm015_clean]:
            if 'aviso' in df_temp.columns: df_temp['aviso'] = df_temp['aviso'].astype(str)
            if 'equipo' in df_temp.columns: df_temp['equipo'] = df_temp['equipo'].astype(str)

        # Mapeo de columnas originales a nombres normalizados para evitar duplicados temporales antes de consolidar
        # y para asegurar que 'total_general_real' se convierta en 'costes_totreales'
        if 'denominacion_de_objeto_tecnico' in iw29_clean.columns:
            iw29_clean.rename(columns={'denominacion_de_objeto_tecnico': 'dot_iw29'}, inplace=True)
        if 'denominacion_de_objeto_tecnico' in iw39_clean.columns:
            iw39_clean.rename(columns={'denominacion_de_objeto_tecnico': 'dot_iw39'}, inplace=True)
        if 'total_general_real' in iw39_clean.columns:
            iw39_clean.rename(columns={'total_general_real': 'costes_totreales'}, inplace=True) # Renombrado directo para el costo principal
        if 'denominacion_de_objeto_tecnico' in ih08_clean.columns:
            ih08_clean.rename(columns={'denominacion_de_objeto_tecnico': 'dot_ih08'}, inplace=True)
        if 'texto' in ih08_clean.columns: # 'texto' en IH08 se convierte a 'texto_equipo'
            ih08_clean.rename(columns={'texto': 'texto_equipo'}, inplace=True)
        if 'denominacion_de_objeto_tecnico' in iw65_clean.columns:
            iw65_clean.rename(columns={'denominacion_de_objeto_tecnico': 'dot_iw65'}, inplace=True)
        
        # Si ZPM015 tiene 'denominacion_objeto' y no 'denominacion_de_objeto_tecnico' en su forma normalizada
        if 'denominacion_objeto' in zpm015_clean.columns and 'denominacion_de_objeto_tecnico' not in zpm015_clean.columns:
            zpm015_clean.rename(columns={'denominacion_objeto': 'dot_zpm015'}, inplace=True)
        elif 'denominacion_de_objeto_tecnico' in zpm015_clean.columns:
             zpm015_clean.rename(columns={'denominacion_de_objeto_tecnico': 'dot_zpm015'}, inplace=True) # Rename if already normalized


        # Define una función helper para obtener columnas de forma segura
        def get_cols_if_exist(df, cols_list):
            return [col for col in cols_list if col in df.columns]

        # Iniciar con IW29 como base
        final_df = iw29_clean.copy()

        # Fusión con IW39 (por 'aviso')
        iw39_merge_cols = get_cols_if_exist(iw39_clean, ['aviso', 'orden', 'costes_totreales', 'dot_iw39', 'fecha_entrada', 'texto_breve', 'tota_general_plan', 'centro_de_coste', 'denominacion_de_la_ubicacion_tecnica'])
        final_df = pd.merge(final_df, iw39_clean[iw39_merge_cols], on='aviso', how='left', suffixes=('', '_iw39'))

        # Fusión con IW65 (por 'aviso')
        iw65_merge_cols = get_cols_if_exist(iw65_clean, ['aviso', 'texto_codigo_accion', 'texto_de_accion', 'descripcion', 'nº_direccion', 'grupo_codigos', 'texto_grupo_accion', 'posicion', 'actividad', 'codigo_de_actividad', 'dot_iw65'])
        # Handle 'descripcion' overlap: keep IW29's first, then IW65's
        if 'descripcion_iw65_merge' in final_df.columns:
            final_df['descripcion'] = final_df['descripcion'].fillna(final_df['descripcion_iw65_merge'])
            final_df.drop(columns=['descripcion_iw65_merge'], errors='ignore', inplace=True)

        final_df = pd.merge(final_df, iw65_clean[iw65_merge_cols], on='aviso', how='left', suffixes=('', '_iw65'))
        
        # Fusión con IH08 (por 'equipo')
        ih08_merge_cols = get_cols_if_exist(ih08_clean, [
            'equipo', 'inic_garantia_prov', 'fin_garantia_prov', 'texto_equipo',
            'indicador_abc', 'dot_ih08', 'centro_de_coste', 'denominacion_de_la_ubicacion_tecnica',
            'fabricante_del_activo_fijo', 'denominacion_de_tipo', 'fabricante_numero_de_serie',
            'numero_de_inventario', 'numero_de_pieza_de_fabricante', 'numero_identificacion_tecnica',
            'cl_objeto_tecnico', 'valor_de_adquisicion', 'fecha_de_adquisicion',
            'tamano_dimension', 'existe_txt_expl'
        ])
        # Handle 'centro_de_coste' overlap: prefer IH08 if present, otherwise IW39
        if 'centro_de_coste_iw39' in final_df.columns and 'centro_de_coste' in ih08_clean.columns:
            ih08_clean.rename(columns={'centro_de_coste': 'centro_de_coste_ih08'}, inplace=True)
            ih08_merge_cols.remove('centro_de_coste')
            ih08_merge_cols.append('centro_de_coste_ih08')
        
        final_df = pd.merge(final_df, ih08_clean[ih08_merge_cols], on='equipo', how='left', suffixes=('', '_ih08'))
        
        # Fusión con ZPM015 (por 'equipo')
        zpm015_merge_cols = get_cols_if_exist(zpm015_clean, [
            'equipo', 'tipo_de_servicio', 'texto_cl_objeto', 'denom_ubic_tecnica',
            'ubicacion_tecnica', 'dot_zpm015', 'fabricante', 'tipo_de_equipo', 'ubicaciones_puntual'
        ])
        final_df = pd.merge(final_df, zpm015_clean[zpm015_merge_cols], on='equipo', how='left', suffixes=('', '_zpm015'))

        # Consolidar 'denominacion_de_objeto_tecnico' (DOT)
        # Priorizar: IW29 > IH08 > IW39 > IW65 > ZPM015 (denominacion_objeto)
        final_df['denominacion_de_objeto_tecnico'] = final_df['dot_iw29']
        if 'dot_ih08' in final_df.columns:
            final_df['denominacion_de_objeto_tecnico'].fillna(final_df['dot_ih08'], inplace=True)
        if 'dot_iw39' in final_df.columns:
            final_df['denominacion_de_objeto_tecnico'].fillna(final_df['dot_iw39'], inplace=True)
        if 'dot_iw65' in final_df.columns:
            final_df['denominacion_de_objeto_tecnico'].fillna(final_df['dot_iw65'], inplace=True)
        if 'dot_zpm015' in final_df.columns:
            final_df['denominacion_de_objeto_tecnico'].fillna(final_df['dot_zpm015'], inplace=True)


        # Consolidar 'costes_totreales': La principal fuente es IW39 ('total_general_real')
        # Ya ha sido renombrada directamente en iw39_clean a 'costes_totreales'
        # No se necesita consolidación adicional a menos que otra hoja tuviera una columna de "costo total" alternativa.
        # Si 'costes_totreales_iw39' existe, se usa, de lo contrario, se mantiene el actual (que puede ser del merge anterior o NaN)
        if 'costes_totreales_iw39' in final_df.columns:
            final_df['costes_totreales'] = final_df['costes_totreales'].fillna(final_df['costes_totreales_iw39'])
            final_df.drop(columns=['costes_totreales_iw39'], errors='ignore', inplace=True)
        
        # Consolidar 'ubicacion_tecnica'
        # Priorizar: IW29 > ZPM015 ('ubicacion_tecnica') > IH08 ('denominacion_de_la_ubicacion_tecnica')
        if 'ubicacion_tecnica_zpm015' in final_df.columns:
            final_df['ubicacion_tecnica'].fillna(final_df['ubicacion_tecnica_zpm015'], inplace=True)
            final_df.drop(columns=['ubicacion_tecnica_zpm015'], errors='ignore', inplace=True)
        if 'denominacion_de_la_ubicacion_tecnica_ih08' in final_df.columns:
            final_df['ubicacion_tecnica'].fillna(final_df['denominacion_de_la_ubicacion_tecnica_ih08'], inplace=True)
            final_df.drop(columns=['denominacion_de_la_ubicacion_tecnica_ih08'], errors='ignore', inplace=True)
        if 'denominacion_de_la_ubicacion_tecnica_iw39' in final_df.columns:
             final_df['ubicacion_tecnica'].fillna(final_df['denominacion_de_la_ubicacion_tecnica_iw39'], inplace=True)
             final_df.drop(columns=['denominacion_de_la_ubicacion_tecnica_iw39'], errors='ignore', inplace=True)
        if 'denominacion_de_la_ubicacion_tecnica_iw65' in final_df.columns:
            final_df['ubicacion_tecnica'].fillna(final_df['denominacion_de_la_ubicacion_tecnica_iw65'], inplace=True)
            final_df.drop(columns=['denominacion_de_la_ubicacion_tecnica_iw65'], errors='ignore', inplace=True)

        # Consolidar 'centro_de_coste'
        # Priorizar: IH08 > IW39 > ZPM015 (if 'centro_de_coste_zpm015' exists)
        if 'centro_de_coste_ih08' in final_df.columns:
            final_df['centro_de_coste'] = final_df['centro_de_coste'].fillna(final_df['centro_de_coste_ih08'])
            final_df.drop(columns=['centro_de_coste_ih08'], errors='ignore', inplace=True)
        if 'centro_de_coste_iw39' in final_df.columns:
            final_df['centro_de_coste'] = final_df['centro_de_coste'].fillna(final_df['centro_de_coste_iw39'])
            final_df.drop(columns=['centro_de_coste_iw39'], errors='ignore', inplace=True)
        if 'centro_de_coste_zpm015' in final_df.columns:
            final_df['centro_de_coste'] = final_df['centro_de_coste'].fillna(final_df['centro_de_coste_zpm015'])
            final_df.drop(columns=['centro_de_coste_zpm015'], errors='ignore', inplace=True)

        # Limpiar columnas temporales de DOT si no se han eliminado
        final_df.drop(columns=[col for col in ['dot_iw29', 'dot_ih08', 'dot_iw39', 'dot_iw65', 'dot_zpm015'] if col in final_df.columns], errors='ignore', inplace=True)
        
        # Define la lista completa de columnas esperadas en el DataFrame final
        # Asegurarse de que no haya duplicados y que los nombres sean los finales y normalizados.
        final_column_names = [
            'aviso', 'orden', 'fecha_de_aviso', 'hora_del_aviso', 'region', 'codigo_postal',
            'status_del_sistema', 'clase_de_aviso', 'texto_para_prioridad', 'status_de_usuario',
            'descripcion', 'duracion_de_parada', 'ubicacion_tecnica', 'cierre_por_fecha',
            'fin_deseado', 'fecha_de_pedido', 'indicador_abc', 'equipo',
            'denominacion_de_objeto_tecnico', 'denominacion_ejecutante', 'centro_de_coste',
            'costes_totreales', 'inic_garantia_prov', 'fin_garantia_prov', 'texto_equipo',
            'texto_codigo_accion', 'texto_de_accion', 'texto_grupo_accion',
            'tipo_de_servicio', 'fecha_entrada', 'texto_breve', 'tota_general_plan',
            'nº_direccion', 'grupo_codigos', 'posicion', 'actividad', 'codigo_de_actividad',
            'grupo_codigos_1', 'fabricante_del_activo_fijo', 'denominacion_de_tipo',
            'fabricante_numero_de_serie', 'numero_de_inventario', 'numero_de_pieza_de_fabricante',
            'numero_identificacion_tecnica', 'cl_objeto_tecnico', 'valor_de_adquisicion',
            'fecha_de_adquisicion', 'tamano_dimension', 'existe_txt_expl', 'texto_cl_objeto',
            'denom_ubic_tecnica', 'denominacion_objeto', 'fabricante', 'tipo_de_equipo',
            'ubicaciones_puntual', 'indicador' # 'indicador' not explicitly found in original lists but in previous code
        ]

        # Asegurar que todas las columnas en final_column_names estén presentes en final_df
        # Si no están, se añadirán con NaN.
        for col in final_column_names:
            if col not in final_df.columns:
                final_df[col] = np.nan
                # st.warning(f"La columna '{col}' no se encontró después de las fusiones. Se añadió con valores nulos.")

        # Filtrar a solo las columnas deseadas y en el orden especificado
        final_df = final_df[final_column_names].copy()

        st.info(f"Columnas del DataFrame después de fusiones y consolidación: {final_df.columns.tolist()}")
        st.info(f"Conteo de nulos en 'denominacion_de_objeto_tecnico' después de consolidación: {final_df['denominacion_de_objeto_tecnico'].isnull().sum()}")
        st.info(f"Conteo de nulos en 'costes_totreales' después de consolidación: {final_df['costes_totreales'].isnull().sum()}")

        return final_df

    @st.cache_data
    def process_data(df: pd.DataFrame) -> pd.DataFrame:
        """
        Realiza filtrado, ajuste de costos y normalización de columnas en el DataFrame.
        Asume que el DataFrame de entrada ya tiene columnas normalizadas.
        """
        # DEBUG: Mostrar columnas del DataFrame al inicio de process_data
        st.info(f"Columnas del DataFrame al inicio de process_data: {df.columns.tolist()}")

        # Las columnas esenciales ya deberían haber sido manejadas en load_and_merge_data,
        # pero mantenemos esta comprobación como un fallback y para asignar nombres amigables.
        
        # Filtrar 'PTBO' del Status del sistema
        if 'status_del_sistema' in df.columns:
            df = df[~df["status_del_sistema"].astype(str).str.contains("PTBO", case=False, na=False)].copy()
        else:
            st.warning("La columna 'status_del_sistema' no se encontró para el filtrado 'PTBO'.")

        # Ajustar costos duplicados por Aviso
        if 'aviso' in df.columns and 'costes_totreales' in df.columns:
            df['costes_totreales'] = pd.to_numeric(df['costes_totreales'], errors='coerce') # Asegurar que es numérico
            df['costes_totreales'] = df.groupby('aviso')['costes_totreales'].transform(
                lambda x: [x.iloc[0]] + [0]*(len(x)-1) if not x.empty else x
            )
        else:
            st.warning("Columnas 'aviso' o 'costes_totreales' no encontradas para el ajuste de costos duplicados.")

        # --- Asignar nombres más simples para uso posterior ---
        # Asegurarse de que estas columnas existan antes de asignarlas
        df['PROVEEDOR'] = df['denominacion_ejecutante'] if 'denominacion_ejecutante' in df.columns else np.nan
        df['COSTO'] = pd.to_numeric(df['costes_totreales'], errors='coerce') if 'costes_totreales' in df.columns else np.nan
        df['TIEMPO PARADA'] = pd.to_numeric(df['duracion_de_parada'], errors='coerce') if 'duracion_de_parada' in df.columns else np.nan
        df['EQUIPO_NUM'] = pd.to_numeric(df['equipo'], errors='coerce') if 'equipo' in df.columns else np.nan
        df['AVISO_NUM'] = pd.to_numeric(df['aviso'], errors='coerce') if 'aviso' in df.columns else np.nan
        df['TIPO DE SERVICIO'] = df['tipo_de_servicio'] if 'tipo_de_servicio' in df.columns else np.nan

        # --- Agregar 'HORA/ DIA' y 'DIAS/ AÑO' basadas en 'texto_equipo' ---
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

        if 'texto_equipo' in df.columns:
            df['HORARIO'] = df['texto_equipo'].astype(str).str.strip().str.upper()
            df['HORA/ DIA'] = df['HORARIO'].map(lambda x: horarios_dict.get(x, (np.nan, np.nan))[0])
            df['DIAS/ AÑO'] = df['HORARIO'].map(lambda x: horarios_dict.get(x, (np.nan, np.nan))[1])
            df['HORA/ DIA'] = pd.to_numeric(df['HORA/ DIA'], errors='coerce')
            df['DIAS/ AÑO'] = pd.to_numeric(df['DIAS/ AÑO'], errors='coerce')
        else:
            df['HORARIO'] = np.nan
            df['HORA/ DIA'] = np.nan
            df['DIAS/ AÑO'] = np.nan

        # Extraer año y mes
        if 'fecha_de_aviso' in df.columns:
            df['fecha_de_aviso'] = pd.to_datetime(df['fecha_de_aviso'], errors='coerce')
            df['año'] = df['fecha_de_aviso'].dt.year
            df['mes'] = df['fecha_de_aviso'].dt.strftime('%B')
        else:
            df['año'] = np.nan
            df['mes'] = np.nan

        # Categorizar 'descripcion'
        def categorize_description(description):
            if pd.isna(description):
                return "Sin Categoría"
            desc = str(description).lower()
            if "reparacion" in desc or "arreglo" in desc:
                return "Reparación"
            elif "mantenimiento" in desc:
                return "Mantenimiento"
            elif "inspeccion" in desc or "revision" in desc:
                return "Inspección/Revisión"
            else:
                return "Otro"
        
        if 'descripcion' in df.columns:
            df['description_category'] = df['descripcion'].apply(categorize_description)
        else:
            df['description_category'] = "Sin Categoría"

        # DEBUG: Mostrar el conteo de nulos para las columnas críticas antes de retornar
        st.info(f"Conteo de nulos en 'denominacion_de_objeto_tecnico' antes de retornar de process_data: {df['denominacion_de_objeto_tecnico'].isnull().sum()}")
        st.info(f"Conteo de nulos en 'costes_totreales' antes de retornar de process_data: {df['costes_totreales'].isnull().sum()}")
        st.info(f"Conteo de nulos en 'COSTO' antes de retornar de process_data: {df['COSTO'].isnull().sum()}")

        return df

    uploaded_file = st.file_uploader("� Sube el archivo Excel", type=["xlsx"], key="file_uploader_initial")

    if uploaded_file:
        try:
            raw_df = load_and_merge_data(uploaded_file)
            st.session_state.df_processed = process_data(raw_df)

            st.success(f"✅ Datos procesados. Filas: {len(st.session_state.df_processed)} | Columnas: {len(st.session_state.df_processed.columns)}")

            st.subheader("📊 Vista previa de los datos procesados")
            st.dataframe(st.session_state.df_processed.head(50), use_container_width=True)

            output_filename = "avisos_filtrados.xlsx"
            # Crear un objeto BytesIO para guardar el archivo Excel en memoria
            excel_buffer = io.BytesIO()
            st.session_state.df_processed.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)  # Rebobinar el búfer al principio

            st.download_button(
                "💾 Descargar archivo procesado",
                excel_buffer,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error al procesar el archivo: {e}")

# Asegurarse de que df esté disponible para las secciones subsiguientes
df = st.session_state.df_processed

if df is not None:
    # --- Funciones de Cálculo Comunes ---
    def calcular_indicadores_servicio(df_sub_filtered):
        if df_sub_filtered.empty:
            # Devolver Series vacías con dtype apropiado para concatenación posterior
            return (pd.Series(dtype=int), pd.Series(dtype=float),
                    pd.Series(dtype=float), pd.Series(dtype=float),
                    pd.Series(dtype=float), pd.Series(dtype=str))

        cnt = df_sub_filtered['tipo_de_servicio'].value_counts() # Usar nombre normalizado
        cost = df_sub_filtered.groupby('tipo_de_servicio')['COSTO'].sum()
        mttr = df_sub_filtered.groupby('tipo_de_servicio')['TIEMPO PARADA'].mean()

        # Calcular el tiempo total de operación por tipo de servicio
        # Asegurarse de que 'DIAS/ AÑO' y 'HORA/ DIA' sean numéricos y manejar NaNs
        df_sub_filtered_copy = df_sub_filtered.copy() # Trabajar en una copia para evitar SettingWithCopyWarning
        df_sub_filtered_copy['DIAS/ AÑO'] = pd.to_numeric(df_sub_filtered_copy['DIAS/ AÑO'], errors='coerce')
        df_sub_filtered_copy['HORA/ DIA'] = pd.to_numeric(df_sub_filtered_copy['HORA/ DIA'], errors='coerce')

        ttot = df_sub_filtered_copy.groupby('tipo_de_servicio').apply(
            lambda g: (g['DIAS/ AÑO'].mean() * g['HORA/ DIA'].mean()) if not g['DIAS/ AÑO'].isnull().all() and not g['HORA/ DIA'].isnull().all() else np.nan
        )
        ttot = pd.to_numeric(ttot, errors='coerce') # Convertir a numérico de forma robusta

        down = df_sub_filtered_copy.groupby('tipo_de_servicio')['TIEMPO PARADA'].sum()
        down = pd.to_numeric(down, errors='coerce') # Convertir a numérico de forma robusta
        
        fails = df_sub_filtered_copy.groupby('tipo_de_servicio')['AVISO_NUM'].count() # Usar nombre normalizado
        
        # Evitar división por cero
        mtbf = (ttot - down) / fails.replace(0, np.nan)
        
        # Evitar división por cero para disponibilidad
        disp = (mtbf / (mtbf + mttr)).replace([np.inf, -np.inf], np.nan) * 100
        rend = disp.apply(lambda v: 'Alto' if v >= 90 else ('Medio' if v >= 75 else 'Bajo') if not pd.isna(v) else np.nan)
        return cnt, cost, mttr, mtbf, disp, rend

    def calcular_indicadores_equipo(df_sub_filtered):
        if df_sub_filtered.empty:
            return (pd.Series(dtype=int), pd.Series(dtype=float),
                    pd.Series(dtype=float), pd.Series(dtype=float),
                    pd.Series(dtype=float), pd.Series(dtype=str), pd.Series(dtype=str))

        cnt = df_sub_filtered['EQUIPO_NUM'].value_counts()
        cost = df_sub_filtered.groupby('EQUIPO_NUM')['COSTO'].sum()
        mttr = df_sub_filtered.groupby('EQUIPO_NUM')['TIEMPO PARADA'].mean()

        # Agrupar por EQUIPMENT_NUM para obtener horas diarias y días/año promedio para el cálculo de MTBF
        equipo_group = df_sub_filtered.groupby('EQUIPO_NUM')
        
        df_sub_filtered_copy = df_sub_filtered.copy() # Trabajar en una copia
        df_sub_filtered_copy['DIAS/ AÑO'] = pd.to_numeric(df_sub_filtered_copy['DIAS/ AÑO'], errors='coerce')
        df_sub_filtered_copy['HORA/ DIA'] = pd.to_numeric(df_sub_filtered_copy['HORA/ DIA'], errors='coerce')
        
        ttot = equipo_group.apply(
            lambda g: (g['DIAS/ AÑO'].mean() * g['HORA/ DIA'].mean()) if not g['DIAS/ AÑO'].isnull().all() and not g['HORA/ DIA'].isnull().all() else np.nan
        )
        ttot = pd.to_numeric(ttot, errors='coerce') # Convertir a numérico de forma robusta

        down = equipo_group['TIEMPO PARADA'].sum()
        down = pd.to_numeric(down, errors='coerce') # Convertir a numérico de forma robusta
        
        fails = equipo_group['AVISO_NUM'].count()
        
        # Evitar división por cero
        mtbf = (ttot - down) / fails.replace(0, np.nan)
        
        # Evitar división por cero para disponibilidad
        disp = (mtbf / (mtbf + mttr)).replace([np.inf, -np.inf], np.nan) * 100

        rend = disp.apply(lambda v: 'Alto' if v >= 90 else ('Medio' if v >= 75 else 'Bajo') if not pd.isna(v) else np.nan)

        # Obtener la categoría de descripción más frecuente para cada equipo
        desc_cat = df_sub_filtered.groupby('EQUIPO_NUM')['description_category'].agg(lambda x: x.mode()[0] if not x.mode().empty else np.nan)

        return cnt, cost, mttr, mtbf, disp, rend, desc_cat

    # --- Sección de Análisis y Evaluación ---
    with st.container():
        st.header("2. Análisis de Costos y Avisos")

        # --- Opciones de Análisis ---
        analysis_options = {
            "Costos por ejecutante": ("denominacion_ejecutante", "costes_totreales", "costos"),
            "Avisos por ejecutante": ("denominacion_ejecutante", None, "avisos"),
            "Costos por objeto técnico": ("denominacion_de_objeto_tecnico", "costes_totreales", "costos"),
            "Avisos por objeto técnico": ("denominacion_de_objeto_tecnico", None, "avisos"),
            "Costos por texto código acción": ("texto_codigo_accion", "costes_totreales", "costos"),
            "Avisos por texto código acción": ("texto_codigo_accion", None, "avisos"),
            "Costos por texto de acción": ("texto_de_accion", "costes_totreales", "costos"),
            "Avisos por texto de acción": ("texto_de_accion", None, "avisos"),
            "Costos por tipo de servicio": ("tipo_de_servicio", "costes_totreales", "costos"),
            "Avisos por tipo de servicio": ("tipo_de_servicio", None, "avisos"),
            "Costos por categoría de descripción": ("description_category", "costes_totreales", "costos"),
            "Avisos por categoría de descripción": ("description_category", None, "avisos"),
        }

        if "analysis_page" not in st.session_state:
            st.session_state.analysis_page = 0

        # Obtener valores únicos para filtros
        # Asegurarse de que las columnas existan antes de intentar usarlas
        all_ejecutantes = sorted(df['denominacion_ejecutante'].dropna().unique()) if 'denominacion_ejecutante' in df.columns else []
        all_cps = sorted(df['codigo_postal'].dropna().unique()) if 'codigo_postal' in df.columns else []
        all_years = sorted(df['año'].dropna().astype(int).unique().tolist()) if 'año' in df.columns else []
        # Asegurar el orden correcto del mes
        month_order = ["January", "February", "March", "April", "May", "June",
                       "July", "August", "September", "October", "November", "December"]
        all_months_raw = df['mes'].dropna().unique().tolist() if 'mes' in df.columns else []
        all_months = sorted(all_months_raw, key=lambda x: month_order.index(x) if x in month_order else len(month_order))


        col1, col2 = st.columns(2)
        with col1:
            selected_ejecutantes = st.multiselect("Ejecutante", all_ejecutantes, default=all_ejecutantes, key="exec_filter")
        with col2:
            selected_cps = st.multiselect("Código postal", all_cps, default=all_cps, key="cp_filter")

        col3, col4 = st.columns(2)
        with col3:
            selected_year = st.selectbox("Año", ["Todos"] + all_years, key="year_filter")
        with col4:
            selected_month = st.selectbox("Mes", ["Todos"] + all_months, key="month_filter")

        selected_analysis_option = st.selectbox("Visualización", list(analysis_options.keys()), key="analysis_option_select")

        if selected_ejecutantes and selected_cps:
            # Filtrar df por las selecciones, asegurando que las columnas existan
            df_filtered_analysis = df.copy()
            if 'denominacion_ejecutante' in df_filtered_analysis.columns:
                df_filtered_analysis = df_filtered_analysis[df_filtered_analysis['denominacion_ejecutante'].isin(selected_ejecutantes)]
            if 'codigo_postal' in df_filtered_analysis.columns:
                df_filtered_analysis = df_filtered_analysis[df_filtered_analysis['codigo_postal'].isin(selected_cps)]
            
            if selected_year != "Todos" and 'año' in df_filtered_analysis.columns:
                df_filtered_analysis = df_filtered_analysis[df_filtered_analysis['año'] == selected_year]
            if selected_month != "Todos" and 'mes' in df_filtered_analysis.columns:
                df_filtered_analysis = df_filtered_analysis[df_filtered_analysis['mes'] == selected_month]

            col_agrup, col_cost, tipo_calc = analysis_options[selected_analysis_option]

            def display_analysis_table(df_to_show, col_agrup_disp, col_cost_disp, type_calc_disp, page_num, items_per_page=20):
                if col_agrup_disp not in df_to_show.columns:
                    st.warning(f"La columna de agrupación '{col_agrup_disp}' no se encontró para el análisis. Asegúrate de que el archivo Excel contenga los datos necesarios.")
                    return pd.DataFrame() # Devuelve un DataFrame vacío si la columna no existe

                if type_calc_disp == "costos":
                    # Verificar si la columna col_cost_disp existe antes de agrupar por ella
                    if col_cost_disp in df_to_show.columns:
                        grouped_df = df_to_show.groupby(col_agrup_disp)[col_cost_disp].sum().sort_values(ascending=False).reset_index()
                        grouped_df.columns = [col_agrup_disp, "Costo total"]
                    else:
                        st.warning(f"La columna de costos '{col_cost_disp}' no se encontró para el análisis.")
                        return pd.DataFrame()
                else: # tipo_calc == "avisos"
                    grouped_df = df_to_show[col_agrup_disp].value_counts().reset_index()
                    grouped_df.columns = [col_agrup_disp, "Cantidad de avisos"]

                total_items = grouped_df.shape[0]
                start = page_num * items_per_page
                end = start + items_per_page
                st.write(grouped_df.iloc[start:end])

                # Paginación
                num_pages = ((total_items - 1) // items_per_page) + 1
                col_nav1, col_nav2, col_nav3 = st.columns([1, 1, 5])
                with col_nav1:
                    if page_num > 0:
                        if st.button("← Página anterior", key="prev_analysis_page"):
                            st.session_state.analysis_page -= 1
                            # Streamlit se recargará automáticamente al cambiar el estado
                with col_nav2:
                    if end < total_items:
                        if st.button("Página siguiente →", key="next_analysis_page"):
                            st.session_state.analysis_page += 1
                            # Streamlit se recargará automáticamente al cambiar el estado
                with col_nav3:
                    st.markdown(f"Página {page_num + 1} de {num_pages}")
                return grouped_df

            display_analysis_table(df_filtered_analysis, col_agrup, col_cost, tipo_calc, st.session_state.analysis_page)
        else:
            st.info("Por favor, selecciona al menos un ejecutante y un código postal para el análisis.")


    # --- Sección de Evaluación de Proveedores ---
    with st.container():
        st.header("3. Evaluación de Proveedores")

        # Definir las preguntas y sus escalas
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
            ("Precio", "Facilita llegar a una negociación (precios)", "2,1,0,-1"),
            ("Precio", "Pone en consideración contratos y trabajos adjudicados en el último periodo de tiempo", "2,1,0,-1"),
            ("Postventa", "¿Tiene disposición y actitud de servicio frente a solicitudes?", "2,1,0,-1"),
            ("Postventa", "¿Conoce necesidades y ofrece alternativas adecuadas?","2,1,0,-1"),
            ("Postventa", "¿Realiza seguimiento a los resultados de los trabajos?", "2,1,0,-1"),
            ("Postventa", "¿Ofrece capacitaciones para el manejo de los equipos?", "2,1,0,-1"),
            ("Postventa", "¿Los métodos de capacitación ofrecidos son efectivos y adecuados?", "2,1,0,-1"), # Añadida la coma aquí
            ("Desempeño técnico", "Disponibilidad promedio (%)", "auto"),
            ("Desempeño técnico", "MTTR promedio (hrs)", "auto"),
            ("Desempeño técnico", "MTBF promedio (hrs)", "auto"),
            ("Desempeño técnico", "Rendimiento promedio equipos", "auto"),
        ]

        # Rangos detallados para mostrar
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

        # Función para mostrar los rangos de respuesta
        def mostrar_rangos_respuesta(preguntas_list, rangos_detallados_dict):
            st.subheader("📊 Rangos de Respuesta para cada Pregunta")

            with st.expander("Ver Escala General"):
                st.markdown("""
                **Escala General:**
                - `2`: Sobresaliente
                - `1`: Bueno
                - `0`: Indiferente
                - `-1`: Malo
                """)

            for cat, texto, escala in preguntas_list:
                st.markdown(f"#### [{cat}] {texto}")
                if escala == "auto":
                    rangos = rangos_detallados_dict.get(cat, {}).get(texto)
                    if rangos:
                        for val, desc in rangos.items():
                            st.markdown(f"- **{val}**: {desc}")
                    else:
                        st.markdown("_(Rangos definidos automáticamente por el sistema)_")
                else:
                    rangos = rangos_detallados_dict.get(cat, {}).get(texto)
                    if rangos:
                        for val, desc in rangos.items():
                            st.markdown(f"- **{val}**: {desc}")
                    else:
                        st.markdown(f"_Rangos: {escala}_")

        # Botón para mostrar los rangos de evaluación
        if st.button("🔍 Ver Rangos de Evaluación"):
            mostrar_rangos_respuesta(preguntas, rangos_detallados)

        # --- Selección de Proveedor ---
        providers = ["Todos"]
        if 'PROVEEDOR' in df.columns and not df['PROVEEDOR'].empty:
            providers.extend(sorted(df["PROVEEDOR"].dropna().unique()))
        selected_provider = st.selectbox("Seleccione un proveedor para evaluar", providers, key="eval_provider_select")

        # --- Cargar datos y métricas específicas del proveedor ---
        def cargar_datos_proveedor(data_df, prov):
            if prov == "Todos":
                sub_df = data_df.copy()
            else:
                if 'PROVEEDOR' in data_df.columns:
                    sub_df = data_df[data_df['PROVEEDOR'] == prov].copy()
                else:
                    st.warning("La columna 'PROVEEDOR' no se encontró en los datos para filtrar.")
                    return None, {}, [], pd.DataFrame(), pd.DataFrame()

            if sub_df.empty:
                st.warning(f"No hay datos disponibles para el proveedor '{prov}'.")
                return None, {}, [], pd.DataFrame(), pd.DataFrame()

            cnt_s, cost_s, mttr_s, mtbf_s, disp_s, rend_s = calcular_indicadores_servicio(sub_df)
            current_metrics = {'cnt': cnt_s, 'cost': cost_s, 'mttr': mttr_s, 'mtbf': mtbf_s, 'disp': disp_s, 'rend': rend_s}

            all_service_types = sorted(sub_df['tipo_de_servicio'].dropna().unique().tolist()) if 'tipo_de_servicio' in sub_df.columns else []

            resumen_servicio_df = pd.DataFrame({
                'Cantidad de Avisos': cnt_s,
                'Costo Total': cost_s,
                'Disponibilidad (%)': disp_s.round(2),
                'MTTR (hrs)': mttr_s.round(2),
                'MTBF (hrs)': mtbf_s.round(2),
                'Rendimiento': rend_s
            }).reset_index().rename(columns={'index': 'TIPO DE SERVICIO'})

            cnt_e, cost_e, mttr_e, mtbf_e, disp_e, rend_e, desc_cat_e = calcular_indicadores_equipo(sub_df)
            resumen_equipo_df = pd.DataFrame({
                'Avisos': cnt_e,
                'Costo total': cost_e,
                'MTTR': mttr_e.round(2),
                'MTBF': mtbf_e.round(2),
                'Disponibilidad (%)': disp_e.round(2),
                'Rendimiento': rend_e,
                'Categoría de Descripción': desc_cat_e
            }).reset_index().rename(columns={'index': 'Denominacion_Equipo'})

            return sub_df, current_metrics, all_service_types, resumen_servicio_df, resumen_equipo_df

        df_sub, current_provider_metrics, all_provider_service_types, summary_servicio_global, resumen_equipo_global = cargar_datos_proveedor(df, selected_provider)

        if df_sub is not None and not df_sub.empty:
            # Inicializar el estado de la sesión para las respuestas de evaluación
            if 'all_evaluation_widgets_map' not in st.session_state:
                st.session_state.all_evaluation_widgets_map = {}
            if 'current_eval_page' not in st.session_state:
                st.session_state.current_eval_page = 0

            st.subheader(f"Evaluación para: {selected_provider}")

            # Asegúrate de que 'cnt' existe y tiene un índice para tipos_servicio_eval
            tipos_servicio_eval = []
            if 'cnt' in current_provider_metrics and not current_provider_metrics['cnt'].empty:
                tipos_servicio_eval = list(current_provider_metrics['cnt'].index)
            else:
                st.warning("No se encontraron tipos de servicio para evaluar. Asegúrate de que los datos de servicio estén presentes.")
                tipos_servicio_eval = ["Servicio Desconocido"] # Proporcionar un valor predeterminado para evitar errores

            # Paginación para las preguntas de evaluación
            eval_services_per_page = 5
            eval_num_pages = len(tipos_servicio_eval) // eval_services_per_page + int(len(tipos_servicio_eval) % eval_services_per_page > 0)
            if eval_num_pages == 0: # Caso para cuando no hay servicios
                eval_num_pages = 1

            # Botones de navegación para las páginas de evaluación
            eval_col1, eval_col2, eval_col3 = st.columns([1, 1, 5])
            with eval_col1:
                if st.session_state.current_eval_page > 0:
                    if st.button("Página Anterior Evaluación", key="prev_eval_page"):
                        st.session_state.current_eval_page -= 1
                        st.rerun() # Para recargar la página y aplicar el cambio de estado
            with eval_col2:
                if st.session_state.current_eval_page < eval_num_pages - 1:
                    if st.button("Página Siguiente Evaluación", key="next_eval_page"):
                        st.session_state.current_eval_page += 1
                        st.rerun() # Para recargar la página y aplicar el cambio de estado
            with eval_col3:
                st.markdown(f"Página de Evaluación {st.session_state.current_eval_page + 1} de {eval_num_pages}")


            eval_start_idx = st.session_state.current_eval_page * eval_services_per_page
            eval_end_idx = eval_start_idx + eval_services_per_page
            services_on_current_page = tipos_servicio_eval[eval_start_idx:eval_end_idx]

            for tipo_servicio in services_on_current_page:
                st.markdown(f"### Servicio: {tipo_servicio}")
                for cat, texto, escala in preguntas:
                    question_key = (cat, texto, tipo_servicio)
                    if escala == "auto":
                        st.write(f"**[{cat}] {texto}**")
                        # Asegúrate de que los tipos de servicio existan en las métricas
                        if tipo_servicio in current_provider_metrics['disp'].index:
                            if "Disponibilidad" in texto:
                                val = current_provider_metrics['disp'][tipo_servicio]
                                st.metric("Disponibilidad (%)", f"{val:.2f}" if not pd.isna(val) else "N/A")
                            elif "MTTR" in texto:
                                val = current_provider_metrics['mttr'][tipo_servicio]
                                st.metric("MTTR (hrs)", f"{val:.2f}" if not pd.isna(val) else "N/A")
                            elif "MTBF" in texto:
                                val = current_provider_metrics['mtbf'][tipo_servicio]
                                st.metric("MTBF (hrs)", f"{val:.2f}" if not pd.isna(val) else "N/A")
                            elif "Rendimiento" in texto:
                                st.write(f"Rendimiento: **{current_provider_metrics['rend'][tipo_servicio]}**")
                        else:
                            st.info(f"No hay datos de métricas automáticas para el servicio '{tipo_servicio}'.")
                    else:
                        options = ["Sobresaliente (2)", "Bueno (1)", "Indiferente (0)", "Malo (-1)"]
                        default_index = 2  # Predeterminado a Indiferente (0)
                        if question_key in st.session_state.all_evaluation_widgets_map:
                             # Buscar el índice de la opción seleccionada previamente
                             prev_score_str = st.session_state.all_evaluation_widgets_map[question_key]
                             for i, opt in enumerate(options):
                                 if prev_score_str in opt:  # Verificar si la puntuación está en la cadena de opción
                                     default_index = i
                                     break

                        selected_option = st.radio(
                            f"**[{cat}] {texto}**",
                            options,
                            index=default_index,
                            horizontal=True,
                            key=f"eval_{cat}_{texto}_{tipo_servicio}"
                        )
                        # Extraer la puntuación numérica de la cadena de opción seleccionada
                        score = selected_option.split('(')[1].split(')')[0]
                        st.session_state.all_evaluation_widgets_map[question_key] = score

            st.markdown("---")
            st.subheader("Visualización de Métricas del Proveedor")

            # --- Gráficos de Rendimiento por Servicio ---
            if not current_provider_metrics['rend'].empty:
                st.subheader("Rendimiento por Servicio")
                fig, ax = plt.subplots(figsize=(8, 6))
                rend_counts = current_provider_metrics['rend'].value_counts().reindex(['Alto', 'Medio', 'Bajo'], fill_value=0)
                colors = ["#66bb6a", "#ffee58", "#ef5350"]  # Verde, Amarillo, Rojo
                ax.bar(rend_counts.index, rend_counts.values, color=colors)
                ax.set_title('Distribución de Rendimiento por Servicio')
                ax.set_xlabel('Rendimiento')
                ax.set_ylabel('Número de Servicios')
                st.pyplot(fig)
            else:
                st.info("No hay datos de rendimiento disponibles para graficar.")


            # --- Gráficos de Resumen de Métricas (MTTR, MTBF, Disponibilidad) ---
            st.subheader("Resumen de Métricas Técnicas por Servicio")
            def graficar_resumen(mttr_data, mtbf_data, disp_data):
                metrics = {
                    'MTTR (hrs)': mttr_data.dropna(),
                    'MTBF (hrs)': mtbf_data.dropna(),
                    'Disponibilidad (%)': disp_data.dropna()
                }
                # Filtrar datos vacíos para graficar
                metrics_to_plot = {k: v for k, v in metrics.items() if not v.empty}

                if not metrics_to_plot:
                    st.warning("No hay datos suficientes para graficar las métricas técnicas.")
                    return

                # Determinar el número de subgráficos necesarios
                num_plots = len(metrics_to_plot)
                if num_plots > 0:
                    fig, axs = plt.subplots(1, num_plots, figsize=(5 * num_plots, 4))
                    if num_plots == 1:  # Manejar el caso de un solo subgráfico
                        axs = [axs]

                    idx = 0
                    for title, data in metrics_to_plot.items():
                        sns.histplot(data, bins=10, kde=True, ax=axs[idx])
                        axs[idx].set_title(title)
                        axs[idx].set_xlabel(title.split(' ')[0])  # Solo el nombre de la métrica
                        axs[idx].set_ylabel('Frecuencia')
                        idx += 1
                    plt.tight_layout()
                    st.pyplot(fig)
                else:
                    st.warning("No hay datos válidos para generar los gráficos de resumen.")

            graficar_resumen(current_provider_metrics['mttr'], current_provider_metrics['mtbf'], current_provider_metrics['disp'])

            # --- Generar y Descargar Resumen de Evaluación ---
            if st.button("Generar Resumen de Evaluación", key="generate_summary_btn"):
                st.subheader("Generando resumen de evaluación...")

                if not st.session_state.all_evaluation_widgets_map:
                    st.warning("No hay evaluaciones para resumir. Completa las evaluaciones antes de generar el resumen.")
                else:
                    unique_service_types_evaluated = sorted({k[2] for k in st.session_state.all_evaluation_widgets_map.keys()})
                    all_categories_evaluated = sorted({p[0] for p in preguntas})

                    category_service_scores = {cat: {st_type: 0 for st_type in unique_service_types_evaluated} for cat in all_categories_evaluated}
                    service_type_question_counts = {st_type: {cat: 0 for cat in all_categories_evaluated} for st_type in unique_service_types_evaluated}
                    service_type_totals = {st_type: 0 for st_type in unique_service_types_evaluated}
                    service_type_overall_question_counts = {st_type: 0 for st_type in unique_service_types_evaluated}


                    for (cat, q_text, st_original), score_str in st.session_state.all_evaluation_widgets_map.items():
                        try:
                            score = int(score_str)
                            category_service_scores[cat][st_original] += score
                            service_type_totals[st_original] += score
                            service_type_question_counts[st_original][cat] += 1
                            service_type_overall_question_counts[st_original] += 1
                        except ValueError:
                            st.warning(f"Valor no numérico encontrado para la pregunta '{q_text}' del servicio '{st_original}'. Saltando.")


                    # Calcular puntuaciones promedio por categoría por tipo de servicio
                    average_category_service_scores = {cat: {st: np.nan for st in unique_service_types_evaluated} for cat in all_categories_evaluated}
                    for cat, service_scores in category_service_scores.items():
                        for st, total_score in service_scores.items():
                            num_questions = service_type_question_counts[st][cat]
                            if num_questions > 0:
                                average_category_service_scores[cat][st] = round(total_score / num_questions, 2) # Redondear a 2 decimales

                    # Calcular puntuación promedio general por tipo de servicio
                    average_service_type_scores = {st: np.nan for st in unique_service_types_evaluated}
                    for st, total_score in service_type_totals.items():
                        num_questions_overall = service_type_overall_question_counts[st]
                        if num_questions_overall > 0:
                            average_service_type_scores[st] = round(total_score / num_questions_overall, 2) # Redondear a 2 decimales


                    # Crear DataFrames para el resumen
                    summary_df_calificacion_raw = pd.DataFrame.from_dict(average_category_service_scores, orient='index')
                    summary_df_calificacion_raw.index.name = 'Categoría'

                    # Agregar fila de promedio general
                    average_scores_row = pd.Series(average_service_type_scores, name='Puntuación Promedio General por Servicio')
                    # Asegurarse de que el índice se alinee para concatenar
                    summary_df_calificacion_raw = pd.concat([summary_df_calificacion_raw, pd.DataFrame(average_scores_row).T])


                    service_type_display_names_cal = {
                        st: f"Servicio {all_provider_service_types.index(st)+1} ({st})"
                        if st in all_provider_service_types else f"Servicio ({st})"
                        for st in unique_service_types_evaluated
                    }
                    summary_df_calificacion = summary_df_calificacion_raw.rename(columns=service_type_display_names_cal)

                    output_filename_summary = f"resumen_evaluacion_{selected_provider.replace(' ', '_').replace('/', '-')}.xlsx"

                    buffer_summary = io.BytesIO()
                    with pd.ExcelWriter(buffer_summary, engine='xlsxwriter') as writer:
                        if summary_servicio_global is not None and not summary_servicio_global.empty:
                            summary_servicio_global.to_excel(writer, sheet_name='Resumen_Servicio', index=False)
                        if resumen_equipo_global is not None and not resumen_equipo_global.empty:
                            resumen_equipo_global.to_excel(writer, sheet_name='Resumen_Equipo', index=False)
                        if not summary_df_calificacion.empty:
                            summary_df_calificacion.to_excel(writer, sheet_name='Resumen_Calificacion')

                    st.success("Resumen generado exitosamente.")
                    buffer_summary.seek(0)
                    st.download_button(
                        "💾 Descargar resumen en Excel",
                        buffer_summary.getvalue(),
                        file_name=output_filename_summary,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.info("Selecciona un proveedor para ver sus métricas de desempeño y realizar la evaluación.")
�
