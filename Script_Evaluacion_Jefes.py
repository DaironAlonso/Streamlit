import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pickle
import os
import numpy as np


# =========================
# SISTEMA DE AUTENTICACIÓN
# =========================

def cargar_usuarios():
    """Carga el archivo de usuarios"""
    try:
        usuarios_df = pd.read_excel('Usuarios.xlsx')
        # Limpiar espacios en blanco
        usuarios_df['Empresa'] = usuarios_df['Empresa'].str.strip()
        usuarios_df['contraseña'] = usuarios_df['contraseña'].str.strip()
        return usuarios_df
    except Exception as e:
        st.error(f"Error al cargar usuarios: {e}")
        return pd.DataFrame()

def verificar_credenciales(empresa, contraseña, usuarios_df):
    """Verifica si las credenciales son válidas"""
    usuario = usuarios_df[
        (usuarios_df['Empresa'].str.lower() == empresa.lower()) & 
        (usuarios_df['contraseña'] == contraseña)
    ]
    return not usuario.empty

def pagina_login():
    """Muestra la página de login"""
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image("https://441041d6dc.imgdist.com/pub/bfra/989mykjl/3jw/n2n/7ki/Logo%20Adecco.png", width=300)
    
    st.title("🔐 Inicio de Sesión")
    st.markdown("### Dashboard de Análisis de Riesgo Psicosocial")
    
    usuarios_df = cargar_usuarios()
    
    if usuarios_df.empty:
        st.error("No se pudo cargar el archivo de usuarios")
        return
    
    with st.form("login_form"):
        empresa = st.text_input("Usuario (Empresa)", placeholder="Ingrese el nombre de la empresa")
        contraseña = st.text_input("Contraseña", type="password", placeholder="Ingrese su contraseña")
        submit = st.form_submit_button("Iniciar Sesión", use_container_width=True)
        
        if submit:
            if empresa and contraseña:
                if verificar_credenciales(empresa, contraseña, usuarios_df):
                    st.session_state['autenticado'] = True
                    st.session_state['usuario'] = empresa
                    st.session_state['es_admin'] = (empresa.upper() == 'ADMIN')
                    st.success(f"¡Bienvenido {empresa}!")
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos")
            else:
                st.warning("Por favor ingrese usuario y contraseña")
    
    

def cerrar_sesion():
    """Cierra la sesión del usuario"""
    st.session_state['autenticado'] = False
    st.session_state['usuario'] = None
    st.session_state['es_admin'] = False
    st.rerun()


# =========================
# Configuración de la página
# =========================
st.set_page_config(
    page_title="Dashboard de Análisis de Riesgo Psicosocial",
    page_icon="🧠",
    layout="wide"
)


# =========================
# CONTROL DE AUTENTICACIÓN
# =========================
if 'autenticado' not in st.session_state:
    st.session_state['autenticado'] = False
    st.session_state['usuario'] = None
    st.session_state['es_admin'] = False

if not st.session_state['autenticado']:
    pagina_login()
    st.stop()


# =========================
# Funciones de carga optimizadas
# =========================
@st.cache_data
def cargar_datos_procesados():
    """Carga todos los datos pre-procesados jerárquicos"""
    
    # Verificar que existen los archivos
    archivos_necesarios = [
        'datos_procesados/predicciones_jerarquicas.pkl',
        'datos_procesados/datos_con_predicciones_jerarquicos.pkl',
        'datos_procesados/variables_demograficas.pkl',
        'datos_procesados/datos_base_procesados.parquet',
        'datos_procesados/empresas.pkl',
        'datos_procesados/segmentacion.pkl'
    ]
    
    for archivo in archivos_necesarios:
        if not os.path.exists(archivo):
            st.error(f"❌ Archivo faltante: {archivo}")
            st.error("Por favor ejecuta primero el script de generación de predicciones actualizado")
            st.stop()
    
    # Cargar datos jerárquicos
    with open('datos_procesados/predicciones_jerarquicas.pkl', 'rb') as f:
        predicciones = pickle.load(f)
    
    with open('datos_procesados/datos_con_predicciones_jerarquicos.pkl', 'rb') as f:
        datos_con_predicciones = pickle.load(f)
    
    with open('datos_procesados/variables_demograficas.pkl', 'rb') as f:
        variables_demograficas = pickle.load(f)
    
    with open('datos_procesados/empresas.pkl', 'rb') as f:
        empresas = pickle.load(f)
    
    with open('datos_procesados/segmentacion.pkl', 'rb') as f:
        segmentacion = pickle.load(f)
    
    df = pd.read_parquet('datos_procesados/datos_base_procesados.parquet', engine='pyarrow')
    
    return df, predicciones, datos_con_predicciones, variables_demograficas, empresas, segmentacion


def crear_mapeo_jerarquico(segmentacion):
    """
    Crea mapeos jerárquicos entre Factor → Dominio → Dimensión
    IMPORTANTE: Factor Extralaboral NO tiene dominios, solo dimensiones directas
    """
    
    # Mapear Factor → Dominios (solo para Factor Intralaboral)
    factor_a_dominios = {}
    for key, (factor, dominio, dimension) in segmentacion.items():
        if factor not in factor_a_dominios:
            factor_a_dominios[factor] = set()
        if dominio is not None:
            factor_a_dominios[factor].add(dominio)
    
    # Convertir sets a listas ordenadas
    for factor in factor_a_dominios:
        factor_a_dominios[factor] = sorted(list(factor_a_dominios[factor]))
    
    # Mapear Dominio → Dimensiones (solo para Factor Intralaboral)
    dominio_a_dimensiones = {}
    for key, (factor, dominio, dimension) in segmentacion.items():
        if dominio is not None and dimension is not None:
            if dominio not in dominio_a_dimensiones:
                dominio_a_dimensiones[dominio] = []
            if dimension not in dominio_a_dimensiones[dominio]:
                dominio_a_dimensiones[dominio].append(dimension)
    
    # Ordenar dimensiones
    for dominio in dominio_a_dimensiones:
        dominio_a_dimensiones[dominio] = sorted(dominio_a_dimensiones[dominio])
    
    # Mapear Factor → Dimensiones DIRECTAS (para Factor Extralaboral)
    factor_a_dimensiones_directas = {}
    for key, (factor, dominio, dimension) in segmentacion.items():
        # Solo agregar dimensiones que NO tienen dominio (Factor Extralaboral)
        if dominio is None and dimension is not None:
            if factor not in factor_a_dimensiones_directas:
                factor_a_dimensiones_directas[factor] = []
            if dimension not in factor_a_dimensiones_directas[factor]:
                factor_a_dimensiones_directas[factor].append(dimension)
    
    # Ordenar dimensiones directas
    for factor in factor_a_dimensiones_directas:
        factor_a_dimensiones_directas[factor] = sorted(factor_a_dimensiones_directas[factor])
    
    # Mapear Dimensión → Factor (para búsqueda inversa)
    dimension_a_factor = {}
    dominio_a_factor = {}
    for key, (factor, dominio, dimension) in segmentacion.items():
        if dominio is not None:
            dominio_a_factor[dominio] = factor
        if dimension is not None:
            dimension_a_factor[dimension] = factor
    
    return factor_a_dominios, dominio_a_dimensiones, factor_a_dimensiones_directas, dimension_a_factor, dominio_a_factor


def limpiar_sin_dominio_dimension_en_grupo(df):
    """
    Filtra un DataFrame para eliminar solo las LÍNEAS/GRUPOS que literalmente
    se llaman 'Sin Dominio' o 'Sin Dimensión' en la columna Grupo.
    NO filtra los factores completos.
    """
    if df.empty:
        return df
    
    df_limpio = df.copy()
    
    # Solo filtrar en la columna "Grupo" si existe
    if 'Grupo' in df_limpio.columns:
        # Filtrar líneas que literalmente se llaman "Sin Dominio"
        df_limpio = df_limpio[~df_limpio['Grupo'].astype(str).str.match(r'^Sin Dominio$', case=False, na=False)]
        # Filtrar líneas que literalmente se llaman "Sin Dimensión"
        df_limpio = df_limpio[~df_limpio['Grupo'].astype(str).str.match(r'^Sin Dimensi[oó]n$', case=False, na=False)]
    
    return df_limpio

def limpiar_valores_negativos_2026(df):
    """
    Convierte valores negativos a 0 para predicciones del año 2026
    """
    if df.empty:
        return df
    
    df_limpio = df.copy()
    
    # Aplicar el filtro solo para el año 2026
    mask_2026 = df_limpio['Año'] == 2026
    mask_negativos = df_limpio['Nivel de riesgo codificado'] < 0
    
    # Convertir valores negativos a 0 solo en 2026
    df_limpio.loc[mask_2026 & mask_negativos, 'Nivel de riesgo codificado'] = 0
    
    return df_limpio


# =========================
# NUEVA FUNCIÓN: Obtener descripciones de etiquetas
# =========================
def obtener_descripcion_etiqueta(nombre_elemento, tipo_elemento):
    """
    Retorna la descripción interpretativa del nivel de riesgo según el elemento.
    tipo_elemento puede ser: 'Factor', 'Dominio', 'Dimension'
    """
    
    descripciones = {
        # Factores
        'Factor Intralaboral': {
            'Sin Riesgo': 'Las condiciones del ambiente laboral favorecen el bienestar y el desempeño.',
            'Riesgo Bajo': 'Las condiciones laborales son mayormente favorables con áreas menores de mejora.',
            'Riesgo Medio': 'Existen condiciones laborales que requieren atención y mejora.',
            'Riesgo Alto': 'Las condiciones laborales representan una fuente importante de estrés.',
            'Riesgo Muy Alto': 'Las condiciones laborales son críticas y requieren intervención inmediata.'
        },
        'Factor Extralaboral': {
            'Sin Riesgo': 'Las condiciones externas al trabajo no afectan el bienestar del trabajador.',
            'Riesgo Bajo': 'Existen condiciones externas leves que podrían afectar el desempeño.',
            'Riesgo Medio': 'Las condiciones externas requieren atención para evitar impactos negativos.',
            'Riesgo Alto': 'Las condiciones externas están afectando significativamente al trabajador.',
            'Riesgo Muy Alto': 'Las condiciones externas representan una amenaza crítica para el bienestar.'
        },
        'Estrés': {
            'Sin Riesgo': 'El trabajador no presenta síntomas significativos de estrés.',
            'Riesgo Bajo': 'Se observan síntomas leves de estrés que son manejables.',
            'Riesgo Medio': 'Los síntomas de estrés requieren atención y manejo activo.',
            'Riesgo Alto': 'El estrés está afectando significativamente la salud y desempeño.',
            'Riesgo Muy Alto': 'Nivel crítico de estrés que requiere intervención urgente.'
        }
    }
    
    # Si el elemento específico tiene descripciones definidas
    if nombre_elemento in descripciones:
        return descripciones[nombre_elemento]
    
    # Descripciones genéricas según tipo
    if tipo_elemento == 'Dominio':
        return {
            'Sin Riesgo': f'El {nombre_elemento} no representa riesgo para el trabajador.',
            'Riesgo Bajo': f'El {nombre_elemento} presenta condiciones levemente desfavorables.',
            'Riesgo Medio': f'El {nombre_elemento} requiere intervenciones preventivas.',
            'Riesgo Alto': f'El {nombre_elemento} requiere intervención prioritaria.',
            'Riesgo Muy Alto': f'El {nombre_elemento} está en estado crítico.'
        }
    elif tipo_elemento == 'Dimension':
        return {
            'Sin Riesgo': f'La dimensión {nombre_elemento} no presenta factores de riesgo.',
            'Riesgo Bajo': f'La dimensión {nombre_elemento} muestra indicadores leves de riesgo.',
            'Riesgo Medio': f'La dimensión {nombre_elemento} necesita medidas correctivas.',
            'Riesgo Alto': f'La dimensión {nombre_elemento} requiere atención inmediata.',
            'Riesgo Muy Alto': f'La dimensión {nombre_elemento} está en nivel crítico.'
        }
    
    # Descripciones por defecto
    return {
        'Sin Riesgo': 'No se identifican condiciones de riesgo.',
        'Riesgo Bajo': 'Se identifican condiciones leves de riesgo.',
        'Riesgo Medio': 'Se requieren acciones preventivas.',
        'Riesgo Alto': 'Se requiere intervención prioritaria.',
        'Riesgo Muy Alto': 'Situación crítica que requiere acción inmediata.'
    }


def obtener_nivel_riesgo_texto(valor):
    """Convierte el valor numérico a texto del nivel de riesgo"""
    if valor < 0.5:
        return "No Referido"
    elif valor < 1.5:
        return "Sin Riesgo"
    elif valor < 2.5:
        return "Riesgo Bajo"
    elif valor < 3.5:
        return "Riesgo Medio"
    elif valor < 4.5:
        return "Riesgo Alto"
    else:
        return "Riesgo Muy Alto"


@st.cache_data
def obtener_datos_filtrados_jerarquicos(_df, _datos_con_predicciones, empresas_sel, 
                                        factores_sel, dominios_sel, dimensiones_sel, 
                                        var_demo, valores_demo, comparar_demo):
    """
    Obtiene los datos filtrados siguiendo la jerarquía:
    Agregado → Empresas (múltiples) → Factores (múltiples) → Dominios/Dimensiones (múltiples) → Demografía
    """
    
    dfs = []
    
    # Convertir listas vacías a "Todos"
    if not empresas_sel:
        empresas_sel = ['Todas']
    if not factores_sel:
        factores_sel = ['Todos']
    if not dominios_sel:
        dominios_sel = ['Todos']
    if not dimensiones_sel:
        dimensiones_sel = ['Todos']
    if not valores_demo:
        valores_demo = ['Todos']
    
    if var_demo == 'Sin filtro demográfico':
        # =============================================
        # DATOS SIN DEMOGRAFÍA
        # =============================================
        
        if 'Todas' in empresas_sel:
            # USAR DATOS AGREGADOS (todas las empresas como una sola)
            
            # 1. Agregar factores seleccionados
            if 'Todos' in factores_sel:
                # Mostrar todos los factores
                if 'Factor' in _datos_con_predicciones['agregadas_nivel']:
                    df_precalc = _datos_con_predicciones['agregadas_nivel']['Factor'].copy()
                    df_precalc['Grupo'] = df_precalc['Factor']
                    df_precalc['Tipo'] = 'Factor'
                    dfs.append(df_precalc[['Año', 'Grupo', 'Tipo', 'Nivel de riesgo codificado']])
            else:
                # Factores específicos seleccionados
                if 'Factor' in _datos_con_predicciones['agregadas_nivel']:
                    df_factor = _datos_con_predicciones['agregadas_nivel']['Factor'].copy()
                    df_factor = df_factor[df_factor['Factor'].isin(factores_sel)]
                    df_factor['Grupo'] = df_factor['Factor']
                    df_factor['Tipo'] = 'Factor'
                    dfs.append(df_factor[['Año', 'Grupo', 'Tipo', 'Nivel de riesgo codificado']])
                
                # 2. Agregar dominios relacionados (SOLO si hay Factor Intralaboral seleccionado)
                if 'Factor Intralaboral' in factores_sel:
                    if 'Dominio' in _datos_con_predicciones['agregadas_nivel']:
                        df_dominio = _datos_con_predicciones['agregadas_nivel']['Dominio'].copy()
                        
                        if 'Todos' in dominios_sel:
                            # Mostrar todos los dominios del Factor Intralaboral
                            dominios_intralaboral = _df[_df['Factor'] == 'Factor Intralaboral']['Dominio'].dropna().unique()
                            df_dominio = df_dominio[df_dominio['Dominio'].isin(dominios_intralaboral)]
                        else:
                            # Dominios específicos
                            df_dominio = df_dominio[df_dominio['Dominio'].isin(dominios_sel)]
                        
                        df_dominio['Grupo'] = df_dominio['Dominio']
                        df_dominio['Tipo'] = 'Dominio'
                        dfs.append(df_dominio[['Año', 'Grupo', 'Tipo', 'Nivel de riesgo codificado']])
                
                # 3. Agregar dimensiones relacionadas
                if 'Dimension' in _datos_con_predicciones['agregadas_nivel']:
                    df_dimension = _datos_con_predicciones['agregadas_nivel']['Dimension'].copy()
                    
                    if 'Todos' in dimensiones_sel:
                        # Mostrar todas las dimensiones de los factores seleccionados
                        dimensiones_de_factores = _df[_df['Factor'].isin(factores_sel)]['Dimension'].dropna().unique()
                        df_dimension = df_dimension[df_dimension['Dimension'].isin(dimensiones_de_factores)]
                    else:
                        # Dimensiones específicas
                        df_dimension = df_dimension[df_dimension['Dimension'].isin(dimensiones_sel)]
                    
                    df_dimension['Grupo'] = df_dimension['Dimension']
                    df_dimension['Tipo'] = 'Dimension'
                    dfs.append(df_dimension[['Año', 'Grupo', 'Tipo', 'Nivel de riesgo codificado']])
        
        else:
            # Empresas específicas - procesar cada una
            for empresa_sel in empresas_sel:
                # Factores
                if 'Todos' in factores_sel:
                    # Mostrar todos los factores de la empresa
                    clave = f"{empresa_sel}_Factor"
                    if clave in _datos_con_predicciones['por_empresa_nivel']:
                        df_precalc = _datos_con_predicciones['por_empresa_nivel'][clave].copy()
                        df_precalc['Grupo'] = f"{empresa_sel} - " + df_precalc['Factor'].astype(str)
                        df_precalc['Tipo'] = 'Factor'
                        dfs.append(df_precalc[['Año', 'Grupo', 'Tipo', 'Nivel de riesgo codificado']])
                
                else:
                    # Factores específicos de empresa
                    clave_factor = f"{empresa_sel}_Factor"
                    if clave_factor in _datos_con_predicciones['por_empresa_nivel']:
                        df_factor = _datos_con_predicciones['por_empresa_nivel'][clave_factor].copy()
                        df_factor = df_factor[df_factor['Factor'].isin(factores_sel)]
                        df_factor['Grupo'] = f"{empresa_sel} - " + df_factor['Factor'].astype(str)
                        df_factor['Tipo'] = 'Factor'
                        dfs.append(df_factor[['Año', 'Grupo', 'Tipo', 'Nivel de riesgo codificado']])
                    
                    # Dominios de la empresa (SOLO si Factor Intralaboral está seleccionado)
                    if 'Factor Intralaboral' in factores_sel:
                        clave_dominio = f"{empresa_sel}_Dominio"
                        if clave_dominio in _datos_con_predicciones['por_empresa_nivel']:
                            df_dominio = _datos_con_predicciones['por_empresa_nivel'][clave_dominio].copy()
                            df_empresa = _df[_df['Empresa'] == empresa_sel]
                            
                            if 'Todos' in dominios_sel:
                                # Todos los dominios del Factor Intralaboral
                                dominios_intralaboral = df_empresa[df_empresa['Factor'] == 'Factor Intralaboral']['Dominio'].dropna().unique()
                                df_dominio = df_dominio[df_dominio['Dominio'].isin(dominios_intralaboral)]
                            else:
                                df_dominio = df_dominio[df_dominio['Dominio'].isin(dominios_sel)]
                            
                            df_dominio['Grupo'] = f"{empresa_sel} - " + df_dominio['Dominio'].astype(str)
                            df_dominio['Tipo'] = 'Dominio'
                            dfs.append(df_dominio[['Año', 'Grupo', 'Tipo', 'Nivel de riesgo codificado']])
                    
                    # Dimensiones de la empresa
                    clave_dimension = f"{empresa_sel}_Dimension"
                    if clave_dimension in _datos_con_predicciones['por_empresa_nivel']:
                        df_dimension = _datos_con_predicciones['por_empresa_nivel'][clave_dimension].copy()
                        df_empresa = _df[_df['Empresa'] == empresa_sel]
                        
                        if 'Todos' in dimensiones_sel:
                            # Todas las dimensiones de los factores seleccionados
                            dimensiones_de_factores = df_empresa[df_empresa['Factor'].isin(factores_sel)]['Dimension'].dropna().unique()
                            df_dimension = df_dimension[df_dimension['Dimension'].isin(dimensiones_de_factores)]
                        else:
                            df_dimension = df_dimension[df_dimension['Dimension'].isin(dimensiones_sel)]
                        
                        df_dimension['Grupo'] = f"{empresa_sel} - " + df_dimension['Dimension'].astype(str)
                        df_dimension['Tipo'] = 'Dimension'
                        dfs.append(df_dimension[['Año', 'Grupo', 'Tipo', 'Nivel de riesgo codificado']])
    
    else:
        # =============================================
        # DATOS CON DEMOGRAFÍA
        # =============================================
        
        if 'Todas' in empresas_sel:
            # Usar datos demográficos agregados
            for nivel in ['Factor', 'Dominio', 'Dimension']:
                clave_demo_agregada = f"{nivel}_{var_demo}"
                
                if clave_demo_agregada in _datos_con_predicciones['agregadas_nivel_demografico']:
                    df_precalc = _datos_con_predicciones['agregadas_nivel_demografico'][clave_demo_agregada].copy()
                    
                    # Filtrar según selección
                    if 'Todos' not in factores_sel:
                        if nivel == 'Factor':
                            df_precalc = df_precalc[df_precalc['Factor'].isin(factores_sel)]
                        elif nivel == 'Dominio':
                            # Solo procesar dominios si Factor Intralaboral está seleccionado
                            if 'Factor Intralaboral' not in factores_sel:
                                continue
                            df_filtro = _df[_df['Factor'] == 'Factor Intralaboral']
                            if 'Todos' not in dominios_sel:
                                dominios_validos = dominios_sel
                            else:
                                dominios_validos = df_filtro['Dominio'].dropna().unique()
                            df_precalc = df_precalc[df_precalc['Dominio'].isin(dominios_validos)]
                        elif nivel == 'Dimension':
                            df_filtro = _df[_df['Factor'].isin(factores_sel)]
                            if 'Todos' not in dimensiones_sel:
                                dimensiones_validas = dimensiones_sel
                            else:
                                dimensiones_validas = df_filtro['Dimension'].dropna().unique()
                            df_precalc = df_precalc[df_precalc['Dimension'].isin(dimensiones_validas)]
                    
                    if 'Todos' not in valores_demo:
                        df_precalc = df_precalc[df_precalc[var_demo].isin(valores_demo)]
                    
                    if comparar_demo:
                        df_precalc['Grupo'] = df_precalc[nivel].astype(str) + ' - ' + df_precalc[var_demo].astype(str)
                    else:
                        if 'Todos' not in valores_demo and len(valores_demo) == 1:
                            df_precalc['Grupo'] = df_precalc[nivel].astype(str) + ' (' + valores_demo[0] + ')'
                        else:
                            df_precalc['Grupo'] = df_precalc[nivel].astype(str) + ' - ' + df_precalc[var_demo].astype(str)
                    
                    df_precalc['Tipo'] = nivel
                    dfs.append(df_precalc[['Año', 'Grupo', 'Tipo', 'Nivel de riesgo codificado']])
        else:
            # Empresas específicas con demografía
            for empresa_sel in empresas_sel:
                for nivel in ['Factor', 'Dominio', 'Dimension']:
                    clave_demo = f"{empresa_sel}_{nivel}_{var_demo}"
                    
                    if clave_demo in _datos_con_predicciones['por_empresa_nivel_demografico']:
                        df_precalc = _datos_con_predicciones['por_empresa_nivel_demografico'][clave_demo].copy()
                        df_empresa = _df[_df['Empresa'] == empresa_sel]
                        
                        # Filtrar según selección
                        if 'Todos' not in factores_sel:
                            if nivel == 'Factor':
                                df_precalc = df_precalc[df_precalc['Factor'].isin(factores_sel)]
                            elif nivel == 'Dominio':
                                # Solo procesar dominios si Factor Intralaboral está seleccionado
                                if 'Factor Intralaboral' not in factores_sel:
                                    continue
                                if 'Todos' not in dominios_sel:
                                    dominios_validos = dominios_sel
                                else:
                                    dominios_validos = df_empresa[df_empresa['Factor'] == 'Factor Intralaboral']['Dominio'].dropna().unique()
                                df_precalc = df_precalc[df_precalc['Dominio'].isin(dominios_validos)]
                            elif nivel == 'Dimension':
                                if 'Todos' not in dimensiones_sel:
                                    dimensiones_validas = dimensiones_sel
                                else:
                                    dimensiones_validas = df_empresa[df_empresa['Factor'].isin(factores_sel)]['Dimension'].dropna().unique()
                                df_precalc = df_precalc[df_precalc['Dimension'].isin(dimensiones_validas)]
                        
                        if 'Todos' not in valores_demo:
                            df_precalc = df_precalc[df_precalc[var_demo].isin(valores_demo)]
                        
                        if comparar_demo:
                            df_precalc['Grupo'] = f"{empresa_sel} - " + df_precalc[nivel].astype(str) + ' - ' + df_precalc[var_demo].astype(str)
                        else:
                            if 'Todos' not in valores_demo and len(valores_demo) == 1:
                                df_precalc['Grupo'] = f"{empresa_sel} - " + df_precalc[nivel].astype(str) + ' (' + valores_demo[0] + ')'
                            else:
                                df_precalc['Grupo'] = f"{empresa_sel} - " + df_precalc[nivel].astype(str) + ' - ' + df_precalc[var_demo].astype(str)
                        
                        df_precalc['Tipo'] = nivel
                        dfs.append(df_precalc[['Año', 'Grupo', 'Tipo', 'Nivel de riesgo codificado']])
    
    # Combinar todos los DataFrames
    if dfs:
        df_combinado = pd.concat(dfs, ignore_index=True)
        df_combinado = df_combinado.drop_duplicates()
        
        # FILTRO FINAL: Eliminar solo las líneas que se llaman exactamente "Sin Dominio" o "Sin Dimensión"
        df_combinado = limpiar_sin_dominio_dimension_en_grupo(df_combinado)

        df_combinado = limpiar_valores_negativos_2026(df_combinado)
        
        return df_combinado
    else:
        return pd.DataFrame(columns=['Año', 'Grupo', 'Tipo', 'Nivel de riesgo codificado'])


# =========================
# Aplicación principal
# =========================
def main():
    # Mostrar información del usuario y botón de cerrar sesión
    col_header1, col_header2, col_header3 = st.columns([1, 2, 1])
    
    with col_header1:
        st.write(f"👤 **Usuario:** {st.session_state['usuario']}")
        if st.session_state['es_admin']:
            st.success("🔑 Administrador")
    
    with col_header2:
        st.image("https://441041d6dc.imgdist.com/pub/bfra/989mykjl/3jw/n2n/7ki/Logo%20Adecco.png", width=300)
    
    with col_header3:
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            cerrar_sesion()
    
    st.title("🧠 Dashboard de Análisis de Riesgo Psicosocial")
    st.subheader("📈 Tendencias y Predicciones para 2026 - Análisis Jerárquico")
    
    # =========================
    # CARGA SUPER RÁPIDA
    # =========================
    with st.spinner("Cargando datos pre-procesados jerárquicos..."):
        df, predicciones, datos_con_predicciones, variables_demograficas, empresas, segmentacion = cargar_datos_procesados()
    
    # **FILTRAR DATOS SEGÚN EL USUARIO**
    if not st.session_state['es_admin']:
        # Usuario normal: solo ver datos de su empresa
        df = df[df['Empresa'] == st.session_state['usuario']]
        # Filtrar también las empresas disponibles
        empresas = [st.session_state['usuario']]
    
    # Crear mapeos jerárquicos
    factor_a_dominios, dominio_a_dimensiones, factor_a_dimensiones_directas, dimension_a_factor, dominio_a_factor = crear_mapeo_jerarquico(segmentacion)
    
    st.success("¡Datos cargados instantáneamente! ⚡")
    
    # =========================
    # SIDEBAR PARA FILTROS JERÁRQUICOS
    # =========================
    
    st.sidebar.header("🏢 NIVEL 1: Empresa")
    
    # FILTRO 1: Empresas (MULTISELECT)
    empresas_disponibles = sorted(empresas)
    
    # Si es admin, puede seleccionar múltiples empresas. Si no, solo la suya
    if st.session_state['es_admin']:
        empresas_seleccionadas = st.sidebar.multiselect(
            "🏢 Selecciona una o más Empresas:",
            options=empresas_disponibles,
            default=None,
            help="Deja vacío para ver datos agregados de todas las empresas"
        )
        
        if not empresas_seleccionadas:
            st.sidebar.info("✅ Se mostrarán datos agregados (todas las empresas)")
            empresas_seleccionadas = ['Todas']
    else:
        # Usuario normal: solo su empresa
        st.sidebar.info(f"📌 Filtrando datos de: **{st.session_state['usuario']}**")
        empresas_seleccionadas = [st.session_state['usuario']]
    
    st.sidebar.markdown("---")
    st.sidebar.header("📊 NIVEL 2: Factores Psicosociales")
    
    # FILTRO 2: Factores (MULTISELECT)
    factores_disponibles = sorted(df['Factor'].dropna().unique())
    factores_seleccionados = st.sidebar.multiselect(
        "📊 Selecciona uno o más Factores:",
        options=factores_disponibles,
        default=None,
        help="Deja vacío para seleccionar todos"
    )
    
    # Si no se selecciona nada, considerar "Todos"
    if not factores_seleccionados:
        st.sidebar.info("✅ Se mostrarán todos los factores")
        factores_seleccionados = ['Todos']
    
    # FILTRO 3: Dominios (MULTISELECT, solo para Factor Intralaboral)
    dominios_seleccionados = ['Todos']
    mostrar_dominios = False
    
    if factores_seleccionados and factores_seleccionados != ['Todos']:
        # Solo mostrar dominios si Factor Intralaboral está seleccionado
        if 'Factor Intralaboral' in factores_seleccionados:
            mostrar_dominios = True
            st.sidebar.markdown("---")
            st.sidebar.header("🗂️ NIVEL 3: Dominios (Factor Intralaboral)")
            
            # Obtener dominios del Factor Intralaboral
            dominios_disponibles = factor_a_dominios.get('Factor Intralaboral', [])
            
            if dominios_disponibles:
                dominios_seleccionados = st.sidebar.multiselect(
                    "🗂️ Selecciona uno o más Dominios:",
                    options=dominios_disponibles,
                    default=None,
                    help="Deja vacío para seleccionar todos los dominios del Factor Intralaboral"
                )
                
                if not dominios_seleccionados:
                    st.sidebar.info("✅ Se mostrarán todos los dominios del Factor Intralaboral")
                    dominios_seleccionados = ['Todos']
    
    # FILTRO 4: Dimensiones (MULTISELECT)
    dimensiones_seleccionadas = ['Todos']
    
    if factores_seleccionados and factores_seleccionados != ['Todos']:
        st.sidebar.markdown("---")
        st.sidebar.header("📈 NIVEL 4: Dimensiones")
        
        dimensiones_disponibles = set()
        
        # Para Factor Intralaboral: obtener dimensiones de dominios seleccionados
        if 'Factor Intralaboral' in factores_seleccionados:
            if dominios_seleccionados and dominios_seleccionados != ['Todos']:
                for dominio in dominios_seleccionados:
                    if dominio in dominio_a_dimensiones:
                        dimensiones_disponibles.update(dominio_a_dimensiones[dominio])
            else:
                # Todas las dimensiones del Factor Intralaboral
                for dominio in factor_a_dominios.get('Factor Intralaboral', []):
                    if dominio in dominio_a_dimensiones:
                        dimensiones_disponibles.update(dominio_a_dimensiones[dominio])
        
        # Para Factor Extralaboral y otros: dimensiones directas
        for factor in factores_seleccionados:
            if factor in factor_a_dimensiones_directas:
                dimensiones_disponibles.update(factor_a_dimensiones_directas[factor])
        
        dimensiones_disponibles = sorted(list(dimensiones_disponibles))
        
        if dimensiones_disponibles:
            dimensiones_seleccionadas = st.sidebar.multiselect(
                "📈 Selecciona una o más Dimensiones:",
                options=dimensiones_disponibles,
                default=None,
                help="Deja vacío para seleccionar todas las dimensiones relacionadas"
            )
            
            if not dimensiones_seleccionadas:
                st.sidebar.info("✅ Se mostrarán todas las dimensiones relacionadas")
                dimensiones_seleccionadas = ['Todos']
        else:
            st.sidebar.info("Los factores/dominios seleccionados no tienen dimensiones asociadas")
    
    st.sidebar.markdown("---")
    st.sidebar.header("👥 NIVEL 5: Filtros Demográficos")
    
    # =========================
    # FILTROS DEMOGRÁFICOS CON NOMBRES AMIGABLES
    # =========================
    
    nombres_amigables = {
        'Sexo': 'Género',
        'Generación': 'Generación', 
        'Rango de Edad': 'Rango de Edad', 
        'Tipo de servicio': 'Tipo de servicio', 
        'Seleccione tipo de cargo que mas se parece': 'Tipo de cargo', 
        'Estrato según servicios Públicos': 'Estrato socioeconómico',
        'Factor a Evaluar': 'Factor a Evaluar'
    }
    
    variables_demo_display = ['Sin filtro demográfico'] + [
        nombres_amigables.get(var, var) for var in variables_demograficas 
        if var != 'Empresa'
    ]
    variable_demografica_display = st.sidebar.selectbox("🧬 Variable Demográfica:", variables_demo_display)
    
    # Convertir de vuelta al nombre real de la columna
    if variable_demografica_display != 'Sin filtro demográfico':
        variable_demografica = next(k for k, v in nombres_amigables.items() if v == variable_demografica_display)
    else:
        variable_demografica = 'Sin filtro demográfico'
    
    if variable_demografica != 'Sin filtro demográfico':
        # Filtrar valores demográficos basados en empresas seleccionadas
        if 'Todas' in empresas_seleccionadas:
            valores_demo_disponibles = sorted(df[variable_demografica].dropna().unique())
        else:
            valores_demo_disponibles = sorted(
                df[df['Empresa'].isin(empresas_seleccionadas)][variable_demografica].dropna().unique()
            )
        
        # MULTISELECT para valores demográficos
        valores_demograficos = st.sidebar.multiselect(
            "💡 Selecciona uno o más valores:",
            options=valores_demo_disponibles,
            default=None,
            help="Deja vacío para seleccionar todos"
        )
        
        if not valores_demograficos:
            st.sidebar.info("✅ Se mostrarán todos los valores")
            valores_demograficos = ['Todos']
            comparar_demograficos = False
        else:
            if len(valores_demograficos) > 1:
                comparar_demograficos = st.sidebar.checkbox(
                    f"📊 Comparar los {len(valores_demograficos)} valores seleccionados",
                    value=True
                )
            else:
                comparar_demograficos = False
                st.sidebar.info(f"Filtrando solo por: {valores_demograficos[0]}")
    else:
        valores_demograficos = ['Todos']
        comparar_demograficos = False
    
    # =========================
    # INFORMACIÓN DE CONTEXTO
    # =========================
    st.sidebar.markdown("---")
    st.sidebar.header("ℹ️ Filtros Activos")
    
    filtros_activos = []
    
    if empresas_seleccionadas and empresas_seleccionadas != ['Todas']:
        filtros_activos.append(f"🏢 Empresas: {len(empresas_seleccionadas)}")
        for empresa in empresas_seleccionadas[:3]:
            filtros_activos.append(f"   • {empresa}")
        if len(empresas_seleccionadas) > 3:
            filtros_activos.append(f"   ... y {len(empresas_seleccionadas) - 3} más")
    
    if factores_seleccionados and factores_seleccionados != ['Todos']:
        filtros_activos.append(f"📊 Factores: {len(factores_seleccionados)}")
        for factor in factores_seleccionados[:3]:
            filtros_activos.append(f"   • {factor}")
        if len(factores_seleccionados) > 3:
            filtros_activos.append(f"   ... y {len(factores_seleccionados) - 3} más")
    
    if dominios_seleccionados and dominios_seleccionados != ['Todos']:
        filtros_activos.append(f"🗂️ Dominios: {len(dominios_seleccionados)}")
        for dominio in dominios_seleccionados[:3]:
            filtros_activos.append(f"   • {dominio}")
        if len(dominios_seleccionados) > 3:
            filtros_activos.append(f"   ... y {len(dominios_seleccionados) - 3} más")
    
    if dimensiones_seleccionadas and dimensiones_seleccionadas != ['Todos']:
        filtros_activos.append(f"📈 Dimensiones: {len(dimensiones_seleccionadas)}")
        for dimension in dimensiones_seleccionadas[:2]:
            filtros_activos.append(f"   • {dimension}")
        if len(dimensiones_seleccionadas) > 2:
            filtros_activos.append(f"   ... y {len(dimensiones_seleccionadas) - 2} más")
    
    if variable_demografica != 'Sin filtro demográfico':
        filtros_activos.append(f"👥 {variable_demografica_display}")
        if valores_demograficos != ['Todos']:
            filtros_activos.append(f"   Valores: {len(valores_demograficos)}")
            for valor in valores_demograficos[:2]:
                filtros_activos.append(f"   • {valor}")
            if len(valores_demograficos) > 2:
                filtros_activos.append(f"   ... y {len(valores_demograficos) - 2} más")
    
    if filtros_activos:
        for filtro in filtros_activos:
            st.sidebar.text(filtro)
    else:
        st.sidebar.text("Sin filtros aplicados")
    
    # =========================
    # GENERAR GRÁFICO CON PREDICCIONES
    # =========================
    
    try:
        # Obtener datos filtrados jerárquicamente (CON PREDICCIONES 2026)
        df_combinado = obtener_datos_filtrados_jerarquicos(
            df, datos_con_predicciones, empresas_seleccionadas,
            factores_seleccionados, dominios_seleccionados, dimensiones_seleccionadas,
            variable_demografica, valores_demograficos, comparar_demograficos
        )
        
        # DEBUG: Mostrar años disponibles
        if not df_combinado.empty:
            años_disponibles = sorted(df_combinado['Año'].unique())
            st.sidebar.text(f"Años: {años_disponibles}")
        
        # Crear título del gráfico
        titulo = 'Tendencia del Nivel de Riesgo'
        
        if empresas_seleccionadas and empresas_seleccionadas != ['Todas']:
            if len(empresas_seleccionadas) == 1:
                titulo += f' - {empresas_seleccionadas[0]}'
            else:
                titulo += f' - {len(empresas_seleccionadas)} Empresas'
        
        if factores_seleccionados and factores_seleccionados != ['Todos']:
            if len(factores_seleccionados) == 1:
                titulo += f' - {factores_seleccionados[0]}'
            else:
                titulo += f' - {len(factores_seleccionados)} Factores'
        
        if dominios_seleccionados and dominios_seleccionados != ['Todos']:
            if len(dominios_seleccionados) == 1:
                titulo += f' > {dominios_seleccionados[0]}'
            else:
                titulo += f' > {len(dominios_seleccionados)} Dominios'
        
        if dimensiones_seleccionadas and dimensiones_seleccionadas != ['Todos']:
            if len(dimensiones_seleccionadas) == 1:
                titulo += f' > {dimensiones_seleccionadas[0]}'
            else:
                titulo += f' > {len(dimensiones_seleccionadas)} Dimensiones'
        
        if variable_demografica != 'Sin filtro demográfico':
            titulo += f' - {variable_demografica_display}'
            if valores_demograficos != ['Todos']:
                if len(valores_demograficos) == 1:
                    titulo += f': {valores_demograficos[0]}'
                else:
                    titulo += f' ({len(valores_demograficos)} valores)'
        
        # Crear gráfico
        if df_combinado.empty:
            st.warning("⚠️ No hay datos disponibles para los filtros seleccionados")
            st.info("💡 Intenta cambiar los filtros o dejar algunos vacíos para ver datos agregados")
        else:
            fig = px.line(
                df_combinado,
                x='Año',
                y='Nivel de riesgo codificado',
                color='Grupo',
                markers=True,
                title=titulo
            )

            # Añadir franjas de riesgo en el fondo
            fig.add_shape(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=-0.5, y1=0.5, 
                          fillcolor="#E0E0E0", opacity=0.2, layer="below", line_width=0)
            fig.add_shape(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=0.5, y1=1.5, 
                          fillcolor="#149e11", opacity=0.2, layer="below", line_width=0)
            fig.add_shape(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=1.5, y1=2.5, 
                          fillcolor="#92d050", opacity=0.2, layer="below", line_width=0)
            fig.add_shape(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=2.5, y1=3.5, 
                          fillcolor="#ffff00", opacity=0.2, layer="below", line_width=0)
            fig.add_shape(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=3.5, y1=4.5, 
                          fillcolor="#ff0000", opacity=0.2, layer="below", line_width=0)
            fig.add_shape(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=4.5, y1=5.5, 
                          fillcolor="#8b0000", opacity=0.2, layer="below", line_width=0)

            # Añadir anotaciones para las franjas
            fig.add_annotation(x=0.02, y=0, text="<b>No Referido</b>", showarrow=False, xref="paper", yref="y", 
                             font=dict(size=14, color="black"))
            fig.add_annotation(x=0.02, y=1, text="<b>Sin Riesgo</b>", showarrow=False, xref="paper", yref="y",
                             font=dict(size=14, color="black"))
            fig.add_annotation(x=0.02, y=2, text="<b>Riesgo Bajo</b>", showarrow=False, xref="paper", yref="y",
                             font=dict(size=14, color="black"))
            fig.add_annotation(x=0.02, y=3, text="<b>Riesgo Medio</b>", showarrow=False, xref="paper", yref="y",
                             font=dict(size=14, color="black"))
            fig.add_annotation(x=0.02, y=4, text="<b>Riesgo Alto</b>", showarrow=False, xref="paper", yref="y",
                             font=dict(size=14, color="black"))
            fig.add_annotation(x=0.02, y=5, text="<b>Riesgo Muy Alto</b>", showarrow=False, xref="paper", yref="y",
                             font=dict(size=14, color="black"))

            # Resaltar el año 2026 con una línea vertical
            if 2026 in df_combinado['Año'].values:
                fig.add_vline(x=2026, line_dash="dash", line_color="red", line_width=2,
                            annotation_text="<b>Predicción 2026</b>", 
                            annotation_position="top",
                            annotation_font_size=16,
                            annotation_font_color="red")

            fig.update_layout(
                template='plotly_white', 
                title_x=0.5, 
                yaxis_title="Nivel de riesgo codificado",
                xaxis_title="Año",
                hovermode='x unified',
                height=600,
                title_font_size=20,
                title_font_family="Arial Black"
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # =========================
            # ESTADÍSTICAS ADICIONALES
            # =========================
            
            if not df_combinado.empty:
                st.header("📊 Estadísticas Adicionales")
                
                # Predicciones para 2026
                df_2026 = df_combinado[df_combinado['Año'] == 2026]
                
                if not df_2026.empty:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.subheader("🔮 Predicciones para 2026")
                        
                        # Crear tabla de predicciones
                        tabla_datos = []
                        for _, row in df_2026.iterrows():
                            grupo = row['Grupo']
                            valor = row['Nivel de riesgo codificado']
                            tipo = row.get('Tipo', 'Desconocido')
                            
                            # Determinar el nivel de riesgo textual y emoji con color
                            nivel_texto = obtener_nivel_riesgo_texto(valor)
                            
                            if valor < 0.5:
                                emoji = "⚪"
                            elif valor < 1.5:
                                emoji = "🟢"
                            elif valor < 2.5:
                                emoji = "🟢"
                            elif valor < 3.5:
                                emoji = "🟡"
                            elif valor < 4.5:
                                emoji = "🔴"
                            else:
                                emoji = "🔴"
                            
                            tabla_datos.append({
                                'Estado': emoji,
                                'Grupo': f"{grupo}",
                                'Valor Predicho': f"{valor:.2f}",
                                'Nivel de Riesgo': f"{nivel_texto}",
                                'Tipo': tipo
                            })
                        
                        df_tabla = pd.DataFrame(tabla_datos)
                        st.dataframe(df_tabla, use_container_width=True, hide_index=True)
                    
                    with col2:
                        st.subheader("📈 Análisis de Tendencias")
                        
                        tendencias = []
                        for grupo in df_combinado['Grupo'].unique():
                            df_grupo = df_combinado[df_combinado['Grupo'] == grupo]
                            if len(df_grupo) >= 2:
                                años = df_grupo['Año'].values
                                valores = df_grupo['Nivel de riesgo codificado'].values
                                
                                if len(años) > 1:
                                    # Ordenar por año
                                    orden = np.argsort(años)
                                    años_ord = años[orden]
                                    valores_ord = valores[orden]
                                    
                                    pendiente = (valores_ord[-1] - valores_ord[0]) / (años_ord[-1] - años_ord[0])
                                    
                                    if pendiente > 0.1:
                                        tendencia = "↑ Aumentando"
                                        color = "🔴"
                                        categoria = "⚠️ Empeorando"
                                    elif pendiente < -0.1:
                                        tendencia = "↓ Disminuyendo"
                                        color = "🟢"
                                        categoria = "✅ Mejorando"
                                    else:
                                        tendencia = "→ Estable"
                                        color = "🟡"
                                        categoria = "🔄 Estable"
                                    
                                    tendencias.append({
                                        'emoji': color,
                                        'grupo': grupo,
                                        'tendencia': tendencia,
                                        'pendiente': pendiente,
                                        'categoria': categoria
                                    })
                        
                        if tendencias:
                            for t in tendencias:
                                st.write(f"{t['emoji']} **{t['grupo']}**")
                                st.write(f"   {t['tendencia']} ({t['pendiente']:.3f} puntos/año)")
                                st.write(f"   {t['categoria']}")
                                st.write("")
                else:
                    st.info("ℹ️ No hay predicciones para 2026 con los filtros seleccionados")
                
                # =========================
                # SECCIÓN DE DESCRIPCIÓN DE ETIQUETAS Y CATEGORÍAS
                # =========================
                st.markdown("---")
                st.header("📖 Descripción de Categorías de Riesgo")
                
                with st.expander("🔍 Ver descripción detallada de categorías y niveles", expanded=False):
                    st.markdown("""
                    ### Niveles de Riesgo Psicosocial
                    
                    Las predicciones del sistema clasifican el riesgo en las siguientes categorías:
                    """)
                    
                    col_a, col_b = st.columns(2)
                    
                    with col_a:
                        st.markdown("""
                        #### ⚪ **No Referido** (0 - 0.5)
                        - No aplica evaluación de riesgo
                        - Sin datos suficientes para clasificar
                        
                        #### 🟢 **Sin Riesgo** (0.5 - 1.5)
                        - Condiciones óptimas
                        - Ambiente favorable para el trabajador
                        - No se requieren intervenciones
                        
                        #### 🟢 **Riesgo Bajo** (1.5 - 2.5)
                        - Condiciones favorables en general
                        - Aspectos menores a mejorar
                        - Monitoreo preventivo recomendado
                        """)
                    
                    with col_b:
                        st.markdown("""
                        #### 🟡 **Riesgo Medio** (2.5 - 3.5)
                        - Condiciones que requieren atención
                        - Intervenciones preventivas necesarias
                        - Seguimiento periódico obligatorio
                        
                        #### 🔴 **Riesgo Alto** (3.5 - 4.5)
                        - Condiciones desfavorables importantes
                        - Intervención prioritaria requerida
                        - Plan de acción inmediato
                        
                        #### 🔴 **Riesgo Muy Alto** (4.5 - 5.5)
                        - Situación crítica
                        - Intervención urgente obligatoria
                        - Alto impacto en salud y bienestar
                        """)
                    
                    st.markdown("---")
                    
                    # Mostrar descripciones específicas de los elementos seleccionados
                    if not df_2026.empty:
                        st.subheader("📝 Interpretación de Predicciones 2026")
                        
                        for _, row in df_2026.iterrows():
                            grupo = row['Grupo']
                            valor = row['Nivel de riesgo codificado']
                            tipo = row.get('Tipo', 'Desconocido')
                            nivel_texto = obtener_nivel_riesgo_texto(valor)
                            
                            # Extraer nombre limpio del elemento
                            nombre_limpio = grupo.split(' - ')[-1] if ' - ' in grupo else grupo
                            
                            # Obtener descripción específica
                            descripciones_elemento = obtener_descripcion_etiqueta(nombre_limpio, tipo)
                            descripcion = descripciones_elemento.get(nivel_texto, "Sin descripción disponible")
                            
                            # Mostrar con formato destacado
                            if valor >= 3.5:
                                st.error(f"**{grupo}**: {nivel_texto}")
                            elif valor >= 2.5:
                                st.warning(f"**{grupo}**: {nivel_texto}")
                            else:
                                st.success(f"**{grupo}**: {nivel_texto}")
                            
                            st.write(f"_{descripcion}_")
                            st.write("")
                
                # Información de rendimiento
                total_registros = len(df_combinado)
                total_grupos = df_combinado['Grupo'].nunique()
                años_con_2026 = 2026 in df_combinado['Año'].values
                
                st.sidebar.markdown("---")
                st.sidebar.info(
                    f"🚀 **Optimización Jerárquica Activa**\n\n"
                    f"📊 Registros: {total_registros}\n\n"
                    f"🎯 Grupos: {total_grupos}\n\n"
                    f"{'✅' if años_con_2026 else '❌'} Predicciones 2026\n\n"
                    f"⚡ Datos pre-calculados"
                )
    
    except Exception as e:
        st.error(f"❌ Error al generar el gráfico: {str(e)}")
        with st.expander("🔍 Ver detalles del error"):
            st.exception(e)


if __name__ == "__main__":
    main()
