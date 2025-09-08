import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import pickle
import os
import numpy as np

# =========================
# Configuración de la página
# =========================
st.set_page_config(
    page_title="Dashboard de Análisis de Riesgo Psicosocial",
    page_icon="🧠",
    layout="wide"
)

# =========================
# Funciones de carga optimizadas
# =========================
@st.cache_data
def cargar_datos_procesados():
    """Carga todos los datos pre-procesados"""
    
    # Verificar que existen los archivos
    archivos_necesarios = [
        'datos_procesados/agrupaciones_base.pkl',
        'datos_procesados/datos_con_predicciones.pkl',
        'datos_procesados/agrupaciones_demograficas.pkl',
        'datos_procesados/variables_demograficas.pkl',
        'datos_procesados/datos_base_procesados.parquet',
        'datos_procesados/segmentacion.pkl'
    ]
    
    for archivo in archivos_necesarios:
        if not os.path.exists(archivo):
            st.error(f"❌ Archivo faltante: {archivo}")
            st.error("Por favor ejecuta primero 'generar_predicciones.py'")
            st.stop()
    
    # Cargar datos
    with open('datos_procesados/agrupaciones_base.pkl', 'rb') as f:
        agrupaciones_base = pickle.load(f)
    
    with open('datos_procesados/datos_con_predicciones.pkl', 'rb') as f:
        datos_con_predicciones = pickle.load(f)
    
    with open('datos_procesados/agrupaciones_demograficas.pkl', 'rb') as f:
        agrupaciones_demograficas = pickle.load(f)
    
    with open('datos_procesados/variables_demograficas.pkl', 'rb') as f:
        variables_demograficas = pickle.load(f)
    
    with open('datos_procesados/segmentacion.pkl', 'rb') as f:
        segmentacion = pickle.load(f)
    
    df = pd.read_parquet('datos_procesados/datos_base_procesados.parquet', engine='pyarrow')
    
    return df, agrupaciones_base, datos_con_predicciones, agrupaciones_demograficas, variables_demograficas, segmentacion

@st.cache_data
def obtener_datos_filtrados(_datos_con_predicciones, _agrupaciones_demograficas, _domain_to_factor, _dimension_to_factor,
                          factor_sel, dominio_sel, dimension_sel, var_demo, valor_demo, comparar_demo):
    """Obtiene los datos filtrados con predicciones pre-computadas"""
    
    dfs = []
    
    if var_demo == 'Sin filtro demográfico':
        # =============================================
        # DATOS GENERALES (sin variables demográficas)
        # =============================================
        
        if factor_sel == 'Todos':
            # Mostrar todos los niveles
            for nivel in ['Factor', 'Dominio', 'Dimension']:
                df_nivel = _datos_con_predicciones[nivel].copy()
                df_nivel['Grupo'] = df_nivel[nivel]
                dfs.append(df_nivel[['Año', 'Grupo', 'Nivel de riesgo codificado']])
        
        else:
            # Factor específico seleccionado
            # 1. Mostrar el factor seleccionado
            df_f = _datos_con_predicciones['Factor'][
                _datos_con_predicciones['Factor']['Factor'] == factor_sel
            ].copy()
            df_f['Grupo'] = df_f['Factor']
            dfs.append(df_f[['Año', 'Grupo', 'Nivel de riesgo codificado']])

            # 2. Mostrar dominios relacionados
            if dominio_sel == 'Todos':
                dominios = [d for d, f in _domain_to_factor.items() if f == factor_sel]
            else:
                dominios = [dominio_sel]

            df_d = _datos_con_predicciones['Dominio'][
                _datos_con_predicciones['Dominio']['Dominio'].isin(dominios)
            ].copy()
            df_d['Grupo'] = df_d['Dominio']
            dfs.append(df_d[['Año', 'Grupo', 'Nivel de riesgo codificado']])

            # 3. Mostrar dimensiones relacionadas
            if dimension_sel == 'Todos':
                dimensiones = [d for d, f in _dimension_to_factor.items() if f == factor_sel]
            else:
                dimensiones = [dimension_sel]

            df_dim = _datos_con_predicciones['Dimension'][
                _datos_con_predicciones['Dimension']['Dimension'].isin(dimensiones)
            ].copy()
            df_dim['Grupo'] = df_dim['Dimension']
            dfs.append(df_dim[['Año', 'Grupo', 'Nivel de riesgo codificado']])
    
    else:
        # =============================================
        # DATOS CON VARIABLES DEMOGRÁFICAS - CORREGIDO
        # =============================================
        
        # Determinar qué niveles mostrar
        niveles_a_mostrar = []
        if factor_sel != 'Todos':
            niveles_a_mostrar.append(('Factor', factor_sel))
        if dominio_sel != 'Todos':
            niveles_a_mostrar.append(('Dominio', dominio_sel))
        if dimension_sel != 'Todos':
            niveles_a_mostrar.append(('Dimension', dimension_sel))
        
        # Si no hay selecciones específicas, mostrar todos los niveles
        if not niveles_a_mostrar:
            niveles_a_mostrar = [('Factor', 'Todos'), ('Dominio', 'Todos'), ('Dimension', 'Todos')]
        
        for nivel_tipo, nivel_valor in niveles_a_mostrar:
            if nivel_tipo in _agrupaciones_demograficas and var_demo in _agrupaciones_demograficas[nivel_tipo]:
                df_demo = _agrupaciones_demograficas[nivel_tipo][var_demo].copy()
                
                # ========================================
                # APLICAR FILTROS DE NIVEL
                # ========================================
                if nivel_valor != 'Todos':
                    df_demo = df_demo[df_demo[nivel_tipo] == nivel_valor]
                
                # ========================================
                # APLICAR FILTROS DEMOGRÁFICOS - CORREGIDO
                # ========================================
                if valor_demo != 'Todos':
                    # Aplicar filtro incluso cuando comparar_demo es True
                    df_demo = df_demo[df_demo[var_demo] == valor_demo]
                
                # ========================================
                # CREAR GRUPOS PARA EL GRÁFICO
                # ========================================
                if comparar_demo:
                    # Si comparar_demo es True pero ya filtramos por valor_demo específico,
                    # solo mostraremos ese valor específico en combinación con los niveles
                    df_demo['Grupo'] = df_demo[nivel_tipo].astype(str) + ' - ' + df_demo[var_demo].astype(str)
                else:
                    df_demo['Grupo'] = df_demo[nivel_tipo]
                
                dfs.append(df_demo[['Año', 'Grupo', 'Nivel de riesgo codificado']])
    
    # Combinar todos los DataFrames
    if dfs:
        df_combinado = pd.concat(dfs, ignore_index=True)
        df_combinado = df_combinado.drop_duplicates()
        return df_combinado
    else:
        return pd.DataFrame(columns=['Año', 'Grupo', 'Nivel de riesgo codificado'])

# =========================
# Funciones auxiliares
# =========================
def obtener_color_riesgo(valor):
    if valor < 0.5:
        return "lightgray"
    elif valor < 1.5:
        return "#149e11"
    elif valor < 2.5:
        return "#92d050"
    elif valor < 3.5:
        return "#ffff00"
    elif valor < 4.5:
        return "red"
    else:
        return "darkred"

# =========================
# Aplicación principal
# =========================
def main():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image("https://441041d6dc.imgdist.com/pub/bfra/989mykjl/3jw/n2n/7ki/Logo%20Adecco.png", width=300)

    st.title("🧠 Dashboard de Análisis de Riesgo Psicosocial")
    st.subheader("📈 Tendencias y Predicciones para 2026")
    
    # =========================
    # CARGA SUPER RÁPIDA
    # =========================
    with st.spinner("Cargando datos pre-procesados..."):
        df, agrupaciones_base, datos_con_predicciones, agrupaciones_demograficas, variables_demograficas, segmentacion = cargar_datos_procesados()
    
    st.success("¡Datos cargados instantáneamente! ⚡")
    
    # Crear mapeos para dominio y dimensión que asocien al factor
    domain_to_factor = {}
    dimension_to_factor = {}

    for key, (f, d, dim) in segmentacion.items():
        if d is not None:
            domain_to_factor[d] = f
        if dim is not None:
            dimension_to_factor[dim] = f
    
    # =========================
    # SIDEBAR PARA FILTROS
    # =========================
    
    st.sidebar.header("🎯 Filtros de Niveles")
    
    # Filtro de Factor
    factores = ['Todos'] + sorted([f for f in datos_con_predicciones['Factor']['Factor'].dropna().unique()])
    factor_seleccionado = st.sidebar.selectbox("🔎 Selecciona el Factor:", factores)
    
    # Filtro de Dominio (dinámico basado en factor)
    if factor_seleccionado == 'Todos':
        dominios_disponibles = sorted(datos_con_predicciones['Dominio']['Dominio'].dropna().unique())
    else:
        dominios_disponibles = sorted([d for d, f in domain_to_factor.items() if f == factor_seleccionado])
    
    dominios = ['Todos'] + dominios_disponibles
    dominio_seleccionado = st.sidebar.selectbox("🗂 Selecciona el Dominio:", dominios)
    
    # Filtro de Dimensión (dinámico basado en dominio/factor)
    if dominio_seleccionado == 'Todos':
        if factor_seleccionado == 'Todos':
            dimensiones_disponibles = sorted(datos_con_predicciones['Dimension']['Dimension'].dropna().unique())
        else:
            dimensiones_disponibles = sorted([dim for dim, f in dimension_to_factor.items() if f == factor_seleccionado])
    else:
        dimensiones_disponibles = sorted([
            dim for dim, f in dimension_to_factor.items()
            if f == factor_seleccionado and segmentacion.get(dim, (None, None, None))[1] == dominio_seleccionado
        ])
    
    dimensiones = ['Todos'] + dimensiones_disponibles
    dimension_seleccionada = st.sidebar.selectbox("📊 Selecciona la Dimensión:", dimensiones)
    
    # =========================
    # FILTROS DEMOGRÁFICOS CON NOMBRES AMIGABLES
    # =========================
    st.sidebar.header("👥 Filtros Demográficos")
    
    # Crear mapeo para mostrar nombres más amigables
    nombres_amigables = {
        'Sexo': 'Género',  # ← Cambio visual clave
        'Generación': 'Generación', 
        'Rango de Edad': 'Rango de Edad', 
        'Tipo de servicio': 'Tipo de servicio', 
        'Seleccione tipo de cargo que mas se parece': 'Tipo de cargo', 
        'Estrato según servicios Públicos': 'Estrato socioeconómico',
        'Empresa': 'Empresa',
        'Factor a Evaluar': 'Factor a Evaluar'
    }
    
    # Crear opciones con nombres amigables
    variables_demo_display = ['Sin filtro demográfico'] + [nombres_amigables.get(var, var) for var in variables_demograficas]
    variable_demografica_display = st.sidebar.selectbox("🧬 Variable Demográfica:", variables_demo_display)
    
    # Convertir de vuelta al nombre real de la columna
    if variable_demografica_display != 'Sin filtro demográfico':
        # Buscar la clave original basada en el valor mostrado
        variable_demografica = next(k for k, v in nombres_amigables.items() if v == variable_demografica_display)
    else:
        variable_demografica = 'Sin filtro demográfico'
    
    if variable_demografica != 'Sin filtro demográfico':
        valores_demo = ['Todos'] + sorted(df[variable_demografica].dropna().unique())
        valor_demografico = st.sidebar.selectbox("💡 Valor de la Variable:", valores_demo)
        
        if valor_demografico == 'Todos':
            comparar_demograficos = st.sidebar.checkbox(f"📊 Comparar todos los valores de {variable_demografica_display}")
        else:
            comparar_demograficos = False
            st.sidebar.info(f"Filtrando solo por: {valor_demografico}")
    else:
        valor_demografico = 'Todos'
        comparar_demograficos = False
    
    # =========================
    # GENERAR GRÁFICO SÚPER RÁPIDO
    # =========================
    
    try:
        # Obtener datos filtrados (ya con predicciones incluidas)
        df_combinado = obtener_datos_filtrados(
            datos_con_predicciones, agrupaciones_demograficas, domain_to_factor, dimension_to_factor,
            factor_seleccionado, dominio_seleccionado, dimension_seleccionada, 
            variable_demografica, valor_demografico, comparar_demograficos
        )
        
        # Crear título del gráfico con nombre amigable
        titulo = 'Tendencia del Nivel de Riesgo'
        if variable_demografica != 'Sin filtro demográfico':
            titulo += f' - {variable_demografica_display}'  # Usar nombre amigable
            if valor_demografico != 'Todos' and not comparar_demograficos:
                titulo += f': {valor_demografico}'
        
        # Crear gráfico
        if df_combinado.empty:
            st.warning("No hay datos disponibles para los filtros seleccionados")
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
                          fillcolor="red", opacity=0.2, layer="below", line_width=0)
            fig.add_shape(type="rect", xref="paper", yref="y", x0=0, x1=1, y0=4.5, y1=5.5, 
                          fillcolor="darkred", opacity=0.2, layer="below", line_width=0)

            # Añadir anotaciones para las franjas
            fig.add_annotation(x=0.02, y=0, text="No Referido", showarrow=False, xref="paper", yref="y")
            fig.add_annotation(x=0.02, y=1, text="Sin Riesgo", showarrow=False, xref="paper", yref="y")
            fig.add_annotation(x=0.02, y=2, text="Riesgo Bajo", showarrow=False, xref="paper", yref="y")
            fig.add_annotation(x=0.02, y=3, text="Riesgo Medio", showarrow=False, xref="paper", yref="y")
            fig.add_annotation(x=0.02, y=4, text="Riesgo Alto", showarrow=False, xref="paper", yref="y")
            fig.add_annotation(x=0.02, y=5, text="Riesgo Muy Alto", showarrow=False, xref="paper", yref="y")

            fig.update_layout(
                template='plotly_white', 
                title_x=0.5, 
                yaxis_title="Nivel de riesgo codificado",
                xaxis_title="Año",
                hovermode='x unified',
                height=600
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
                    st.subheader("🔮 Predicciones para 2026")
                    
                    # Crear tabla de predicciones
                    tabla_datos = []
                    for grupo in df_2026['Grupo'].unique():
                        valor = df_2026[df_2026['Grupo'] == grupo]['Nivel de riesgo codificado'].iloc[0]
                        
                        # Determinar el nivel de riesgo textual
                        if valor < 0.5:
                            nivel_texto = "No Referido"
                        elif valor < 1.5:
                            nivel_texto = "Sin Riesgo"
                        elif valor < 2.5:
                            nivel_texto = "Riesgo Bajo"
                        elif valor < 3.5:
                            nivel_texto = "Riesgo Medio"
                        elif valor < 4.5:
                            nivel_texto = "Riesgo Alto"
                        else:
                            nivel_texto = "Riesgo Muy Alto"
                        
                        tabla_datos.append({
                            'Grupo': grupo,
                            'Valor Predicho': f"{valor:.2f}",
                            'Nivel de Riesgo': nivel_texto
                        })
                    
                    df_tabla = pd.DataFrame(tabla_datos)
                    st.dataframe(df_tabla, use_container_width=True)
                    
                    # Análisis de tendencias
                    st.subheader("📈 Análisis de Tendencias")
                    
                    tendencias = []
                    for grupo in df_combinado['Grupo'].unique():
                        df_grupo = df_combinado[df_combinado['Grupo'] == grupo]
                        if len(df_grupo) >= 2:
                            años = df_grupo['Año'].values
                            valores = df_grupo['Nivel de riesgo codificado'].values
                            
                            if len(años) > 1:
                                # Ordenar por año para calcular tendencia correctamente
                                orden = np.argsort(años)
                                años_ord = años[orden]
                                valores_ord = valores[orden]
                                
                                pendiente = (valores_ord[-1] - valores_ord[0]) / (años_ord[-1] - años_ord[0])
                                
                                if pendiente > 0.1:
                                    tendencia = "↑ Aumentando"
                                    color = "🔴"
                                elif pendiente < -0.1:
                                    tendencia = "↓ Disminuyendo"
                                    color = "🟢"
                                else:
                                    tendencia = "→ Estable"
                                    color = "🟡"
                                
                                tendencias.append(f"{color} **{grupo}**: {tendencia} ({pendiente:.3f} puntos/año)")
                    
                    if tendencias:
                        for tendencia in tendencias:
                            st.write(f"{tendencia}")
                    
                    # Información de rendimiento
                    st.sidebar.info(f"🚀 **Optimización activa**\n\nPredicciones pre-calculadas para máxima velocidad.\n\n📊 Datos mostrados: {len(df_combinado)} registros")
    
    except Exception as e:
        st.error(f"Error al generar el gráfico: {str(e)}")
        st.exception(e)

if __name__ == "__main__":
    main()
