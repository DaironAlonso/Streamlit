import streamlit as st
import time
from PIL import Image
import requests
from io import BytesIO

# --- CONFIGURACIÓN INICIAL ---
st.set_page_config(
    page_title="NPS Zonamerica 2025 – BC", 
    page_icon="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABwAAAAcCAMAAABF0y+mAAAAKlBMVEVHcEwwrGgwrGgwrGgwrGgwrGgwrGgwrGgwrGgwrGgwrGgwrGgwrGgwrGjV3fMIAAAADXRSTlMATcoq3TlrrwhiivAW61nCVAAAAJ9JREFUKJHNkUsSxCAIREURxMT7X3ckxM9osp2a3hh8dtEQ5/5XwJ7hhfmiyvqZQGKaWSim4E575tPqPBhLk//yEkU9G8T7PgrdUSAfpTdQiRUZMhaWpLGRyR5z70LnOkLsTK46hTBg7hDrLVxpfVidPYkqtrXZVLWkiY4JAWwdk3nf8uSUDY7dlbDB0dRvbCwEH1iNrj8H6ZGp0iv5hT6F0gp3kbmSPQAAAABJRU5ErkJggg==", 
    layout="centered"
)

def validar_formulario(data):
    """Valida que todos los campos de calificación estén seleccionados"""
    campos_incompletos = []
    
    # Lista de campos que deben ser numéricos (no pueden ser string "Seleccione...")
    campos_calificacion = [
        ("p4_1_comunicacion", "4.1 - Comunicación"),
        ("p4_2_soporte", "4.2 - Soporte"),
        ("p5_1_resolucion_necesidades", "5.1 - Resolución de necesidades"),
        ("p5_2_cumplimiento_acuerdos", "5.2 - Cumplimiento de acuerdos"),
        ("p6_1_velocidad_respuesta", "6.1 - Velocidad de respuesta"),
        ("p6_2_eficiencia_operativa", "6.2 - Eficiencia operativa"),
        ("p7_1_integridad_etica", "7.1 - Integridad ética"),
        ("p7_2_equipo_humano", "7.2 - Equipo humano"),
        ("p8_1_comprension_necesidades", "8.1 - Comprensión de necesidades"),
        ("p8_2_sensibilidad", "8.2 - Sensibilidad"),
        ("p9_1_oficinas_espacios", "9.1 - Oficinas y espacios"),
        ("p9_2_instalaciones_servicios", "9.2 - Instalaciones y servicios"),
        ("p10_tic_tiempo", "10 - TIC Tiempo"),
        ("p10_tic_calidad", "10 - TIC Calidad"),
        ("p10_tic_atencion", "10 - TIC Atención"),
        ("p11_oper_tiempo", "11 - Operaciones Tiempo"),
        ("p11_oper_calidad", "11 - Operaciones Calidad"),
        ("p11_oper_atencion", "11 - Operaciones Atención"),
        ("p12_comercial_tiempo", "12 - Comercial Tiempo"),
        ("p12_comercial_calidad", "12 - Comercial Calidad"),
        ("p12_comercial_atencion", "12 - Comercial Atención"),
        ("p13_admin_tiempo", "13 - Admin Tiempo"),
        ("p13_admin_calidad", "13 - Admin Calidad"),
        ("p13_admin_atencion", "13 - Admin Atención"),
        ("p14_gci_tiempo", "14 - GCI Tiempo"),
        ("p14_gci_calidad", "14 - GCI Calidad"),
        ("p14_gci_atencion", "14 - GCI Atención"),
        ("p15_mobiliario", "15 - Mobiliario"),
        ("p15_limpieza", "15 - Limpieza"),
        ("p15_aire_acondicionado", "15 - Aire acondicionado"),
        ("p15_telefonia", "15 - Telefonía"),
        ("p15_internet_datos", "15 - Internet"),
        ("p15_impresiones", "15 - Impresiones"),
        ("p15_atencion_cliente", "15 - Atención al cliente"),
        ("p15_sala_reuniones", "15 - Sala de reuniones"),
        ("p16_registro_ingreso", "16 - Registro e ingreso"),
        ("p16_seguridad", "16 - Seguridad"),
        ("p16_areas_esparcimiento", "16 - Áreas de esparcimiento"),
        ("p16_alquiler_espacios", "16 - Alquiler de espacios"),
        ("p16_actividades_bienestar", "16 - Actividades de bienestar"),
        ("p16_mantenimiento_limpieza", "16 - Mantenimiento"),
        ("p17_restaurante", "17 - Restaurante"),
        ("p17_cafe", "17 - Café"),
        ("p17_cajero", "17 - Cajero"),
        ("p17_vending", "17 - Vending"),
        ("p17_esparcimiento", "17 - Esparcimiento"),
    ]
    
    for campo, nombre in campos_calificacion:
        valor = data.get(campo)
        if isinstance(valor, str) and "Seleccione" in valor:
            campos_incompletos.append(nombre)
    
    return campos_incompletos

# --- FUNCIÓN GUARDAR EN SUPABASE (Solo para BC) ---
def guardar_en_supabase(datos,tabla):
    """
    Guarda los datos de la encuesta en Supabase.
    """
    try:
        from supabase import create_client, Client
        
        url = "https://ekyfwvxmkagwaonrbafk.supabase.co"
        key = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVreWZ3dnhta2Fnd2FvbnJiYWZrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjQ2NjE5MTIsImV4cCI6MjA4MDIzNzkxMn0.VD1QFtqxHAkfp1D_TQj4GUCD8YKzmu14oQMpiOkrDX0"
        supabase: Client = create_client(url, key)

        if tabla == "Cuestionario_BC":
            # Mapear los datos al esquema de la tabla Cuestionario_BC
            supabase_data = {
                # Datos generales
                "aceptacion_politica": datos.get("aceptacion_politica"),
                "razon_social": datos.get("razon_social"),
                "nombre_completo": datos.get("nombre_completo"),
                
                # Pregunta 4: Comunicación y Soporte
                "calificacion_comunicacion_soporte_comunicacion": datos.get("p4_1_comunicacion"),
                "calificacion_comunicacion_soporte_soporte": datos.get("p4_2_soporte"),
                
                # Pregunta 5: Fiabilidad
                "calificacion_fiabilidad_resuelve": datos.get("p5_1_resolucion_necesidades"),
                "calificacion_fiabilidad_cumplir": datos.get("p5_2_cumplimiento_acuerdos"),
                
                # Pregunta 6: Capacidad de respuesta
                "calificacion_capacidad_respuesta_velocidad": datos.get("p6_1_velocidad_respuesta"),
                "calificacion_capacidad_respuesta_capacidad": datos.get("p6_2_eficiencia_operativa"),
                
                # Pregunta 7: Confianza
                "calificacion_confianza_etica": datos.get("p7_1_integridad_etica"),
                "calificacion_confianza_humano": datos.get("p7_2_equipo_humano"),
                
                # Pregunta 8: Empatía
                "calificacion_empatia_comprende": datos.get("p8_1_comprension_necesidades"),
                "calificacion_empatia_demuestra": datos.get("p8_2_sensibilidad"),
                
                # Pregunta 9: Elementos Tangibles
                "calificacion_elementos_tangibles_oficinas": datos.get("p9_1_oficinas_espacios"),
                "calificacion_elementos_tangibles_instalaciones": datos.get("p9_2_instalaciones_servicios"),
                
                # Pregunta 10: Área TIC
                "calificacion_area_tecnologia_espera": datos.get("p10_tic_tiempo"),
                "calificacion_area_tecnologia_calidad": datos.get("p10_tic_calidad"),
                "calificacion_area_tecnologia_atencion": datos.get("p10_tic_atencion"),
                
                # Pregunta 11: Operaciones y Mantenimiento
                "calificacion_area_operaciones_tiempo": datos.get("p11_oper_tiempo"),
                "calificacion_area_operaciones_calidad": datos.get("p11_oper_calidad"),
                "calificacion_area_operaciones_atencion": datos.get("p11_oper_atencion"),
                
                # Pregunta 12: Comercial, Bienestar y ATC
                "calificacion_area_comercial_tiempo": datos.get("p12_comercial_tiempo"),
                "calificacion_area_comercial_calidad": datos.get("p12_comercial_calidad"),
                "calificacion_area_comercial_atencion": datos.get("p12_comercial_atencion"),
                
                # Pregunta 13: Administrativa y Financiera
                "calificacion_area_administrativa_tiempo": datos.get("p13_admin_tiempo"),
                "calificacion_area_administrativa_calidad": datos.get("p13_admin_calidad"),
                "calificacion_area_administrativa_atencion": datos.get("p13_admin_atencion"),
                
                # Pregunta 14: GCI y Control de Inventarios
                "calificacion_area_gci_tiempo": datos.get("p14_gci_tiempo"),
                "calificacion_area_gci_calidad": datos.get("p14_gci_calidad"),
                "calificacion_area_gci_atencion": datos.get("p14_gci_atencion"),
                
                # Pregunta 15: Servicios Business Center
                "calificacion_business_center_mobiliario": datos.get("p15_mobiliario"),
                "calificacion_business_center_calidad": datos.get("p15_limpieza"),
                "calificacion_business_center_aire": datos.get("p15_aire_acondicionado"),
                "calificacion_business_center_telefonia": datos.get("p15_telefonia"),
                "calificacion_business_center_internet": datos.get("p15_internet_datos"),
                "calificacion_business_center_impresiones": datos.get("p15_impresiones"),
                "calificacion_business_center_cliente": datos.get("p15_atencion_cliente"),
                "calificacion_business_center_reuniones": datos.get("p15_sala_reuniones"),
                
                # Pregunta 16: Servicios Generales Zonamerica
                "satisfaccion_servicios_zonamerica_proceso": datos.get("p16_registro_ingreso"),
                "satisfaccion_servicios_zonamerica_seguridad": datos.get("p16_seguridad"),
                "satisfaccion_servicios_zonamerica_comunes": datos.get("p16_areas_esparcimiento"),
                "satisfaccion_servicios_zonamerica_alquiler": datos.get("p16_alquiler_espacios"),
                "satisfaccion_servicios_zonamerica_actividades": datos.get("p16_actividades_bienestar"),
                "satisfaccion_servicios_zonamerica_mantenimiento": datos.get("p16_mantenimiento_limpieza"),
                
                # Pregunta 17: Necesidades del personal
                "satisfaccion_necesidades_personal_restaurante": datos.get("p17_restaurante"),
                "satisfaccion_necesidades_personal_cafe": datos.get("p17_cafe"),
                "satisfaccion_necesidades_personal_cajero": datos.get("p17_cajero"),
                "satisfaccion_necesidades_personal_vending": datos.get("p17_vending"),
                "satisfaccion_necesidades_personal_areas": float(datos.get("p17_esparcimiento", 0)),
                
                # Pregunta 18: Satisfacción general (texto porque la columna es text)
                "satisfaccion_general": str(datos.get("p18_satisfaccion_general")),
                
                # Preguntas 19-22: Observaciones
                "sugerencias_servicios_2025": datos.get("p19_servicios_2025"),
                "observaciones_servicios_generales": datos.get("p20_obs_servicios_generales"),
                "observaciones_business_center": datos.get("p21_obs_business_center"),
                "observaciones_amenities": datos.get("p22_obs_amenities"),
                
                # Pregunta 23-24: Cierre
                "experiencia_emocional": datos.get("p23_emocion_experiencia"),
                "nps_score": datos.get("p24_nps_score")
            }
            
            with st.spinner("Guardando respuestas..."):
                # Insertar en Supabase
                result = supabase.table(tabla).insert(supabase_data).execute()
            
            if result.data:
                st.success("✅ ¡Datos guardados exitosamente!")
                return True
            else:
                st.warning("⚠️ No se pudo confirmar el guardado de los datos")
                return False
    except Exception as e:
        st.error(f"❌ Error al guardar: {e}")
        # Mostrar datos para depuración en caso de error
        with st.expander("Ver datos que se intentaron guardar"):
            st.json(supabase_data)
        return False

# --- ENCUESTA 1: BUSINESS CENTER (BC) ---
def show_survey_bc():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col1:
        st.image("https://web.zonamerica.com/colombia/wp-content/themes/zaco/assets/img/logo.png", width=200)
    with col3:
        st.image("https://cdn.brandfetch.io/ida2XlnzHx/theme/dark/logo.svg?c=1bxid64Mup7aczewSAYMX&t=1680282084001", width=200)
    
    # Encabezado principal del formulario
    st.title("NPS Zonamerica 2025 - BC")
    st.header("¡Queremos conocer tu experiencia con nosotros!")
    
    # Mensaje introductorio
    st.markdown("""
    Este contenido lo creó el propietario del formulario. Los datos que envíes se enviarán al propietario del formulario.
    
    Adecco no se hace responsable de las prácticas de privacidad o seguridad de sus clientes, incluidas las que adopte el propietario de este formulario.
    
    Sus datos personales serán tratados conforme a los lineamientos establecidos en nuestra 
    [Política de Protección de Datos Personales](https://web.zonamerica.com/colombia/wp-content/uploads/2022/11/JUR.PO_.01-Politica-de-Tratamiento-de-Datos-Personales-v1.pdf)
    
    **Te invitamos a participar en esta encuesta.**

    Tus respuestas son muy valiosas para nosotros. A continuación, encontrarás una serie de preguntas en las que deberás seleccionar una opción en escalas del 1 al 5, del 1 al 10 según corresponda, así como preguntas abiertas
    Gracias por ayudarnos a mejorar.
    """)
    
    st.markdown("---")

    with st.form("form_bc"):
        # Pregunta 1: Política de datos
        st.markdown("<h4 style='font-size: 20px;'>1. Usted acepta la política de tratamiento de datos. (*)</h4>", unsafe_allow_html=True)
        aceptacion = st.radio("", ["Sí", "No"], horizontal=True, label_visibility="collapsed")

        # Pregunta 2: Razón social
        st.markdown("<h4 style='font-size: 20px;'>2. Escriba la razón social de su Organización/Empresa. (*)</h4>", unsafe_allow_html=True)
        razon_social = st.selectbox("Razón Social",  [
                                                        "Green Leaf",
                                                        "Inelco",
                                                        "Learn English",
                                                        "Recman",
                                                        "Speedcast",
                                                        "Wekall",
                                                        "Cobranza",
                                                        "Xmaasrt",
                                                        "Holberton",
                                                        "Lidarit",
                                                        "Factoramos",
                                                        "Optima Corp."
                                                    ],index=1)
        
        # Pregunta 3: Nombre
        st.markdown("<h4 style='font-size: 20px;'>3. Escriba su nombre y apellidos. (*)</h4>", unsafe_allow_html=True)
        nombre_completo = st.text_input("Nombre completo", label_visibility="collapsed")

        st.markdown("---")
        
        # Pregunta 4: Comunicación y Soporte
        st.markdown("<h4 style='font-size: 20px;'>4. Comunicación y Soporte</h4>", unsafe_allow_html=True)
        st.caption("En una escala del 1 al 5, siendo 1=muy insatisfecho y 5=muy satisfecho, califica: (*)")
        
        st.markdown("**4.1** ¿Cómo evalúas la comunicación proporcionada por la administración de la Zonamerica?")
        comunicacion_4_1 = st.selectbox("4.1", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="com_4_1", label_visibility="collapsed")
        
        st.markdown("**4.2** ¿Estás satisfecho con el nivel de soporte ofrecido por la zona franca?")
        comunicacion_4_2 = st.selectbox("4.2", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="com_4_2", label_visibility="collapsed")

        # Pregunta 5: Fiabilidad
        st.markdown("<h4 style='font-size: 20px;'>5. Fiabilidad</h4>", unsafe_allow_html=True)
        st.caption("En una escala del 1 al 5, siendo 1=muy insatisfecho y 5=muy satisfecho, califica: (*)")
        
        st.markdown("**5.1** Zonamerica resuelve íntegramente sus necesidades como usuario/cliente.")
        fiabilidad_5_1 = st.selectbox("5.1", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="fia_5_1", label_visibility="collapsed")
        
        st.markdown("**5.2** La capacidad de Zonamerica para cumplir con los acuerdos establecidos ha influido en la toma de decisiones estratégicas de tu empresa.")
        fiabilidad_5_2 = st.selectbox("5.2", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="fia_5_2", label_visibility="collapsed")

        # Pregunta 6: Capacidad de respuesta
        st.markdown("<h4 style='font-size: 20px;'>6. Capacidad de respuesta</h4>", unsafe_allow_html=True)
        st.caption("En una escala del 1 al 5, siendo 1=muy insatisfecho y 5=muy satisfecho, califica: (*)")
        
        st.markdown("**6.1** ¿Te sientes satisfecho/a con la velocidad y eficacia de la respuesta a tus solicitudes y necesidades por parte de la administración de Zonamerica?")
        capacidad_6_1 = st.selectbox("6.1", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="cap_6_1", label_visibility="collapsed")
        
        st.markdown("**6.2** ¿La capacidad de respuesta ágil de Zonamerica ha influido en la eficiencia operativa de tu empresa?")
        capacidad_6_2 = st.selectbox("6.2", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="cap_6_2", label_visibility="collapsed")

        # Pregunta 7: Confianza
        st.markdown("<h4 style='font-size: 20px;'>7. Confianza</h4>", unsafe_allow_html=True)
        st.caption("En una escala del 1 al 5, siendo 1=muy insatisfecho y 5=muy satisfecho, califica: (*)")
        
        st.markdown("**7.1** ¿La integridad y ética empresarial de Zonamerica han contribuido a consolidar la confianza de tu empresa en su relación?")
        confianza_7_1 = st.selectbox("7.1", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="con_7_1", label_visibility="collapsed")
        
        st.markdown("**7.2** ¿El equipo humano de Zonamerica, es amable y diligente e infunde confianza en usted?")
        confianza_7_2 = st.selectbox("7.2", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="con_7_2", label_visibility="collapsed")

        # Pregunta 8: Empatía
        st.markdown("<h4 style='font-size: 20px;'>8. Empatía</h4>", unsafe_allow_html=True)
        st.caption("En una escala del 1 al 5, siendo 1=muy insatisfecho y 5=muy satisfecho, califica: (*)")
        
        st.markdown("**8.1** Zonamerica comprende y se preocupa por las necesidades específicas de tu empresa.")
        empatia_8_1 = st.selectbox("8.1", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="emp_8_1", label_visibility="collapsed")
        
        st.markdown("**8.2** Zonamerica demuestra sensibilidad hacia las circunstancias individuales de las empresas establecidas en ella.")
        empatia_8_2 = st.selectbox("8.2", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="emp_8_2", label_visibility="collapsed")

        # Pregunta 9: Elementos Tangibles
        st.markdown("<h4 style='font-size: 20px;'>9. Elementos Tangibles</h4>", unsafe_allow_html=True)
        st.caption("En una escala del 1 al 5, siendo 1=muy insatisfecho y 5=muy satisfecho, califica: (*)")
        
        st.markdown("**9.1** Las oficinas y espacios de trabajo, cumplen con tus expectativas y necesidades empresariales.")
        tangibles_9_1 = st.selectbox("9.1", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="tan_9_1", label_visibility="collapsed")
        
        st.markdown("**9.2** La calidad de las instalaciones y servicios de Zonamerica ha impactado positivamente en la imagen y operación de tu empresa.")
        tangibles_9_2 = st.selectbox("9.2", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="tan_9_2", label_visibility="collapsed")

        st.markdown("---")
        
        # Pregunta 10: Área de Tecnología/TIC'S
        st.markdown("<h4 style='font-size: 20px;'>10. Área de Tecnología/TIC'S</h4>", unsafe_allow_html=True)
        st.caption("En una escala del 1 al 5, siendo 1=muy insatisfecho y 5=muy satisfecho, califica: (*)")
        col10_1, col10_2, col10_3 = st.columns(3)
        with col10_1:
            tic_tiempo = st.selectbox("Tiempo de espera", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="tic_tiempo")
        with col10_2:
            tic_calidad = st.selectbox("Calidad de servicio", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="tic_calidad")
        with col10_3:
            tic_atencion = st.selectbox("Atención del personal", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="tic_atencion")

        # Pregunta 11: Área de Operaciones y Equipos de Mantenimiento
        st.markdown("<h4 style='font-size: 20px;'>11. Área de Operaciones y Equipos de Mantenimiento</h4>", unsafe_allow_html=True)
        st.caption("En una escala del 1 al 5, siendo 1=muy insatisfecho y 5=muy satisfecho, califica: (*)")
        col11_1, col11_2, col11_3 = st.columns(3)
        with col11_1:
            oper_tiempo = st.selectbox("Tiempo de espera", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="oper_tiempo")
        with col11_2:
            oper_calidad = st.selectbox("Calidad de servicio", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="oper_calidad")
        with col11_3:
            oper_atencion = st.selectbox("Atención del personal", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="oper_atencion")

        # Pregunta 12: Área Comercial, Bienestar y ATC
        st.markdown("<h4 style='font-size: 20px;'>12. Área Comercial, Bienestar y ATC</h4>", unsafe_allow_html=True)
        st.caption("En una escala del 1 al 5, siendo 1=muy insatisfecho y 5=muy satisfecho, califica: (*)")
        col12_1, col12_2, col12_3 = st.columns(3)
        with col12_1:
            com_tiempo = st.selectbox("Tiempo de espera", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="com_tiempo")
        with col12_2:
            com_calidad = st.selectbox("Calidad de servicio", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="com_calidad")
        with col12_3:
            com_atencion = st.selectbox("Atención del personal", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="com_atencion")

        # Pregunta 13: Área Administrativa y Financiera
        st.markdown("<h4 style='font-size: 20px;'>13. Área Administrativa y Financiera</h4>", unsafe_allow_html=True)
        st.caption("En una escala del 1 al 5, siendo 1=muy insatisfecho y 5=muy satisfecho, califica: (*)")
        col13_1, col13_2, col13_3 = st.columns(3)
        with col13_1:
            admin_tiempo = st.selectbox("Tiempo de espera", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="admin_tiempo")
        with col13_2:
            admin_calidad = st.selectbox("Calidad de servicio", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="admin_calidad")
        with col13_3:
            admin_atencion = st.selectbox("Atención del personal", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="admin_atencion")

        # Pregunta 14: Área GCI y Control de Inventarios
        st.markdown("<h4 style='font-size: 20px;'>14. Área GCI y Control de Inventarios</h4>", unsafe_allow_html=True)
        st.caption("En una escala del 1 al 5, siendo 1=muy insatisfecho y 5=muy satisfecho, califica: (*)")
        col14_1, col14_2, col14_3 = st.columns(3)
        with col14_1:
            gci_tiempo = st.selectbox("Tiempo de espera", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="gci_tiempo")
        with col14_2:
            gci_calidad = st.selectbox("Calidad de servicio", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="gci_calidad")
        with col14_3:
            gci_atencion = st.selectbox("Atención del personal", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="gci_atencion")

        # Pregunta 15: Servicios Business Center
        st.markdown("<h4 style='font-size: 20px;'>15. Servicios Business Center</h4>", unsafe_allow_html=True)
        st.caption("En una escala del 1 al 5, siendo 1=muy insatisfecho y 5=muy satisfecho, califica: (*)")
        col15_a, col15_b = st.columns(2)
        with col15_a:
            bc_mobiliario = st.selectbox("Mobiliario", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="bc_mob")
            bc_limpieza = st.selectbox("Calidad de Limpieza", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="bc_limp")
            bc_aire = st.selectbox("Aire Acondicionado", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="bc_aire")
            bc_telefonia = st.selectbox("Servicio de Telefonía", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="bc_tel")
        with col15_b:
            bc_internet = st.selectbox("Servicio de Internet y Datos", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="bc_int")
            bc_impresiones = st.selectbox("Servicios de Impresiones", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="bc_imp")
            bc_atencion = st.selectbox("Calidad y Disponibilidad del equipo de Atención al Cliente", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="bc_atc")
            bc_salas = st.selectbox("Calidad y Disponibilidad del servicio de sala de reuniones", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="bc_salas")

        # Pregunta 16: Nivel de satisfacción servicios generales
        st.markdown("<h4 style='font-size: 20px;'>16. Nivel de satisfacción - Servicios Generales</h4>", unsafe_allow_html=True)
        st.caption("En términos generales, ¿Cuál diría que es el Nivel de satisfacción de su empresa en relación a los siguientes servicios que ofrece Zonamerica? (*)")
        col16_a, col16_b = st.columns(2)
        with col16_a:
            gen_registro = st.selectbox("Proceso de Registro e Ingreso de Funcionarios, Visitantes y Proveedores", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="gen_reg")
            gen_seguridad = st.selectbox("Seguridad en general", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="gen_seg")
            gen_areas = st.selectbox("Áreas comunes de esparcimiento (internas y externas)", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="gen_areas")
        with col16_b:
            gen_alquiler = st.selectbox("Alquiler de espacios y equipos para actividades de cada usuario/cliente", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="gen_alq")
            gen_bienestar = st.selectbox("Actividades promovidas por Zonamerica para el bienestar de la comunidad", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="gen_bien")
            gen_mantenimiento = st.selectbox("Mantenimiento y limpieza de la infraestructura de Zonamerica", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="gen_mant")

        # Pregunta 17: Nivel de satisfacción - Necesidades del personal
        st.markdown("<h4 style='font-size: 20px;'>17. Necesidades del Personal</h4>", unsafe_allow_html=True)
        st.caption("Teniendo en cuenta las necesidades de su personal, por favor califique su Nivel de Satisfacción: (*)")
        col17_a, col17_b = st.columns(2)
        with col17_a:
            pers_restaurante = st.selectbox("Restaurante", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="pers_rest")
            pers_cafe = st.selectbox("Café", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="pers_cafe")
            pers_cajero = st.selectbox("Cajero", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="pers_cajero")
        with col17_b:
            pers_vending = st.selectbox("Vending Machine", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="pers_vend")
            pers_esparcimiento = st.selectbox("Áreas de esparcimiento en general", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="pers_espa")

        st.markdown("---")

        # Pregunta 18: Satisfacción general
        st.markdown("<h4 style='font-size: 20px;'>18. ¿En general, qué tan satisfecho está con Zonamerica? (*)</h4>", unsafe_allow_html=True)
        satisfaccion_general = st.selectbox("Satisfacción general", ["Seleccione el puntaje", 1, 2, 3, 4, 5], index=0, key="sat_gen", label_visibility="collapsed")

        # Pregunta 19: Servicios nuevos 2025
        st.markdown("<h4 style='font-size: 20px;'>19. ¿Qué tipo de servicios nuevos o mejorados le gustaría ver de nuestra empresa en 2025? (*)</h4>", unsafe_allow_html=True)
        servicios_2025 = st.text_area("Servicios 2025", key="serv_2025", label_visibility="collapsed")

        # Pregunta 20: Observaciones servicios generales
        st.markdown("<h4 style='font-size: 20px;'>20. Observaciones - Servicios Generales (*)</h4>", unsafe_allow_html=True)
        st.caption("Con respecto a los Servicios Generales, por favor amplíe sus observaciones y/o sugerencias y/o empresas de apoyo.")
        obs_servicios_generales = st.text_area("Observaciones generales", key="obs_gen", label_visibility="collapsed")

        # Pregunta 21: Observaciones Business Center
        st.markdown("<h4 style='font-size: 20px;'>21. Observaciones - Business Center (*)</h4>", unsafe_allow_html=True)
        st.caption("Con respecto a los Servicios Business Center por favor amplíe sus observaciones y/o sugerencias.")
        obs_bc = st.text_area("Observaciones BC", key="obs_bc", label_visibility="collapsed")

        # Pregunta 22: Observaciones Amenities
        st.markdown("<h4 style='font-size: 20px;'>22. Observaciones - Amenities (*)</h4>", unsafe_allow_html=True)
        st.caption("Con respecto a los Amenities y/o empresas de apoyo, por favor amplíe sus observaciones y/o sugerencias.")
        obs_amenities = st.text_area("Observaciones amenities", key="obs_amen", label_visibility="collapsed")

        # Pregunta 23: Experiencia emocional
        st.markdown("<h4 style='font-size: 20px;'>23. ¿Cómo se siente acerca de su experiencia general con Zonamerica? (*)</h4>", unsafe_allow_html=True)
        emocion = st.radio(
            "Emoción",
            ["Feliz 😄", "Encantado 🤩", "Neutral 😐", "Decepcionado 😞", "Molesto 😠"],
            horizontal=True, key="emocion", label_visibility="collapsed"
        )

        # Pregunta 24: NPS
        st.markdown("<h4 style='font-size: 20px;'>24. De 1 a 10, ¿Qué tanto recomendaría a sus colegas o amigos Zonamerica? (*)</h4>", unsafe_allow_html=True)
        nps = st.selectbox("NPS", options=range(1, 11), index=8, key="nps", label_visibility="collapsed")

        submitted = st.form_submit_button("Enviar Encuesta BC", type="primary", use_container_width=True)

        if submitted:
            if aceptacion == "No":
                st.error("⚠️ Debe aceptar la política de tratamiento de datos para continuar.")
            elif not razon_social or not nombre_completo:
                st.error("⚠️ Por favor complete la Razón Social y su Nombre.")
            else:
                data_payload = {
                    "aceptacion_politica": aceptacion == "Sí",
                    "razon_social": razon_social,
                    "nombre_completo": nombre_completo,
                    "p4_1_comunicacion": comunicacion_4_1,
                    "p4_2_soporte": comunicacion_4_2,
                    "p5_1_resolucion_necesidades": fiabilidad_5_1,
                    "p5_2_cumplimiento_acuerdos": fiabilidad_5_2,
                    "p6_1_velocidad_respuesta": capacidad_6_1,
                    "p6_2_eficiencia_operativa": capacidad_6_2,
                    "p7_1_integridad_etica": confianza_7_1,
                    "p7_2_equipo_humano": confianza_7_2,
                    "p8_1_comprension_necesidades": empatia_8_1,
                    "p8_2_sensibilidad": empatia_8_2,
                    "p9_1_oficinas_espacios": tangibles_9_1,
                    "p9_2_instalaciones_servicios": tangibles_9_2,
                    "p10_tic_tiempo": tic_tiempo,
                    "p10_tic_calidad": tic_calidad,
                    "p10_tic_atencion": tic_atencion,
                    "p11_oper_tiempo": oper_tiempo,
                    "p11_oper_calidad": oper_calidad,
                    "p11_oper_atencion": oper_atencion,
                    "p12_comercial_tiempo": com_tiempo,
                    "p12_comercial_calidad": com_calidad,
                    "p12_comercial_atencion": com_atencion,
                    "p13_admin_tiempo": admin_tiempo,
                    "p13_admin_calidad": admin_calidad,
                    "p13_admin_atencion": admin_atencion,
                    "p14_gci_tiempo": gci_tiempo,
                    "p14_gci_calidad": gci_calidad,
                    "p14_gci_atencion": gci_atencion,
                    "p15_mobiliario": bc_mobiliario,
                    "p15_limpieza": bc_limpieza,
                    "p15_aire_acondicionado": bc_aire,
                    "p15_telefonia": bc_telefonia,
                    "p15_internet_datos": bc_internet,
                    "p15_impresiones": bc_impresiones,
                    "p15_atencion_cliente": bc_atencion,
                    "p15_sala_reuniones": bc_salas,
                    "p16_registro_ingreso": gen_registro,
                    "p16_seguridad": gen_seguridad,
                    "p16_areas_esparcimiento": gen_areas,
                    "p16_alquiler_espacios": gen_alquiler,
                    "p16_actividades_bienestar": gen_bienestar,
                    "p16_mantenimiento_limpieza": gen_mantenimiento,
                    "p17_restaurante": pers_restaurante,
                    "p17_cafe": pers_cafe,
                    "p17_cajero": pers_cajero,
                    "p17_vending": pers_vending,
                    "p17_esparcimiento": pers_esparcimiento,
                    "p18_satisfaccion_general": satisfaccion_general,
                    "p19_servicios_2025": servicios_2025,
                    "p20_obs_servicios_generales": obs_servicios_generales,
                    "p21_obs_business_center": obs_bc,
                    "p22_obs_amenities": obs_amenities,
                    "p23_emocion_experiencia": emocion,
                    "p24_nps_score": nps
                }

                # Validar campos incompletos
                campos_faltantes = validar_formulario(data_payload)
                
                if campos_faltantes:
                    st.error(f"⚠️ Por favor complete todas las calificaciones. Faltan {len(campos_faltantes)} campo(s):")
                    with st.expander("Ver campos faltantes"):
                        for campo in campos_faltantes:
                            st.write(f"- {campo}")
                else:
                    if guardar_en_supabase(data_payload, "Cuestionario_BC"):
                        st.balloons()
                        time.sleep(2)

if __name__ == '__main__':
    show_survey_bc()