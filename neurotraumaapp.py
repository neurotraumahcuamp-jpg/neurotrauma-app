import streamlit as st
import pandas as pd
from datetime import datetime, date
import io
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Neurotrauma App", layout="wide", page_icon="🧠")

# --- DIRECTORIO DE USUARIOS Y ROLES ---
USUARIOS = {
    "mcorona":    {"nombre": "Dra. María Corona",     "rol": "CEO",       "cat": "R3 / Creadora",    "pass": "NeuroCEO2026*"},
    "warrieche":  {"nombre": "Dr. William Arrieche",  "rol": "Jefatura",  "cat": "Jefe de Servicio", "pass": "HCUAMP2026"},
    "vgonzalez":  {"nombre": "Dr. Víctor González",   "rol": "Adjunto",   "cat": "Adjunto",          "pass": "HCUAMP2026"},
    "ehernandez": {"nombre": "Dra. Eliana Hernández", "rol": "Adjunto",   "cat": "Adjunto",          "pass": "HCUAMP2026"},
    "lboschetti": {"nombre": "Dr. Levino Boschetti",  "rol": "Adjunto",   "cat": "Adjunto",          "pass": "HCUAMP2026"},
    "guerrero":   {"nombre": "Dr. Guerrero",          "rol": "Adjunto",   "cat": "Adjunto",          "pass": "HCUAMP2026"},
    "fmansaray":  {"nombre": "Dr. Foday Mansaray",    "rol": "Adjunto",   "cat": "Adjunto (R5)",     "pass": "HCUAMP2026"},
    "mortega":    {"nombre": "Dr. Michael Ortega",    "rol": "Adjunto",   "cat": "Adjunto (R5)",     "pass": "HCUAMP2026"},
    "alopez":     {"nombre": "Dr. Alejandro López",   "rol": "Adjunto",   "cat": "Adjunto (R5)",     "pass": "HCUAMP2026"},
    "dalvarado":  {"nombre": "Dr. Daniel Alvarado",   "rol": "Adjunto",   "cat": "Adjunto (R5)",     "pass": "HCUAMP2026"},
    "kmejias":    {"nombre": "Dr. Kenny Mejías",      "rol": "Adjunto",   "cat": "Adjunto (R5)",     "pass": "HCUAMP2026"},
    "yperez":     {"nombre": "Dr. Yeixon Pérez",      "rol": "Residente", "cat": "R4",               "pass": "NeuroR4"},
    "jalvarez":   {"nombre": "Dr. Jhorbis Álvarez",   "rol": "Residente", "cat": "R3",               "pass": "NeuroR3"},
    "rtelleria":  {"nombre": "Dra. Rudy Tellería",    "rol": "Residente", "cat": "R2",               "pass": "NeuroR2"},
    "ocartaya":   {"nombre": "Dr. Osman Cartaya",     "rol": "Residente", "cat": "R2",               "pass": "NeuroR2"},
    "mdun":       {"nombre": "Dra. Marialys Dun",     "rol": "Residente", "cat": "R1",               "pass": "NeuroR1"},
    "jolivar":    {"nombre": "Dr. Jonathan Olivar",   "rol": "Residente", "cat": "R1",               "pass": "NeuroR1"},
    "avalera":    {"nombre": "Dr. Alejandro Valera",  "rol": "Residente", "cat": "R1",               "pass": "NeuroR1"},
}

# ==========================================
# --- FUNCIONES CIENTÍFICAS EXACTAS ---
# ==========================================
def calcular_pam(tas, tad):
    if tas == 0 or tad == 0:
        return 0
    return int((2 * tad + tas) / 3)

def calcular_ppc(pam, pic):
    return int(pam - pic)

def calcular_indice_choque(fc, tas):
    if tas == 0:
        return 0
    return fc / tas

def calcular_volumen_hematoma(a, b, c):
    return (a * b * c) / 2

def calcular_rotterdam_exacto(cisternas, desviacion_mm, tiene_epidural, hsa_estado, hiv_estado):
    score = 1
    if "Comprimidas" in cisternas:
        score += 1
    elif cisternas == "Ausentes":
        score += 2
    if desviacion_mm > 5.0:
        score += 1
    if not tiene_epidural:
        score += 1
    if hsa_estado != "Ausente" or hiv_estado == "Presente":
        score += 1
    return min(score, 6)

def calcular_marshall_exacto(tc_normal, cisternas, desviacion_mm, vol_max, evacuada):
    if tc_normal:
        return 1, "Grado I (Sin Patología Visible)"
    if evacuada == "SI":
        return 5, "Grado V (Masa Evacuada)"
    if vol_max > 25.0:
        return 6, "Grado VI (Masa No Evacuada >25 cc)"
    if desviacion_mm > 5.0:
        return 4, "Grado IV (Lesión Difusa, Shift >5 mm)"
    if "Comprimidas" in cisternas or cisternas == "Ausentes":
        return 3, f"Grado III (Edema / Cisternas {cisternas})"
    return 2, "Grado II (Lesión Difusa, Cisternas Presentes)"

def calcular_greene_exacto(hsa_estado, efecto_masa, desviacion_mm):
    if hsa_estado == "Ausente":
        return "0", "No aplica"
    es_mayor_5 = "> 5 mm" in hsa_estado
    desv_mayor_5 = desviacion_mm > 5.0

    if not efecto_masa:
        if not es_mayor_5:
            return "1", "GRADO 1: HSA igual o menor a 5 mm"
        else:
            return "2", "GRADO 2: HSA > 5 mm"
    else:
        if not es_mayor_5:
            if not desv_mayor_5:
                return "3A", "GRADO 3A: HSA ≤ 5 mm, desviación < 5 mm y efecto de masa"
            else:
                return "3B", "GRADO 3B: HSA ≤ 5 mm, desviación > 5 mm y efecto de masa"
        else:
            if not desv_mayor_5:
                return "4A", "GRADO 4A: HSA > 5 mm, desviación < 5 mm y efecto de masa"
            else:
                return "4B", "GRADO 4B: HSA > 5 mm, desviación > 5 mm y efecto de masa"

def evaluar_conducta_clinica(marshall_score, vol_epi, gros_sub, vol_ip, vol_cont, shift_mm, tlics_score):
    criterios_qx = []
    if vol_epi > 30:
        criterios_qx.append(f"Hematoma Epidural > 30 cc ({vol_epi:.1f} cc)")
    if gros_sub > 10:
        criterios_qx.append(f"Hematoma Subdural > 10 mm ({gros_sub:.1f} mm)")
    if vol_ip > 25:
        criterios_qx.append(f"Hematoma Intraparenquimatoso > 25 cc ({vol_ip:.1f} cc)")
    if vol_cont > 25:
        criterios_qx.append(f"Contusión con efecto de masa > 25 cc ({vol_cont:.1f} cc)")
    if shift_mm > 5:
        criterios_qx.append(f"Desviación de Línea Media > 5 mm ({shift_mm:.1f} mm)")
    if tlics_score >= 5:
        criterios_qx.append(f"Trauma Raquimedular Inestable (TLICS {tlics_score})")

    if criterios_qx or marshall_score in [4, 6]:
        return "QUIRURGICO", "🔴 POTENCIAL QUIRÚRGICO / CRITERIO QX (BTF)", criterios_qx
    elif marshall_score == 3:
        return "ANTIEDEMA", "🟠 ALTO RIESGO DE HIC (Marshall III) — Manejo Médico Antiedema / Neuroprotección (No Qx de entrada salvo HIC refractaria)", []
    elif marshall_score == 5:
        return "POSTQX", "🔵 POST-EVACUACIÓN (Marshall V) — Vigilancia UCI / Neuroprotección", []
    else:
        return "MEDICO", "🟢 BAJO RIESGO DE HIC (Marshall I–II) — Manejo Médico / Observación Neurológica", []

def clasificar_tce(gcs):
    if gcs >= 13:
        return "Leve"
    elif gcs >= 9:
        return "Moderado"
    else:
        return "Severo"

def calcular_tlics(morfologia, ligamento, neuro):
    score = 0
    if "Compresión" in morfologia:
        score += 1
    elif "Estallido" in morfologia:
        score += 2
    elif "Traslación" in morfologia:
        score += 3
    elif "Distracción" in morfologia:
        score += 4
    if "Indeterminado" in ligamento:
        score += 2
    elif "Roto" in ligamento:
        score += 3
    if "Raíz" in neuro:
        score += 2
    elif "Completa" in neuro:
        score += 2
    elif "Incompleta" in neuro:
        score += 3
    elif "Cauda" in neuro:
        score += 3
    return score

def encontrar_logo(nombre_base):
    for ext in [".png", ".jpg", ".jpeg"]:
        if os.path.exists(nombre_base + ext):
            return nombre_base + ext
    return None

def generar_excel_investigacion(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Base_Madre", index=False)

        total = len(df)
        fallecidos = len(df[df["Condicion_Egreso"] == "FALLECIDO"]) if "Condicion_Egreso" in df.columns else 0
        qx_total = len(df[df["Quirurgico_Si_No"] == "SI"]) if "Quirurgico_Si_No" in df.columns else 0
        txa_total = len(df[df["Acido_Tranexamico"] == "SI"]) if "Acido_Tranexamico" in df.columns else 0
        leve = len(df[df["Severidad_TCE"] == "Leve"]) if "Severidad_TCE" in df.columns else 0
        mod = len(df[df["Severidad_TCE"] == "Moderado"]) if "Severidad_TCE" in df.columns else 0
        sev = len(df[df["Severidad_TCE"] == "Severo"]) if "Severidad_TCE" in df.columns else 0
        motos = len(df[df["Mecanismo"].str.contains("Moto", na=False)]) if "Mecanismo" in df.columns else 0

        resumen = pd.DataFrame([
            {"Indicador": "Total Pacientes", "Cantidad": total, "Porcentaje (%)": 100.0 if total else 0.0},
            {"Indicador": "Fallecidos (Mortalidad)", "Cantidad": fallecidos, "Porcentaje (%)": round(fallecidos / total * 100, 1) if total else 0.0},
            {"Indicador": "Casos Quirúrgicos", "Cantidad": qx_total, "Porcentaje (%)": round(qx_total / total * 100, 1) if total else 0.0},
            {"Indicador": "Uso de Ácido Tranexámico", "Cantidad": txa_total, "Porcentaje (%)": round(txa_total / total * 100, 1) if total else 0.0},
            {"Indicador": "TCE Leve (GCS 13-15)", "Cantidad": leve, "Porcentaje (%)": round(leve / total * 100, 1) if total else 0.0},
            {"Indicador": "TCE Moderado (GCS 9-12)", "Cantidad": mod, "Porcentaje (%)": round(mod / total * 100, 1) if total else 0.0},
            {"Indicador": "TCE Severo (GCS ≤ 8)", "Cantidad": sev, "Porcentaje (%)": round(sev / total * 100, 1) if total else 0.0},
            {"Indicador": "Accidentes en Moto", "Cantidad": motos, "Porcentaje (%)": round(motos / total * 100, 1) if total else 0.0},
        ])
        resumen.to_excel(writer, sheet_name="Resumen_Epidemiologico", index=False)

        dicc = pd.DataFrame([
            {"Variable": "GCS_Ingreso", "Tipo": "Numérica (3-15)", "Descripcion": "Puntaje Glasgow al ingreso"},
            {"Variable": "Severidad_TCE", "Tipo": "Categórica", "Descripcion": "Leve (13-15), Moderado (9-12), Severo (≤8)"},
            {"Variable": "Tipo_TCE", "Tipo": "Categórica", "Descripcion": "Cerrado vs Abierto"},
            {"Variable": "Marshall", "Tipo": "Ordinal (1-6)", "Descripcion": "Clasificación tomográfica de Marshall"},
            {"Variable": "Rotterdam", "Tipo": "Numérica (1-6)", "Descripcion": "Score pronóstico de Rotterdam"},
            {"Variable": "Greene", "Tipo": "Ordinal (1-4B)", "Descripcion": "Escala de HSA postraumática de Greene (1995)"},
            {"Variable": "Conducta_Sugerida", "Tipo": "Categórica", "Descripcion": "QUIRURGICO / ANTIEDEMA / MEDICO / POSTQX"},
            {"Variable": "Acido_Tranexamico", "Tipo": "Dicotómica", "Descripcion": "Protocolo CRASH-3 (SI/NO)"},
            {"Variable": "Quirurgico_Si_No", "Tipo": "Dicotómica", "Descripcion": "Intervención neuroquirúrgica (SI/NO)"},
            {"Variable": "GOSE", "Tipo": "Ordinal (1-8)", "Descripcion": "Glasgow Outcome Scale Extended"},
        ])
        dicc.to_excel(writer, sheet_name="Diccionario_Variables", index=False)
    return output.getvalue()

# --- DATOS GLOBALES ---
municipios_data = {
    "Lara": ["Iribarren", "Palavecino", "Torres", "Jiménez", "Morán", "Crespo", "Andrés Eloy Blanco", "Urdaneta", "Simón Planas"],
    "Yaracuy": ["San Felipe", "Independencia", "Bastidas", "Bolívar", "Bruzual", "Cocorote", "La Trinidad", "Manuel Monge", "Nirgua", "Páez", "Peña", "Sucre", "Urachiche", "Veroes"],
    "Portuguesa": ["Guanare", "Araure", "Esteller", "Guanarito", "Monseñor José Vicente de Unda", "Ospino", "Páez", "Papelón", "San Genaro de Boconoíto", "San Rafael de Onoto", "Santa Rosalía", "Sucre", "Turén"],
    "Falcón": ["Miranda", "Carirubana", "Silva", "Bolívar", "Buchivacoa", "Cacique Manaure", "Colina", "Dabajuro", "Democracia", "Falcón", "Federación", "Jacura", "Los Taques", "Mauroa", "Monseñor Iturriza", "Palmasola", "Petit", "Píritu", "San Francisco", "Sucre", "Tocópero", "Unión", "Urumaco", "Zamora"],
    "Barinas": ["Barinas", "Alberto Arvelo Torrealba", "Andrés Eloy Blanco", "Antonio José de Sucre", "Arismendi", "Bolívar", "Cruz Paredes", "Ezequiel Zamora", "Obispos", "Pedraza", "Rojas", "Sosa"],
    "Trujillo": ["Trujillo", "Valera", "Andrés Bello", "Boconó", "Bolívar", "Candelaria", "Carache", "Escuque", "José Felipe Márquez Cañizales", "Juan Vicente Campo Elías", "La Ceiba", "Miranda", "Monte Carmelo", "Motatán", "Pampán", "Pampanito", "Rafael Rangel", "San Rafael de Carvajal", "Sucre", "Urdaneta"],
    "Zulia": ["Maracaibo", "San Francisco", "Cabimas", "Lagunillas", "Colón", "Mara", "Miranda", "Jesús Enrique Lossada", "Semprún", "Machiques", "La Cañada", "Baralt", "Rosario de Perijá", "Valmore Rodríguez", "Santa Rita", "Guajira", "Sucre", "Almirante Padilla", "Catatumbo", "Simón Bolívar", "Pulgar"]
}
codigos_tlf = ["0412", "0414", "0424", "0416", "0426", "0251 (Lara)", "0252", "0253", "0254 (Yaracuy)", "0255 (Port)", "0256", "0257", "0261 (Zulia)", "0268 (Falcón)", "0273 (Barinas)", "Otro"]
ocupaciones_comunes = ["Mototaxista / Delivery", "Comerciante", "Obrero / Construcción", "Agricultor", "Estudiante", "Oficios del Hogar", "Chofer / Conductor", "Profesional / Técnico", "Desempleado", "Otro"]
motivos_comunes = ["Traumatismo Craneoencefálico (TCE)", "Pérdida del Estado de Conciencia (LOC)", "Hecho Vial en Moto", "Cefalea Postraumática", "Herida en Scalp", "Convulsión Postraumática", "Déficit Motor / Focalización", "Vómitos Postraumáticos", "Trauma Raquimedular"]
loc_craneo = ["Frontal", "Temporal", "Parietal", "Occipital", "Frontotemporal", "Temporoparietal", "Frontoparietotemporal", "Fosa Posterior"]
loc_parenquima = ["Frontal", "Temporal", "Parietal", "Occipital", "Frontotemporal", "Ganglios Basales / Tálamo", "Tronco Encefálico", "Cerebelo"]
loc_subdural = ["Hemisférico", "Frontotemporoparietal", "Frontal", "Temporal", "Parietal", "Interhemisférico", "Tentorial", "Fosa Posterior"]

# --- PANTALLA DE LOGIN ---
def pantalla_login():
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🧠 Neurotrauma App")
        st.markdown("**Sistema Integral Clínico y de Investigación en Neurotrauma**")
        st.caption("Hospital Central Universitario Dr. Antonio María Pineda — Servicio de Neurocirugía")
        st.divider()
        with st.form("login_form"):
            user_in = st.text_input("Usuario").strip().lower()
            pass_in = st.text_input("Contraseña", type="password")
            if st.form_submit_button("🔐 Iniciar Sesión", use_container_width=True):
                if user_in in USUARIOS and USUARIOS[user_in]["pass"] == pass_in:
                    st.session_state["auth"] = True
                    st.session_state["user_info"] = USUARIOS[user_in]
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos.")
        st.markdown("---")
        st.caption("© Neurotrauma App | Creada por **Dra. María Corona** (CEO & Fundadora)")

def main():
    if not st.session_state.get("auth", False):
        pantalla_login()
        return

    user_info = st.session_state["user_info"]
    es_ceo = (user_info["rol"] == "CEO")

    st.title("🧠 Neurotrauma App")
    st.markdown("**Sistema Integral Clínico y de Investigación en Neurotrauma**")

    # ==============================================================================
    # 1. INICIALIZACIÓN GLOBAL DE VARIABLES
    # ==============================================================================
    lenguaje = "Conservado"; orientacion = "Orientado"
    cn_1 = "No eval"; cn_2_visual = "Conservada"; cn_2_campimetria = "Conservada"; cn_2_fondo = "No eval"
    cn_346_motilidad = "Conservada"; cn_346_pupilas = "Normo"; cn_5_sens = "Conservada"; cn_7 = "Simétrico"
    cn_8 = "Conservado"; cn_bajos = "Sin alt"; sens_sup = "Conservada"; sens_prof = "Conservada"
    nivel_sensitivo = "Sin nivel"; tono_muscular = "Normal"; cerebelo_txt = ["Sin alt"]
    signos_meningeos = []; reflejos_txt = "2+"; babinski = "Flexor (Normal)"

    rx_segmentos = []; rx_hallazgos = "Sin alteraciones óseas evidentes"
    sindrome_medular = "Ninguno"; tlics_score = 0; tlics_decision = "N/A"
    herniacion_syndrome = "Ausente"
    uploaded_images = []

    scalp_desc = "Sin heridas evidentes"; trauma_txt = "Sin estigmas"; piel_desc = "Indemne"
    tx_str = "S/P"; cv_str = "S/P"; abd_str = "S/P"; ext_str = "S/P"
    lab_hb = 0.0; lab_hto = 0.0; lab_wbc = 0; lab_plt = 0; lab_tp = 0.0; lab_tpt = 0.0; lab_inr = 0.0
    lab_glu = 0; lab_urea = 0; lab_crea = 0.0; lab_na = 140; lab_k = 4.0; lab_cl = 100; lab_ca = 8.5
    lab_vih = "Pendiente"; lab_vdrl = "Pendiente"
    tipo_va = "Espontánea"; detalles_va = ""; accesos_list = []
    exposicion_dural_osea = False

    tipo_qx = "N/A"; hallazgos_qx = ""; cirujano = ""; qx_tamano = "No aplica"; qx_dura = "No aplica"
    patologias_list = []; var_alc = "No"; var_tab = "No"; var_chimo = "No"; var_drogas = "No"
    total_gcs = 15; gcs_display = "15/15"; clasif_tce = "Leve"
    mm_d = 3; mm_i = 3; r_d = "+"; r_i = "+"; anisocoria = "NO"
    fd = "5/5"; fi = "5/5"; asia_score = "E"; nivel_lesion = "N/A"
    dvno = 0.0; dvno_metodo = "No realizado"; dvno_resumen = "No realizado"; ppc = 0; pic = 0

    # ==========================================
    # 2. BARRA LATERAL
    # ==========================================
    with st.sidebar:
        st.markdown(f"### 👤 {user_info['nombre']}")
        st.caption(f"Rol: **{user_info['rol']}** ({user_info['cat']})")
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            st.session_state["auth"] = False
            st.rerun()

        st.divider()
        st.header("1. Configuración")
        logo_h = encontrar_logo("logo_hospital")
        logo_s = encontrar_logo("logo_servicio")
        logo_h_up = st.file_uploader("Escudo Hospital", type=["png", "jpg"], key="lh")
        if logo_h_up:
            logo_h = logo_h_up
        elif logo_h:
            st.success("✅ Logo Hospital")
        logo_s_up = st.file_uploader("Escudo Servicio", type=["png", "jpg"], key="ls")
        if logo_s_up:
            logo_s = logo_s_up
        elif logo_s:
            st.success("✅ Logo Servicio")

        if es_ceo:
            st.divider()
            st.header("👑 BASE MADRE (CEO)")
            if os.path.exists("neuro_db.csv"):
                df_exp = pd.read_csv("neuro_db.csv")
                try:
                    excel_bytes = generar_excel_investigacion(df_exp)
                    st.download_button(
                        "📥 Bajar Base Madre (Excel .xlsx)",
                        excel_bytes,
                        f"Base_Madre_Neurotrauma_{date.today()}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception:
                    with open("neuro_db.csv", "rb") as f:
                        st.download_button("📥 Bajar Base Madre (CSV)", f, "DB_Neuro_HCUAMP.csv", "text/csv", use_container_width=True)
            else:
                st.info("Sin registros guardados aún.")

        st.divider()
        st.caption("Creada por **Dra. María Corona**")

    # ==========================================
    # 3. PESTAÑAS
    # ==========================================
    tabs = st.tabs([
        "👤 1. Ingreso",
        "📜 2. Historia",
        "🚨 3. Vitales/Físico",
        "🧠 4. Neuro/Imagen",
        "📈 5. Protocolo/Labs",
        "💊 6. Qx/Plan",
        "🏁 7. Egreso/GOSE",
        "📊 8. Dashboard"
    ])

    # --- TAB 1: INGRESO ---
    with tabs[0]:
        st.subheader("Filiación y Fecha de Ingreso")
        c_fecha, c_hora, c_med, c_hosp = st.columns([1, 1, 1.5, 2])
        fecha_ingreso = c_fecha.date_input("Fecha de Ingreso", datetime.now())
        hora_ingreso = c_hora.time_input("Hora de Ingreso", datetime.now().time())
        medico = c_med.text_input("Médico que examina:", value=user_info["nombre"])
        hospital = c_hosp.selectbox("Hospital:", ["Hospital Central Universitario Antonio María Pineda"])
        st.divider()
        c1, c2, c3, c4 = st.columns(4)
        nombre = c1.text_input("Nombre y Apellido")
        cedula = c2.text_input("Cédula")
        edad = c3.number_input("Edad", min_value=0, value=25)
        sexo = c4.selectbox("Sexo", ["Masculino", "Femenino"])

        co1, co2 = st.columns(2)
        ocupacion_sel = co1.selectbox("Ocupación / Oficio", ocupaciones_comunes)
        ocupacion = co2.text_input("Especifique Ocupación") if ocupacion_sel == "Otro" else ocupacion_sel

        st.markdown("**Procedencia y Traslado**")
        cp1, cp2 = st.columns(2)
        estado = cp1.selectbox("Estado", ["Lara", "Yaracuy", "Portuguesa", "Falcón", "Barinas", "Trujillo", "Zulia", "Otro"])
        if estado in municipios_data:
            municipio = cp2.selectbox(f"Municipio ({estado})", municipios_data[estado])
        else:
            municipio = cp2.text_input("Municipio")
        traslado = st.selectbox("Medio de Traslado", ["Ambulancia (Bomberos/PC/Cruz Roja)", "Vehículo Particular", "Patrulla Policial/Militar", "Moto Particular", "Otro"])
        prehospitalaria = st.checkbox("¿Recibió Atención Prehospitalaria? (Collarín, Vía, O2)")

        st.divider()
        st.subheader("📞 Contacto y Emergencia")
        ct1, ct2 = st.columns([1, 3])
        with ct1:
            prefijo_pac = st.selectbox("Código", codigos_tlf, key="pp")
        with ct2:
            numero_pac = st.text_input("Número Teléfono Paciente", placeholder="Ej: 5555555")
        telefono_full = f"{prefijo_pac}-{numero_pac}"

        st.markdown("**Familiar / Emergencia**")
        cf1, cf2, cf3 = st.columns([2, 1, 2])
        with cf1:
            nombre_fam = st.text_input("Nombre y Apellido Familiar")
        with cf2:
            parentesco = st.selectbox("Parentesco", ["Madre", "Padre", "Esposo/a", "Hijo/a", "Hermano/a", "Tío/a", "Nieto/a", "Amigo/a", "Otro"])
        with cf3:
            c_sub1, c_sub2 = st.columns([1, 2])
            prefijo_fam = c_sub1.selectbox("Cód", codigos_tlf, key="pf")
            numero_fam = c_sub2.text_input("Número Familiar", placeholder="Ej: 5555555")

        st.divider()
        st.markdown("⏱️ **Tiempo de Latencia**")
        delta_t = st.number_input("Horas de Evolución (Desde trauma hasta ingreso)", min_value=0.0, step=0.5)

    # --- TAB 2: HISTORIA ---
    with tabs[1]:
        st.subheader("Anamnesis")
        c_mot, c_mec = st.columns([2, 1])
        motivos_sel = c_mot.multiselect("Motivo de Consulta (Selección rápida)", motivos_comunes)
        mecanismo = c_mec.selectbox("Mecanismo", ["Moto-Casco", "Moto-Sin Casco", "Auto", "Arrollamiento", "Caída", "Arma Fuego", "Otro", "No Trauma"])
        motivo_libre = st.text_input("Otro Motivo de Consulta / Detalle")
        motivo = ", ".join(motivos_sel) + (f". {motivo_libre}" if motivo_libre else "")
        ea = st.text_area("Enfermedad Actual", height=100)

        st.markdown("---")
        st.subheader("Estado Mental al Trauma")
        c_intox, c_loc = st.columns(2)
        with c_intox:
            intoxicacion = st.selectbox("Estado Intoxicación Aguda", ["No aparente", "Aliento Etílico", "Ebriedad Evidente", "Drogas"])
        with c_loc:
            loc = st.radio("Pérdida de Consciencia (LOC)", ["NO", "SI (< 30 min)", "SI (> 30 min)", "Desconocido"], horizontal=True)
            amnesia = st.checkbox("Amnesia del Evento")

        st.markdown("---")
        st.subheader("Protocolo CRASH-3 (Ácido Tranexámico)")
        c_txa1, c_txa2 = st.columns(2)
        txa_admin = c_txa1.radio("¿Se administró Ácido Tranexámico (TXA)?", ["NO", "SI"], horizontal=True)
        txa_tiempo = c_txa2.selectbox("Tiempo desde el trauma", ["< 3 horas (Ideal)", "> 3 horas", "Desconocido"]) if txa_admin == "SI" else "No aplica"

        st.markdown("---")
        st.subheader("Antecedentes")
        col_ant1, col_ant2, col_ant3 = st.columns(3)
        with col_ant1:
            st.markdown("🔴 **Patológicos**")
            alergia_check = st.radio("Alergias", ["NO", "SI"], horizontal=True)
            alergia_txt = st.text_input("Medicamento:") if alergia_check == "SI" else ""
            if st.checkbox("HTA"):
                patologias_list.append(f"HTA ({st.text_input('Tto HTA:')})")
            if st.checkbox("Diabetes"):
                patologias_list.append(f"DM ({st.text_input('Tto DM:')})")
            if st.checkbox("Asma"):
                patologias_list.append(f"Asma ({st.text_input('Crisis:')})")
            st.markdown("⚠️ **Medicación de Riesgo**")
            anticoagulantes_list = st.multiselect("¿Anticoagulantes/Antiagregantes?", ["Niega", "Aspirina", "Clopidogrel", "Warfarina", "Rivaroxaban", "Enoxaparina", "Desconocido"])
            otras = st.text_input("Otras Enf.")
            if otras:
                patologias_list.append(otras)
            quirurgicos = st.text_input("Qx Previas")
        with col_ant2:
            st.markdown("👪 **Familiares**")
            m_st = st.radio("Madre", ["Viva", "Fallecida"], horizontal=True)
            m_enf = st.text_input("Enf. Madre", placeholder="Ej: Sana")
            p_st = st.radio("Padre", ["Vivo", "Fallecido"], horizontal=True)
            p_enf = st.text_input("Enf. Padre", placeholder="Ej: Sano")
            st.divider()
            h_yn = st.radio("¿Tiene Hermanos?", ["SI", "NO"], horizontal=True)
            h_txt = st.text_input("Detalle Hermanos (¿Sanos?)") if h_yn == "SI" else "No tiene"
            hi_yn = st.radio("¿Tiene Hijos?", ["SI", "NO"], horizontal=True)
            hi_txt = st.text_input("Detalle Hijos (¿Sanos?)") if hi_yn == "SI" else "No tiene"
        with col_ant3:
            st.markdown("☕ **Hábitos**")
            if st.checkbox("Alcohol (Social/Frecuente)"):
                var_alc = "Si"
            if st.checkbox("Tabaco"):
                var_tab = "Si"
            if st.checkbox("Chimó"):
                var_chimo = "Si"
            if st.checkbox("Drogas"):
                var_drogas = "Si"
            if st.checkbox("Café"):
                t_caf = st.number_input("¿Cuántas tazas al día?", min_value=1)
                hab_cafe_txt = f"Café ({t_caf} tazas/día)"
            else:
                hab_cafe_txt = ""

    # --- TAB 3: VITALES Y FÍSICO ---
    with tabs[2]:
        st.header("Evaluación Primaria (ATLS)")
        st.subheader("C - Circulación y Hemodinamia")
        c_v1, c_v2 = st.columns(2)
        with c_v1:
            tas = st.number_input("TAS", 50, 250, 120)
            tad = st.number_input("TAD", 30, 150, 80)
            fc = st.number_input("FC", 30, 200, 80)
            pam = calcular_pam(tas, tad)
            shock_index = calcular_indice_choque(fc, tas)
            st.metric("PAM", f"{pam} mmHg")
            st.metric("Índice de Choque (FC/TAS)", f"{shock_index:.2f}")
            if shock_index > 0.9:
                st.error("⚠️ ALERTA CHOQUE: Índice > 0.9 (Posible sangrado oculto)")
            if tas < 90:
                st.error("🚨 HIPOTENSIÓN DOCUMENTADA (Factor de mal pronóstico BTF)")
            if tas > 160 and fc < 60:
                st.error("⚠️ ALERTA: TRIADA DE CUSHING (HIC Inminente)")
        with c_v2:
            st.markdown("**B - Ventilación**")
            fr = st.number_input("Frecuencia Resp", 8, 60, 18)
            sat = st.number_input("SatO2 (%)", 50, 100, 98)
            if sat < 90:
                st.error("🚨 HIPOXIA DOCUMENTADA (Factor de mal pronóstico BTF)")
            temp = st.number_input("Temp (ºC)", 30.0, 42.0, 37.0)
            st.markdown("**Antropometría**")
            peso = st.number_input("Peso (Kg) Aprox", 1.0, 150.0, 70.0)
            talla = st.number_input("Talla (m)", 0.50, 2.20, 1.70)
            imc = peso / (talla * talla)
            st.caption(f"IMC: {imc:.1f}")

        st.divider()
        st.subheader("A - Vía Aérea")
        col_va1, col_va2, col_va3 = st.columns(3)
        tipo_va = col_va1.selectbox("Dispositivo Vía Aérea", ["Espontánea (Aire Ambiente)", "Mascarilla / Cánula", "Tubo Endotraqueal (TET)", "Traqueostomía"])
        if "Tubo" in tipo_va or "Traqueo" in tipo_va:
            num_tet = col_va2.selectbox("Número TET/TQT", ["6.0", "6.5", "7.0", "7.5", "8.0", "8.5", "9.0"])
            fijacion_tet = col_va2.number_input("Fijación (cm)", 18, 26, 22)
            modo_vent = col_va3.selectbox("Modo Ventilatorio", ["Asisto-Controlado", "SIMV", "CPAP/PS", "T-Piece"])
            fio2 = col_va3.number_input("FiO2 (%)", 21, 100, 40)
            detalles_va = f"Nº {num_tet} fijado a {fijacion_tet}cm. Modo: {modo_vent} FiO2 {fio2}%"
        else:
            detalles_va = f"FR: {fr} rpm | SatO2: {sat}%"

        st.divider()
        st.subheader("E - Exposición y Adjuntos")
        col_inv1, col_inv2 = st.columns(2)
        with col_inv1:
            if st.checkbox("Vía Periférica"):
                accesos_list.append("VVP permeable")
            if st.checkbox("Vía Central (CVC)"):
                sitio_cvc = st.selectbox("Sitio CVC", ["Subclavia Der", "Subclavia Izq", "Yugular Der", "Yugular Izq", "Femoral"])
                accesos_list.append(f"CVC ({sitio_cvc})")
            if st.checkbox("Línea Arterial"):
                accesos_list.append("Línea Arterial")
        with col_inv2:
            if st.checkbox("Sonda Vesical (Foley)"):
                diuresis = st.text_input("Características Diuresis", "Clara / Colúrica")
                accesos_list.append(f"Foley ({diuresis})")
            if st.checkbox("Sonda Nasogástrica"):
                accesos_list.append("SNG a derivación")

        st.divider()
        st.subheader("Examen Físico Segmentario")
        st.markdown("**💀 Cabeza / Signos de Base de Cráneo**")
        ct1, ct2, ct3 = st.columns(3)
        signos_txt = []
        with ct1:
            st.caption("Fosa Anterior")
            mapache = st.selectbox("Ojos de Mapache", ["Ausente", "Derecho", "Izquierdo", "Bilateral"])
            if mapache != "Ausente":
                signos_txt.append(f"Ojos de Mapache ({mapache})")
            if st.checkbox("Rinorrea / Rinorraquia"):
                signos_txt.append("Rinorrea")
        with ct2:
            st.caption("Fosa Media/Posterior")
            battle = st.selectbox("Signo de Battle", ["Ausente", "Derecho", "Izquierdo", "Bilateral"])
            if battle != "Ausente":
                signos_txt.append(f"Signo de Battle ({battle})")
            otorragia = st.selectbox("Otorragia / Otorrea", ["Ausente", "Derecha", "Izquierda", "Bilateral"])
            if otorragia != "Ausente":
                signos_txt.append(f"Otorragia/Otorrea ({otorragia})")
        with ct3:
            st.caption("Heridas")
            if st.checkbox("Herida en Scalp"):
                h_ubic = st.selectbox("Ubicación", ["Frontal", "Parietal", "Temporal", "Occipital", "Vertex"])
                h_long = st.number_input("Longitud (cm)", 1, 30, 5)
                h_prof = st.selectbox("Profundidad", ["Piel/Celular", "Galea/Aponeurosis", "Exposición Ósea", "Duramadre expuesta"])
                if h_prof in ["Exposición Ósea", "Duramadre expuesta"]:
                    exposicion_dural_osea = True
                h_estado = st.selectbox("Estado", ["Sangrado activo", "Coágulos", "Suturada/Afrontada", "Bordes irregulares"])
                scalp_desc = f"Herida en región {h_ubic} de {h_long}cm, plano: {h_prof}. {h_estado}."
                signos_txt.append(f"Herida Scalp {h_ubic}")
        trauma_txt = ", ".join(signos_txt) if signos_txt else "Normocéfalo, sin estigmas de trauma."

        st.markdown("**🩹 Piel y Faneras (General)**")
        piel_desc = st.text_area("Describir: Abrasiones, Equimosis (Cinturón, Volante), Laceraciones en cuerpo", "Piel morena/blanca, normotérmica, turgencia conservada. Sin otras lesiones evidentes.")

        with st.expander("Cardiopulmonar, Abdomen y Extremidades", expanded=True):
            ce1, ce2 = st.columns(2)
            with ce1:
                st.markdown("**Tórax y Respiratorio**")
                tx_exp = st.selectbox("Expansibilidad", ["Normoexpansible", "Hipoexpansible", "Tórax Inestable"])
                tx_mv = st.selectbox("Murmullo Vesicular", ["Audible en ambos campos", "Disminuido derecho", "Disminuido izquierdo", "Abolido"])
                tx_agreg = st.multiselect("Ruidos Agregados", ["Ninguno", "Crepitantes", "Roncus", "Sibilancias"])
                tx_str = f"{tx_exp}. MV: {tx_mv}. Agregados: {', '.join(tx_agreg) if tx_agreg else 'Niega'}."
                st.markdown("**Cardiovascular**")
                cv_rs = st.radio("Ruidos Cardíacos", ["Rítmicos", "Arrítmicos"], horizontal=True)
                cv_soplo = st.checkbox("Soplos")
                cv_str = f"RsCs {cv_rs}, {'con soplos' if cv_soplo else 'sin soplos'}."
            with ce2:
                st.markdown("**Abdomen**")
                abd_ruidos = st.radio("Ruidos Hidroaéreos", ["Presentes", "Ausentes/Abolidos"], horizontal=True)
                abd_palp = st.selectbox("Palpación", ["Blando/Depresible", "Doloroso", "Defensa/Tabla", "Distendido"])
                abd_str = f"RHA {abd_ruidos}, {abd_palp}."
                st.markdown("**Extremidades**")
                ext_sim = st.selectbox("Simetría/Edema", ["Simétricas s/edema", "Edema MsIs", "Asimétricas"])
                ext_pulsos = st.radio("Pulsos Periféricos", ["Presentes", "Débiles/Ausentes"], horizontal=True)
                ext_str = f"{ext_sim}, pulsos {ext_pulsos}."
            st.markdown("**Cuello**")
            cuello_opt = st.radio("Estado Cuello", ["Móvil/No doloroso", "Collarín Cervical Puesto", "Doloroso/Rigidez", "Ingurgitación Yugular"], horizontal=True)

    # --- TAB 4: NEUROLOGÍA, NEUROIMAGEN Y COLUMNA ---
    with tabs[3]:
        tab_clinica, tab_craneo, tab_columna = st.tabs(["🩺 Examen Clínico Detallado", "☢️ Cráneo (TC/RMN)", "🦴 Columna (TLICS/ASIA)"])

        # 1. CLÍNICA EXPANDIDA V57
        with tab_clinica:
            st.subheader("Evaluación Neurológica Integral")
            with st.expander("1. Estado Mental y Glasgow", expanded=True):
                c_gcs1, c_gcs2 = st.columns(2)
                with c_gcs1:
                    gao = st.selectbox("Ocular (O)", ["4-Espontánea", "3-Orden", "2-Dolor", "1-Nada"])
                    grv = st.selectbox("Verbal (V)", ["5-Orientado", "4-Confuso", "3-Palabras", "2-Sonidos", "1-Nada", "T-Tubo"])
                    grm = st.selectbox("Motor (M)", ["6-Obedece", "5-Localiza", "4-Flex.N", "3-Flex.A", "2-Ext", "1-Nada"])
                    es_intubado = grv.startswith("T")
                    total_gcs = int(gao[0]) + (1 if es_intubado else int(grv[0])) + int(grm[0])
                    gcs_display = f"{total_gcs}T" if es_intubado else f"{total_gcs}/15"
                    clasif_tce = clasificar_tce(total_gcs)
                    st.info(f"🧠 GCS: {gcs_display} (TCE {clasif_tce})")
                with c_gcs2:
                    lenguaje = st.selectbox("Lenguaje", ["Conservado", "Disartria", "Afasia Motora (Broca)", "Afasia Sensitiva", "Afasia Global", "Mutismo"])
                    orientacion = st.text_input("Orientación", "T/E/P Conservada")

            with st.expander("2. Pares Craneales (Detallado)", expanded=True):
                c_pc1, c_pc2 = st.columns(2)
                with c_pc1:
                    cn_1 = st.selectbox("I (Olfatorio)", ["No evaluado", "Conservado", "Anosmia", "Hiposmia"])
                    st.markdown("**II (Visual)**")
                    cn_2_visual = st.selectbox("Agudeza", ["Conservada", "Disminuida", "Amaurosis", "No evaluable"])
                    cn_2_campimetria = st.selectbox("Campimetría", ["Conservada", "Hemianopsia", "Cuadrantanopsia", "No evaluable"])
                    cn_2_fondo = st.selectbox("Fondo Ojo", ["No realizado", "Normal", "Papiledema", "Hemorragia Retiniana"])
                with c_pc2:
                    st.markdown("**III, IV, VI (Oculomotores)**")
                    mm_d = st.slider("Pupila Der (mm)", 1, 9, 3); r_d = st.selectbox("R. Der", ["+", "-"], key="rd")
                    mm_i = st.slider("Pupila Izq (mm)", 1, 9, 3); r_i = st.selectbox("R. Izq", ["+", "-"], key="ri")
                    anisocoria = "SI" if abs(mm_d - mm_i) >= 2 else "NO"
                    if anisocoria == "SI" or r_d == "-" or r_i == "-":
                        st.warning("⚠️ Alerta Pupilar: Anisocoria (≥ 2 mm) o arreactividad.")
                    cn_346_motilidad = st.selectbox("Motilidad", ["Conservada", "Paresia III", "Paresia VI", "Oftalmoplejía"])
                    cn_346_pupilas = f"OD {mm_d}mm {r_d} | OI {mm_i}mm {r_i}"

                st.divider()
                c_pc3, c_pc4 = st.columns(2)
                with c_pc3:
                    cn_5_sens = st.selectbox("V (Trigémino)", ["Sensibilidad Conservada", "Hipoestesia V1/V2/V3", "Neuralgia"])
                    cn_7 = st.selectbox("VII (Facial)", ["Simétrico", "Parálisis Central (Respeta frente)", "Parálisis Periférica (Bell)", "Asimetría leve"])
                    cn_8 = st.selectbox("VIII (Vestibulo-Coclear)", ["Conservado", "Hipoacusia", "Nistagmo", "Vértigo"])
                with c_pc4:
                    cn_bajos = st.selectbox("Pares Bajos (IX, X, XI, XII)", ["Sin alteraciones", "Disfagia/Disfonía", "Desviación Uvula/Lengua", "Debilidad Trapecio"])

            with st.expander("3. Motor, Sensitivo y Reflejos", expanded=True):
                cm1, cm2 = st.columns(2)
                with cm1:
                    st.markdown("**Fuerza (Daniels)**")
                    fd = st.selectbox("Fuerza Der", ["5/5", "4/5", "3/5", "2/5", "1/5", "0/5"])
                    fi = st.selectbox("Fuerza Izq", ["5/5", "4/5", "3/5", "2/5", "1/5", "0/5"])
                    tono_muscular = st.selectbox("Tono", ["Normal", "Hipotonía", "Espasticidad", "Rigidez"])
                with cm2:
                    st.markdown("**Sensibilidad**")
                    sens_sup = st.selectbox("Superficial (Táctil/Dolor)", ["Conservada", "Hipoestesia", "Anestesia", "Parestesias"])
                    sens_prof = st.selectbox("Profunda (Propiocepción)", ["Conservada", "Alterada"])
                    nivel_sensitivo = st.text_input("Nivel Sensitivo (Si aplica)", "Sin nivel")

                st.markdown("**Reflejos y Patológicos**")
                cr1, cr2 = st.columns(2)
                with cr1:
                    reflejos_txt = st.selectbox("Reflejos Osteotendinosos", ["2+ (Normal)", "0 (Arreflexia)", "1+ (Hiporreflexia)", "3+ (Hiperreflexia)", "4+ (Clonus)"])
                with cr2:
                    babinski = st.selectbox("Reflejo Plantar", ["Flexor (Normal)", "Extensor (Babinski Presente)", "Indiferente"])

            with st.expander("4. Cerebelo y Meninges"):
                cerebelo_txt = st.multiselect("Cerebelo", ["Sin alteraciones", "Dismetría", "Disdiadococinesia", "Ataxia", "Temblor"])
                signos_meningeos = st.multiselect("Meninges", ["Ausentes", "Rigidez Nuca", "Kernig", "Brudzinski"])

            herniacion_syndrome = st.selectbox("Síndrome Herniación", ["Ausente", "Uncal", "Central", "Amigdalina"])
            if herniacion_syndrome != "Ausente":
                st.error(f"⚠️ URGENCIA: {herniacion_syndrome}")

        # 2. CRÁNEO / IMAGEN (DISEÑO NEUROQUIRÚRGICO EXACTO)
        with tab_craneo:
            st.subheader("☢️ Neuroimagen (TC de Cráneo)")
            with st.expander("📷 Técnica y Archivos", expanded=False):
                c_tec1, c_tec2, c_tec3 = st.columns(3)
                fecha_estudio_img = c_tec1.date_input("Fecha Estudio", datetime.now())
                tc_formato = c_tec2.selectbox("Formato", ["Digital", "CD", "Placas físicas", "Pendiente por TC"])
                tc_contraste = c_tec3.radio("Fase", ["Simple", "Con Contraste"], horizontal=True)
                tc_cortes = st.multiselect("Cortes", ["Axial", "Coronal", "Sagital"], default=["Axial", "Sagital", "Coronal"])
                uploaded_images = st.file_uploader("Subir Fotos", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
                if uploaded_images:
                    cols = st.columns(3)
                    for i, img_file in enumerate(uploaded_images):
                        cols[i % 3].image(img_file, caption=f"Img {i+1}", use_column_width=True)

            tc_normal = st.checkbox("✅ TC Normal (Sin patología intracraneal visible — Marshall I)")

            st.markdown("#### 1. Cisternas Basales y Línea Media")
            c_cl1, c_cl2, c_cl3 = st.columns([2, 1.5, 1.5])
            tc_cisternas = c_cl1.selectbox(
                "Cisternas Basales",
                ["Normales", "Comprimidas (Unilateral)", "Comprimidas (Bilateral)", "Ausentes"],
                disabled=tc_normal
            )
            desviacion_lm_mm = c_cl2.number_input("Desviación Línea Media (mm)", 0.0, 30.0, 0.0, step=0.5, disabled=tc_normal)
            lado_shift = c_cl3.selectbox("Desviada hacia", ["Derecha", "Izquierda"], disabled=(tc_normal or desviacion_lm_mm == 0))

            st.markdown("#### 2. Lesiones de Masa (Marque las presentes)")
            cb_m1, cb_m2, cb_m3, cb_m4 = st.columns(4)
            chk_epi = cb_m1.checkbox("Hematoma Epidural", disabled=tc_normal)
            chk_sub = cb_m2.checkbox("Hematoma Subdural", disabled=tc_normal)
            chk_ip  = cb_m3.checkbox("H. Intraparenquimatoso", disabled=tc_normal)
            chk_cnt = cb_m4.checkbox("Contusión Hemorrágica", disabled=tc_normal)

            vol_epi = 0.0; gros_sub = 0.0; vol_sub = 0.0; vol_ip = 0.0; vol_cont = 0.0
            lesiones_detalles = []

            if not tc_normal and chk_epi:
                with st.container(border=True):
                    st.markdown("**Hematoma Epidural**")
                    e1, e2, e3, e4, e5 = st.columns([1.2, 1.5, 1, 1, 1])
                    epi_lat = e1.selectbox("Lado", ["Derecho", "Izquierdo", "Bilateral"], key="epi_l")
                    epi_loc = e2.selectbox("Localización", loc_craneo, key="epi_loc")
                    ea_cm = e3.number_input("A (cm)", 0.0, 20.0, 0.0, key="ea")
                    eb_cm = e4.number_input("B (cm)", 0.0, 20.0, 0.0, key="eb")
                    ec_cm = e5.number_input("C (cm)", 0.0, 20.0, 0.0, key="ec")
                    vol_epi = calcular_volumen_hematoma(ea_cm, eb_cm, ec_cm)
                    st.caption(f"Volumen Epidural (ABC/2): **{vol_epi:.1f} cc**")
                    lesiones_detalles.append(f"Hematoma Epidural {epi_loc} {epi_lat} ({vol_epi:.1f} cc)")

            if not tc_normal and chk_sub:
                with st.container(border=True):
                    st.markdown("**Hematoma Subdural**")
                    s1, s2, s3, s4 = st.columns([1.2, 1.5, 1.2, 1.2])
                    sub_lat = s1.selectbox("Lado", ["Derecho", "Izquierdo", "Bilateral"], key="sub_l")
                    sub_loc = s2.selectbox("Localización", loc_subdural, key="sub_loc")
                    gros_sub = s3.number_input("Grosor máximo (mm)", 0.0, 50.0, 0.0, step=0.5, key="sub_g")
                    vol_sub = s4.number_input("Volumen aprox. (cc, opc.)", 0.0, 200.0, 0.0, key="sub_v")
                    lesiones_detalles.append(f"Hematoma Subdural {sub_loc} {sub_lat} (grosor {gros_sub:.1f} mm)")

            if not tc_normal and chk_ip:
                with st.container(border=True):
                    st.markdown("**Hematoma Intraparenquimatoso**")
                    p1, p2, p3, p4, p5 = st.columns([1.2, 1.5, 1, 1, 1])
                    ip_lat = p1.selectbox("Lado", ["Derecho", "Izquierdo", "Bilateral"], key="ip_l")
                    ip_loc = p2.selectbox("Localización", loc_parenquima, key="ip_loc")
                    pa_cm = p3.number_input("A (cm)", 0.0, 20.0, 0.0, key="pa")
                    pb_cm = p4.number_input("B (cm)", 0.0, 20.0, 0.0, key="pb")
                    pc_cm = p5.number_input("C (cm)", 0.0, 20.0, 0.0, key="pc")
                    vol_ip = calcular_volumen_hematoma(pa_cm, pb_cm, pc_cm)
                    st.caption(f"Volumen Intraparenquimatoso (ABC/2): **{vol_ip:.1f} cc**")
                    lesiones_detalles.append(f"Hematoma Intraparenquimatoso {ip_loc} {ip_lat} ({vol_ip:.1f} cc)")

            if not tc_normal and chk_cnt:
                with st.container(border=True):
                    st.markdown("**Contusión Hemorrágica**")
                    k1, k2, k3, k4, k5 = st.columns([1.2, 1.5, 1, 1, 1])
                    cnt_lat = k1.selectbox("Lado", ["Derecho", "Izquierdo", "Bilateral"], key="cnt_l")
                    cnt_loc = k2.selectbox("Localización", loc_parenquima, key="cnt_loc")
                    ka_cm = k3.number_input("A (cm)", 0.0, 20.0, 0.0, key="ka")
                    kb_cm = k4.number_input("B (cm)", 0.0, 20.0, 0.0, key="kb")
                    kc_cm = k5.number_input("C (cm)", 0.0, 20.0, 0.0, key="kc")
                    vol_cont = calcular_volumen_hematoma(ka_cm, kb_cm, kc_cm)
                    st.caption(f"Volumen Contusión (ABC/2): **{vol_cont:.1f} cc**")
                    lesiones_detalles.append(f"Contusión Hemorrágica {cnt_loc} {cnt_lat} ({vol_cont:.1f} cc)")

            hay_masa = (not tc_normal) and (chk_epi or chk_sub or chk_ip or chk_cnt)
            vol_max = max(vol_epi, vol_sub, vol_ip, vol_cont)
            tc_evacuada = st.radio("¿Masa ya evacuada quirúrgicamente? (Marshall V)", ["NO", "SI"], horizontal=True, disabled=tc_normal)

            st.markdown("#### 3. Hemorragia Subaracnoidea (HSA) e Intraventricular (HIV)")
            cs1, cs2, cs3 = st.columns([2, 1.5, 1.5])
            tc_hsa_estado = cs1.selectbox(
                "HSA Traumática (Escala de Greene)",
                ["Ausente", "Igual o menor a 5 mm (≤ 5 mm)", "Mayor a 5 mm (> 5 mm)"],
                disabled=tc_normal
            )
            efecto_masa_greene = cs2.checkbox(
                "¿Presenta efecto de masa?",
                value=(hay_masa and desviacion_lm_mm > 0),
                disabled=(tc_normal or tc_hsa_estado == "Ausente")
            )
            tc_hiv_estado = cs3.selectbox("Hemorragia Intraventricular (HIV)", ["Ausente", "Presente"], disabled=tc_normal)

            st.markdown("#### 4. Ventana Ósea y Neumoencéfalo")
            cf1, cf2, cf3 = st.columns(3)
            tc_fractura = cf1.selectbox("Fractura Craneal", ["Ausente", "Lineal", "Deprimida Cerrada", "Deprimida Abierta", "Conminuta", "Base de Cráneo"])
            tc_fx_ubic = cf2.text_input("Ubicación Fractura", placeholder="Ej: Temporal derecha") if tc_fractura != "Ausente" else ""
            tc_neumo = cf3.selectbox("Neumoencéfalo", ["Ausente", "Presente", "A Tensión"])

            if tc_normal:
                marshall_score, marshall_desc = calcular_marshall_exacto(True, "Normales", 0.0, 0.0, "NO")
                rotterdam_score = calcular_rotterdam_exacto("Normales", 0.0, False, "Ausente", "Ausente")
                greene_score, greene_desc = "0", "No aplica"
            else:
                marshall_score, marshall_desc = calcular_marshall_exacto(False, tc_cisternas, desviacion_lm_mm, vol_max, tc_evacuada)
                rotterdam_score = calcular_rotterdam_exacto(tc_cisternas, desviacion_lm_mm, chk_epi, tc_hsa_estado, tc_hiv_estado)
                greene_score, greene_desc = calcular_greene_exacto(tc_hsa_estado, efecto_masa_greene, desviacion_lm_mm)

            st.divider()
            r1, r2, r3 = st.columns(3)
            r1.metric("Escala de Marshall", f"Grado {marshall_score}")
            r1.caption(marshall_desc)
            r2.metric("Score de Rotterdam", f"{rotterdam_score} / 6")
            r3.metric("Escala de Greene (HSA)", f"Grado {greene_score}" if greene_score != "0" else "No aplica")
            r3.caption(greene_desc)

            cortes_str = ", ".join(tc_cortes) if tc_cortes else "axial"
            shift_txt = f"desviación de línea media de {desviacion_lm_mm} mm hacia la {lado_shift.lower()}" if desviacion_lm_mm > 0 else "línea media centrada"
            plantilla_tc = (
                f"Se evidencia TC de cráneo {tc_contraste.lower()} en formato {tc_formato.lower()}, cortes {cortes_str.lower()}. "
                f"Ventana parenquimatosa: adecuada relación cortico-subcortical, sistema ventricular normoconfigurado, "
                f"cisternas basales {tc_cisternas.lower()}, {shift_txt}, fosa posterior indemne."
            )
            if lesiones_detalles:
                plantilla_tc += f" Lesiones de masa: {'; '.join(lesiones_detalles)}."
            if tc_hsa_estado != "Ausente":
                plantilla_tc += f" HSA traumática ({greene_desc})."
            if tc_hiv_estado == "Presente":
                plantilla_tc += " Hemorragia intraventricular presente."
            if tc_fractura != "Ausente":
                plantilla_tc += f" Ventana ósea: fractura {tc_fractura.lower()} ({tc_fx_ubic})."
            if tc_neumo != "Ausente":
                plantilla_tc += f" Neumoencéfalo: {tc_neumo}."
            desc_tc = st.text_area("Informe Descriptivo de TC (Editable)", value=plantilla_tc, height=110)

        # 3. COLUMNA (RX / TC / RMN + PLANTILLA RX CERVICAL NORMAL + SÍNDROME COMPLETO/INCOMPLETO)
        with tab_columna:
            st.subheader("🦴 Columna y Médula")
            rx_segmentos = st.multiselect(
                "Estudios de Imagen Solicitados (RX / TC / RMN)",
                [
                    "RX Cervical (AP/Lat/Transoral)", "RX Dorsal (AP/Lat)", "RX Lumbar (AP/Lat)", "RX Sacra",
                    "TC Columna Cervical", "TC Columna Dorsal", "TC Columna Lumbar",
                    "RMN Columna Cervical", "RMN Columna Dorsal", "RMN Columna Lumbar"
                ],
                default=["RX Cervical (AP/Lat/Transoral)"]
            )
            plantilla_rx_normal = (
                "Radiografía de columna cervical en proyecciones AP, lateral y transoral (visualizándose adecuadamente hasta C7-T1): "
                "adecuada alineación de los cuerpos vertebrales y de las líneas espinales (vertebral anterior, vertebral posterior, "
                "espinolaminar y espinosa), altura de cuerpos vertebrales y espacios intervertebrales conservados, "
                "sin trazos de fractura ni listesis, articulación atlanto-axoidea conservada y espacio de partes blandas prevertebral dentro de límites normales."
            )
            valor_rx_inicial = plantilla_rx_normal if "RX Cervical (AP/Lat/Transoral)" in rx_segmentos else "Sin alteraciones óseas evidentes."
            rx_hallazgos = st.text_area("Informe Radiológico / Imagen de Columna (Editable)", value=valor_rx_inicial, height=110)

            st.divider()
            c_asia1, c_asia2, c_asia3 = st.columns(3)
            asia_score = c_asia1.selectbox("ASIA", ["E", "D", "C", "B", "A"])
            nivel_lesion = c_asia2.text_input("Nivel (NLI)", "N/A")
            sindrome_medular = c_asia3.selectbox("Síndrome Medular", ["Ninguno", "Síndrome Medular Completo", "Síndrome Medular Incompleto"])

            st.markdown("**TLICS (Toracolumbar)**")
            morf = st.selectbox("Morfología", ["Ninguna (0)", "Compresión (1)", "Estallido (2)", "Traslación (3)", "Distracción (4)"])
            plc = st.selectbox("Ligamento PLC", ["Intacto (0)", "Indeterminado (2)", "Roto (3)"])
            neuro_st = st.selectbox("Estado Neuro", ["Intacto (0)", "Raíz (2)", "Completa (2)", "Incompleta (3)", "Cauda (3)"])
            tlics_score = calcular_tlics(morf, plc, neuro_st)
            tlics_decision = "Conservador" if tlics_score <= 3 else ("QX" if tlics_score >= 5 else "Indeterminado")
            st.metric("TLICS", f"{tlics_score} ({tlics_decision})")

    # --- TAB 5: PROTOCOLO Y LABS ---
    with tabs[4]:
        st.subheader("👁️ Neuromonitoreo no invasivo")
        c_chk1, c_chk2 = st.columns(2)
        chk_dvno_eco = c_chk1.checkbox("Ecografía Ocular")
        chk_dvno_tc  = c_chk2.checkbox("Tomografía (TC)")

        dvno_partes = []
        dvno_metodos_sel = []
        dvno = 0.0

        if chk_dvno_eco:
            dvno_metodos_sel.append("Ecografía Ocular")
            with st.container(border=True):
                st.markdown("**DVNO por Ecografía Ocular**")
                ce1, ce2, ce3 = st.columns(3)
                eco_od = ce1.number_input("OD Eco (mm)", 0.0, 10.0, 4.5, step=0.1, key="eco_od")
                eco_oi = ce2.number_input("OI Eco (mm)", 0.0, 10.0, 4.5, step=0.1, key="eco_oi")
                prom_eco = (eco_od + eco_oi) / 2
                ce3.metric("Promedio Eco", f"{prom_eco:.2f} mm")
                if prom_eco > 5.2:
                    st.error("🚨 Eco DVNO > 5.2 mm (Signo indirecto de HIC)")
                else:
                    st.success("✅ Eco DVNO en rango normal")
                dvno_partes.append(f"Eco Ocular (OD {eco_od:.1f} / OI {eco_oi:.1f} mm | Prom {prom_eco:.2f} mm)")
                dvno = max(dvno, round(prom_eco, 2))

        if chk_dvno_tc:
            dvno_metodos_sel.append("Tomografía (TC)")
            with st.container(border=True):
                st.markdown("**DVNO por Tomografía (TC)**")
                ct1, ct2, ct3 = st.columns(3)
                tc_od = ct1.number_input("OD TC (mm)", 0.0, 10.0, 4.5, step=0.1, key="tc_od")
                tc_oi = ct2.number_input("OI TC (mm)", 0.0, 10.0, 4.5, step=0.1, key="tc_oi")
                prom_tc = (tc_od + tc_oi) / 2
                ct3.metric("Promedio TC", f"{prom_tc:.2f} mm")
                if prom_tc > 5.2:
                    st.error("🚨 TC DVNO > 5.2 mm (Signo indirecto de HIC)")
                else:
                    st.success("✅ TC DVNO en rango normal")
                dvno_partes.append(f"TC (OD {tc_od:.1f} / OI {tc_oi:.1f} mm | Prom {prom_tc:.2f} mm)")
                dvno = max(dvno, round(prom_tc, 2))

        dvno_metodo = " + ".join(dvno_metodos_sel) if dvno_metodos_sel else "No realizado"
        dvno_resumen = " | ".join(dvno_partes) if dvno_partes else "No realizado"

        with st.expander("➕ PIC Invasiva (Opcional)", expanded=False):
            pic = st.number_input("PIC Invasiva (mmHg)", 0, 100, 0)
            if pic > 0:
                ppc = calcular_ppc(pam, pic)
                st.metric("PPC", f"{ppc} mmHg")
                if ppc < 60:
                    st.error("⚠️ PPC < 60 mmHg (Meta BTF no cumplida)")
                if pic > 22:
                    st.warning("⚠️ PIC > 22 mmHg (Umbral BTF 4.ª ed.)")

        st.markdown("---")
        st.subheader("🧪 Laboratorios")
        cl1, cl2, cl3, cl4 = st.columns(4)
        lab_hb = cl1.number_input("Hb (g/dL)", 0.0, 25.0, 12.0)
        lab_hto = cl2.number_input("Hto (%)", 0.0, 70.0, 36.0)
        lab_wbc = cl3.number_input("Leucocitos", 0, 50000, 8000)
        lab_na = cl4.number_input("Sodio Na+ (mEq/L)", 100, 180, 140)
        cl5, cl6, cl7 = st.columns(3)
        lab_tp = cl5.number_input("TP (seg)", 0.0, 100.0, 12.0)
        lab_tpt = cl6.number_input("TPT (seg)", 0.0, 100.0, 30.0)
        lab_glu = cl7.number_input("Glucosa (mg/dL)", 0, 600, 90)

        with st.expander("➕ Otros Laboratorios (Opcional)", expanded=False):
            co1, co2, co3, co4 = st.columns(4)
            lab_plt = co1.number_input("Plaquetas", 0, 500000, 250000)
            lab_inr = co2.number_input("INR", 0.5, 10.0, 1.0)
            lab_urea = co3.number_input("Urea (mg/dL)", 0, 300, 30)
            lab_crea = co4.number_input("Creatinina (mg/dL)", 0.0, 20.0, 0.8)
            co5, co6, co7, co8 = st.columns(4)
            lab_k = co5.number_input("Potasio K+", 1.0, 10.0, 4.0)
            lab_cl = co6.number_input("Cloro Cl-", 60, 150, 100)
            lab_ca = co7.number_input("Calcio Ca++", 4.0, 15.0, 8.5)
            lab_vih = co8.selectbox("VIH", ["Pendiente", "No Reactivo", "Reactivo"])
            lab_vdrl = st.selectbox("VDRL", ["Pendiente", "No Reactivo", "Reactivo"])

        st.markdown("---")
        st.subheader("🛡️ Tiers de Manejo BTF")
        tier_manejo = st.selectbox("Escalón de Manejo (BTF)", ["Tier 0 (Observación)", "Tier 1 (Básico: Cabecera 30°, Sedación)", "Tier 2 (Hiperosmolar, RMN)", "Tier 3 (Barbitúricos, Hipotermia, Descompresiva)"])

    # --- TAB 6: PLAN, ORIENTACIÓN CLÍNICA, QX Y WORD ---
    with tabs[5]:
        conducta_cod, conducta_msg, criterios_qx = evaluar_conducta_clinica(
            marshall_score, vol_epi, gros_sub, vol_ip, vol_cont, desviacion_lm_mm, tlics_score
        )

        st.subheader("📋 Diagnóstico Jerárquico Automático")
        es_abierto = (tc_neumo != "Ausente" or tc_fractura == "Deprimida Abierta" or exposicion_dural_osea)
        tipo_tce_str = "Abierto" if es_abierto else "Cerrado"

        diag_lines = [
            "1. Politraumatismo",
            f"1.1 Traumatismo Craneoencefálico {clasif_tce} {tipo_tce_str} (GCS {gcs_display} | Marshall {marshall_score})"
        ]
        sub_i = 1
        for ld in lesiones_detalles:
            diag_lines.append(f"1.1.{sub_i} {ld}")
            sub_i += 1
        if not tc_normal and tc_hsa_estado != "Ausente":
            diag_lines.append(f"1.1.{sub_i} Hemorragia Subaracnoidea Postraumática (Greene {greene_score})")
            sub_i += 1
        if tc_fractura != "Ausente":
            diag_lines.append(f"1.1.{sub_i} Fractura Craneal {tc_fractura} ({tc_fx_ubic})")
            sub_i += 1
        if tc_neumo != "Ausente":
            diag_lines.append(f"1.1.{sub_i} Neumoencéfalo ({tc_neumo})")
            sub_i += 1
        if tlics_score > 0 or asia_score != "E":
            diag_lines.append(f"1.2 Trauma Raquimedular (ASIA {asia_score} | TLICS {tlics_score})")

        diagnostico_final = st.text_area("Diagnósticos (Editable)", value="\n".join(diag_lines), height=140)

        st.subheader("🚦 Orientación Clínica y Alertas (Marshall + BTF)")
        if conducta_cod == "QUIRURGICO":
            st.error(conducta_msg + ("\n\n" + "\n".join([f"• {c}" for c in criterios_qx]) if criterios_qx else ""))
        elif conducta_cod == "ANTIEDEMA":
            st.warning(conducta_msg)
        elif conducta_cod == "POSTQX":
            st.info(conducta_msg)
        else:
            st.success(conducta_msg)

        if total_gcs <= 8:
            st.error("🚨 BANDERA ROJA: GCS ≤ 8 (TCE Severo — Asegurar vía aérea definitiva / UCI)")
        if anisocoria == "SI" or r_d == "-" or r_i == "-":
            st.error("🚨 BANDERA ROJA: Alteración Pupilar (Anisocoria ≥ 2 mm o arreactividad)")

        plan_lines = [
            "1. Hospitalización / Cabecera a 30° en posición neutra.",
            "2. Hidratación isotónica (Solución 0.9% — evitar soluciones hipotónicas).",
            "3. Monitoreo neurológico estricto (GCS y pupilas horario).",
            "4. Neuroprotección (normotermia, normoglicemia, normonatremia)."
        ]
        if conducta_cod == "ANTIEDEMA":
            plan_lines.append("5. Terapia antiedema / hiperosmolar (Solución Salina Hipertónica o Manitol según protocolo BTF) + monitoreo DVNO seriado.")
        elif conducta_cod == "QUIRURGICO":
            plan_lines.append("5. Preparar para intervención neuroquirúrgica de urgencia: dieta absoluta, tipiaje, hemoderivados y consentimiento informado.")
        else:
            plan_lines.append("5. Analgesia, protección gástrica y vigilancia de signos de alarma.")

        plan = st.text_area("Órdenes Médicas Sugeridas (Editables)", value="\n".join(plan_lines), height=130)

        st.markdown("---")
        st.subheader("🔪 Registro Quirúrgico")
        qx_realizada = st.checkbox("¿Paciente llevado a Quirófano?")
        if qx_realizada:
            cirujano = st.text_input("Cirujano Principal", value=user_info["nombre"])
            tipo_qx = st.selectbox("Procedimiento", ["Hemicraniectomía Descompresiva", "Craneotomía Evacuadora", "Trepanación", "Cirugía Espinal", "DVE"])
            qx_tamano = st.selectbox("Tamaño Craniectomía (Calidad)", ["No aplica", "Adecuada (>12x15 cm)", "Pequeña (<12 cm)"])
            qx_dura = st.selectbox("Manejo Dural", ["No aplica", "Duroplastia Expansiva", "Cierre Primario Hermético", "Abierta"])
            hallazgos_qx = st.text_area("Hallazgos Quirúrgicos")

    # --- TAB 7: EGRESO Y GOSE ---
    with tabs[6]:
        st.subheader("🏁 Resultados al Alta (Outcomes)")
        egreso_check = st.checkbox("Egresar Paciente")
        condicion_egreso = "Hospitalizado"
        gose_desc = "No evaluado"
        complicaciones = []
        if egreso_check:
            condicion_egreso = st.radio("Condición al Egreso", ["VIVO", "FALLECIDO", "REFERIDO", "CONTRA OPINIÓN MÉDICA"], horizontal=True)
            gose_sel = st.selectbox("Glasgow Outcome Scale Extended (GOSE)", [
                "1. Fallecido", "2. Estado Vegetativo", "3. Discapacidad Severa Inferior",
                "4. Discapacidad Severa Superior", "5. Discapacidad Moderada Inferior",
                "6. Discapacidad Moderada Superior", "7. Buena Recuperación Inferior", "8. Buena Recuperación Superior"
            ])
            gose_desc = "1. Fallecido" if condicion_egreso == "FALLECIDO" else gose_sel
            st.markdown("**Complicaciones Intrahospitalarias**")
            complicaciones = st.multiselect("Marque si presentó:", ["Ninguna", "Neumonía Asociada a Ventilación (NAV)", "Meningitis/Ventriculitis", "Hidrocefalia Post-traumática", "Crisis Convulsivas", "Fístula de LCR"])

    # ==========================================
    # GENERADOR DE WORD
    # ==========================================
    def generar_reporte_word():
        try:
            doc = Document()
            section = doc.sections[0]
            section.top_margin = Inches(0.5); section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.6); section.right_margin = Inches(0.6)

            t = doc.add_table(1, 3); t.autofit = False
            t.columns[0].width = Inches(1); t.columns[1].width = Inches(5); t.columns[2].width = Inches(1)
            def p_img(c, i):
                if i:
                    try:
                        c.paragraphs[0].add_run().add_picture(i, width=Inches(0.9))
                    except Exception:
                        pass
            p_img(t.cell(0, 0), logo_h); p_img(t.cell(0, 2), logo_s)

            p = t.cell(0, 1).paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run("HOSPITAL CENTRAL UNIVERSITARIO\nDR. “ANTONIO MARÍA PINEDA”\n").bold = True
            p.add_run("SERVICIO DE NEUROCIRUGÍA").bold = True

            doc.add_paragraph()
            p = doc.add_paragraph("HISTORIA CLÍNICA NEUROCIRUGÍA (PROTOCOLO TRAUMA)")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.runs[0].bold = True

            doc.add_paragraph(f"PACIENTE: {nombre}. C.I: {cedula}. EDAD: {edad}. SEXO: {sexo}. OCUPACIÓN: {ocupacion}.")
            doc.add_paragraph(f"INGRESO: {fecha_ingreso} {hora_ingreso}. MÉDICO: {medico}. HOSPITAL: {hospital}.")
            doc.add_paragraph(f"PROCEDENCIA: {municipio}, {estado}. TLF: {telefono_full}.")
            doc.add_paragraph(f"MOTIVO: {motivo}.")
            doc.add_paragraph(f"EA: {ea}.")
            doc.add_paragraph(f"ANTECEDENTES: {', '.join(patologias_list) if patologias_list else 'Niega'}. Hábitos: Alc {var_alc}, Tab {var_tab}, Chimó {var_chimo}, Drogas {var_drogas}.")

            doc.add_paragraph("EXAMEN FÍSICO:").runs[0].bold = True
            doc.add_paragraph(f"VITALES: PA {tas}/{tad} (PAM {pam}). FC {fc}. FR {fr}. SatO2 {sat}%. Índice Choque: {shock_index:.2f}.")
            doc.add_paragraph(f"- GENERAL: {piel_desc}.")
            doc.add_paragraph(f"- SEGMENTARIO: {trauma_txt}. {tx_str}. {abd_str}. {ext_str}.")

            doc.add_paragraph("EXAMEN NEUROLÓGICO:").runs[0].bold = True
            doc.add_paragraph(f"1. MENTAL: GCS {gcs_display}. Lenguaje: {lenguaje}. Orientación: {orientacion}.")
            doc.add_paragraph("2. PARES CRANEALES:")
            doc.add_paragraph(f"   - I: {cn_1}. II: {cn_2_visual} / {cn_2_campimetria} / FO: {cn_2_fondo}.")
            doc.add_paragraph(f"   - III, IV, VI: {cn_346_motilidad}. Pupilas: {cn_346_pupilas}.")
            doc.add_paragraph(f"   - V: {cn_5_sens}. VII: {cn_7}. VIII: {cn_8}.")
            doc.add_paragraph(f"   - IX, X, XI, XII: {cn_bajos}.")
            doc.add_paragraph(f"3. MOTOR: Der {fd} | Izq {fi}. Tono: {tono_muscular}.")
            doc.add_paragraph(f"4. SENSITIVO: Sup {sens_sup} | Prof {sens_prof}. Nivel: {nivel_sensitivo}.")
            doc.add_paragraph(f"5. REFLEJOS: {reflejos_txt}. Babinski: {babinski}.")
            doc.add_paragraph(f"6. CEREBELO/MENÍNGEOS: {', '.join(cerebelo_txt) if cerebelo_txt else 'Sin alt'}. {', '.join(signos_meningeos)}.")

            if sindrome_medular != "Ninguno" or asia_score != "E":
                doc.add_paragraph(f"COLUMNA CLÍNICA: ASIA {asia_score}. Síndrome: {sindrome_medular}. TLICS: {tlics_score}.")

            doc.add_paragraph("IMAGENOLOGÍA Y ESCALAS:").runs[0].bold = True
            doc.add_paragraph(f"ESCALAS: Marshall {marshall_desc} | Rotterdam {rotterdam_score}/6 | Greene {greene_desc}.")
            if rx_segmentos:
                doc.add_paragraph(f"ESTUDIOS DE COLUMNA: {', '.join(rx_segmentos)}. Hallazgos: {rx_hallazgos}.")
            doc.add_paragraph(f"INFORME TC: {desc_tc}")

            doc.add_paragraph("PROTOCOLO Y PARACLÍNICOS:").runs[0].bold = True
            doc.add_paragraph(f"DVNO: {dvno_resumen}. PIC: {pic} mmHg. PPC: {ppc} mmHg.")
            doc.add_paragraph(f"Labs: Hb {lab_hb} | Hto {lab_hto}% | WBC {lab_wbc} | TP {lab_tp} | TPT {lab_tpt} | Na+ {lab_na} | Glu {lab_glu}.")
            doc.add_paragraph(f"TXA (Ácido Tranexámico): {txa_admin} (Tiempo: {txa_tiempo}). Manejo: {tier_manejo}.")

            doc.add_paragraph("DIAGNÓSTICOS:").runs[0].bold = True
            doc.add_paragraph(diagnostico_final)

            doc.add_paragraph("PLAN DE TRATAMIENTO:").runs[0].bold = True
            doc.add_paragraph(plan)

            if qx_realizada:
                doc.add_paragraph("REGISTRO QUIRÚRGICO:").runs[0].bold = True
                doc.add_paragraph(f"Cx: {cirujano}. Proc: {tipo_qx}. Tamaño: {qx_tamano}. Dura: {qx_dura}. Hallazgos: {hallazgos_qx}.")

            if egreso_check:
                doc.add_paragraph("EGRESO Y OUTCOMES:").runs[0].bold = True
                doc.add_paragraph(f"Condición: {condicion_egreso}. GOSE: {gose_desc}. Complicaciones: {', '.join(complicaciones) if complicaciones else 'Ninguna'}.")

            if uploaded_images:
                doc.add_page_break()
                doc.add_paragraph("ANEXO FOTOGRÁFICO").alignment = WD_ALIGN_PARAGRAPH.CENTER
                for img in uploaded_images:
                    try:
                        img.seek(0)
                        doc.add_picture(img, width=Inches(3.5))
                        doc.add_paragraph("")
                    except Exception:
                        pass

            bio = io.BytesIO(); doc.save(bio); return bio
        except Exception:
            return None

    # --- GUARDAR E IMPRIMIR AL FINAL DE TAB 6 (QX/PLAN) ---
    with tabs[5]:
        st.divider()
        cb1, cb2 = st.columns(2)
        if cb1.button("💾 GUARDAR EN BASE DE DATOS", use_container_width=True):
            nuevo = {
                "Fecha": fecha_ingreso, "Hora": hora_ingreso, "Hospital": hospital, "Medico": medico,
                "Cedula": cedula, "Nombre": nombre, "Edad": edad, "Sexo": sexo, "Ocupacion": ocupacion,
                "Estado": estado, "Municipio": municipio, "Mecanismo": mecanismo, "Horas_Evolucion": delta_t,
                "Alcohol": var_alc, "Tabaco": var_tab, "Chimo": var_chimo, "Drogas": var_drogas,
                "Acido_Tranexamico": txa_admin, "Ventana_TXA": txa_tiempo,
                "GCS_Ingreso": total_gcs, "GCS_Texto": gcs_display, "Severidad_TCE": clasif_tce, "Tipo_TCE": tipo_tce_str,
                "Anisocoria": anisocoria, "Cisternas": tc_cisternas, "Desviacion_LM_mm": desviacion_lm_mm,
                "Marshall": marshall_score, "Rotterdam": rotterdam_score, "Greene": greene_score,
                "Conducta_Sugerida": conducta_cod, "Lesiones_Masa": "; ".join(lesiones_detalles) if lesiones_detalles else "Ninguna",
                "Fractura": tc_fractura, "Neumoencefalo": tc_neumo,
                "ASIA": asia_score, "Sindrome_Medular": sindrome_medular, "TLICS": tlics_score,
                "DVNO_Metodo": dvno_metodo, "DVNO_mm": dvno, "PIC": pic, "PPC": ppc,
                "Hb": lab_hb, "Hto": lab_hto, "Leucocitos": lab_wbc, "TP": lab_tp, "TPT": lab_tpt, "Na": lab_na, "Glucosa": lab_glu,
                "Quirurgico_Si_No": "SI" if qx_realizada else "NO", "Tipo_QX": tipo_qx if qx_realizada else "N/A",
                "Condicion_Egreso": condicion_egreso, "GOSE": gose_desc
            }
            df_new = pd.DataFrame([nuevo])
            f_db = "neuro_db.csv"
            df_new.to_csv(f_db, mode="a" if os.path.exists(f_db) else "w", header=not os.path.exists(f_db), index=False)
            st.success("✅ Registro guardado en la Base Madre.")

        wf = generar_reporte_word()
        if wf:
            cb2.download_button(
                "🖨️ IMPRIMIR / DESCARGAR HISTORIA WORD",
                wf.getvalue(),
                f"Historia_{nombre or 'Paciente'}.docx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

    # --- TAB 8: DASHBOARD ---
    with tabs[7]:
        st.subheader("📊 Dashboard Epidemiológico y de Investigación")
        if os.path.exists("neuro_db.csv"):
            df = pd.read_csv("neuro_db.csv")
            if not df.empty:
                total_p = len(df)
                fallecidos = len(df[df["Condicion_Egreso"] == "FALLECIDO"]) if "Condicion_Egreso" in df.columns else 0
                qx_cnt = len(df[df["Quirurgico_Si_No"] == "SI"]) if "Quirurgico_Si_No" in df.columns else 0
                txa_cnt = len(df[df["Acido_Tranexamico"] == "SI"]) if "Acido_Tranexamico" in df.columns else 0

                k1, k2, k3, k4 = st.columns(4)
                k1.metric("Total Pacientes", total_p)
                k2.metric("Fallecidos (Mortalidad)", f"{fallecidos} ({(fallecidos/total_p*100):.1f}%)")
                k3.metric("Casos Quirúrgicos", f"{qx_cnt} ({(qx_cnt/total_p*100):.1f}%)")
                k4.metric("Ácido Tranexámico", f"{txa_cnt} ({(txa_cnt/total_p*100):.1f}%)")

                st.divider()
                if "Severidad_TCE" in df.columns:
                    s1, s2, s3 = st.columns(3)
                    n_leve = len(df[df["Severidad_TCE"] == "Leve"])
                    n_mod = len(df[df["Severidad_TCE"] == "Moderado"])
                    n_sev = len(df[df["Severidad_TCE"] == "Severo"])
                    s1.metric("TCE Leve", f"{n_leve} ({(n_leve/total_p*100):.1f}%)")
                    s2.metric("TCE Moderado", f"{n_mod} ({(n_mod/total_p*100):.1f}%)")
                    s3.metric("TCE Severo", f"{n_sev} ({(n_sev/total_p*100):.1f}%)")

                st.divider()
                g1, g2 = st.columns(2)
                if "Mecanismo" in df.columns:
                    g1.markdown("**Distribución por Mecanismo**")
                    g1.bar_chart(df["Mecanismo"].value_counts())
                if "Severidad_TCE" in df.columns:
                    g2.markdown("**Distribución por Severidad de TCE**")
                    g2.bar_chart(df["Severidad_TCE"].value_counts())
            else:
                st.info("Sin datos registrados aún.")
        else:
            st.info("Aún no hay pacientes guardados en la base de datos.")

if __name__ == "__main__":
    main()