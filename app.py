import streamlit as st
import pandas as pd
from datetime import datetime
import unicodedata
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Digitación Encuesta 7C", layout="wide")

# =====================================================
# ESTILOS
# =====================================================

st.markdown("""
<style>
    .stApp { background-color: #f5f7fa; }
    .block-container { padding-top: 1.5rem; }
    div[data-testid="stMetric"] {
        background: white;
        border-radius: 10px;
        padding: 12px 16px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }
    .pregunta-label {
        font-size: 0.78rem;
        color: #555;
        margin-bottom: 2px;
    }
    .badge-respondidas {
        background: #d1fae5;
        color: #065f46;
        border-radius: 6px;
        padding: 2px 10px;
        font-size: 0.85rem;
        font-weight: 600;
    }
    .badge-vacias {
        background: #fee2e2;
        color: #991b1b;
        border-radius: 6px;
        padding: 2px 10px;
        font-size: 0.85rem;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

st.title("📋 Digitación Encuesta 7C")

# =====================================================
# CONFIGURACIÓN
# =====================================================

SHEET_ID = "1BvuZXujUBX3Ri1Q7lAtxHTWq1FSPGhlvZyhhffNuADw"
WORKSHEET_NAME = "Hoja 1"
CATALOGO_PATH = "catalogos_7c.xlsx"

REGIONES = [
    "REGION DE TARAPACA",
    "REGION DE ANTOFAGASTA",
    "REGION DE VALPARAISO",
    "REGION METROPOLITANA",
    "REGION DEL LIB. GRAL. BERNARDO O'HIGGINS",
    "REGION DEL MAULE",
    "REGION DEL BIOBIO",
    "REGION DE LA ARAUCANIA",
    "REGION DE LOS RIOS",
    "REGION DE LOS LAGOS",
    "REGION DE AYSEN",
    "REGION DE MAGALLANES"
]

NIVELES = ["", "1° Básico", "2° Básico", "3° Básico", "4° Básico", "5° Básico", "6° Básico"]
LETRAS  = ["", "A", "B", "C", "D", "E", "F", "G", "H"]

PREGUNTAS = [
    "1. Mi profesor(a) es amable conmigo cuando hago preguntas",
    "2. Me gusta la forma en que me trata mi profesor(a) cuando necesito ayuda",
    "3. Mi profesor(a) me pone atención cuando le hablo",
    "4. Me gustan las cosas que estamos aprendiendo en esta clase",
    "5. En esta clase, aprendemos a corregir nuestros errores",
    "6. Cuando nos enseña algo, mi profesor(a) nos pregunta si le entendemos",
    "7. Mi profesor(a) explica muy bien las cosas",
    "8. Mi profesor(a) dedica tiempo a ayudarnos a recordar lo que aprendemos",
    "9. Para ayudarnos a recordar, mi profesor(a) habla de cosas que ya hemos aprendido",
    "10. En esta clase aprendemos mucho casi todos los días",
    "11. Mi profesor(a) se asegura de que me esfuerzo al máximo",
    "12. Cuando algo es difícil para mí, mi profesor(a) igual me hace intentarlo",
    "13. Mi profesor(a) me pide que explique mis respuestas",
    "14. En esta clase, en general, nos mantenemos ocupados y no se pierde el tiempo",
    "15. Mis compañeros(as) actúan como mi profesor(a) espera",
]

COLUMNAS = [
    "Submitted Date", "Completion Time", "Completion Status",
    "Selecciona tu región", "Selecciona tu colegio",
    "Selecciona el nombre de tu profesor o profesora",
    "Selecciona el nivel de tu curso", "Selecciona la letra de tu curso",
    "E.1. ¿Cuál?",
    "1. Mi profesor(a) es amable conmigo cuando hago preguntas",
    "2. Me gusta la forma en que me trata mi profesor(a) cuando necesito ayuda",
    "3. Mi profesor(a) me pone atención cuando le hablo",
    "4. Me gustan las cosas que estamos aprendiendo en esta clase",
    "5. En esta clase, aprendemos a corregir nuestros errores",
    "6. Cuando nos enseña algo, mi profesor(a) nos pregunta si le entendemos",
    "7. Mi profesor(a) explica muy bien las cosas",
    "8. Mi profesor(a) dedica tiempo a ayudarnos a recordar lo que aprendemos",
    "9. Para ayudarnos a recordar, mi profesor(a) habla de cosas que ya hemos aprendido",
    "10. En esta clase aprendemos mucho casi todos los días",
    "11. Mi profesor(a) se asegura de que me esfuerzo al máximo",
    "12. Cuando algo es difícil para mí, mi profesor(a) igual me hace intentarlo",
    "13. Mi profesor(a) me pide que explique mis respuestas",
    "14. En esta clase, en general, nos mantenemos ocupados y no se pierde el tiempo",
    "15. Mis compañeros(as) actúan como mi profesor(a) espera",
    "16. Escribe cualquier comentario, sugerencia o agradecimiento que permita a tu profesor(a) hacer mejores clases para ti.",
    "17. ¿Cómo identificas tu género?",
    "18. ¿Cúantos libros hay en tu hogar?",
    "18. ¿Hay computador o tablet en tu casa? Si es así, ¿Cuántos?",
    "Dado que la encuesta es anónima y para asegurarnos de que no se generen respuestas duplicadas, piensa y escribe un número del 1 al 100.",
    "Response Url", "Referrer", "Ip Address", "Unprotected File List",
    # Columna extra para auditoría
    "Digitador"
]

OPCIONES_RESPUESTA = {
    0: "⬜ Sin respuesta",
    1: "❌ No",
    2: "🤔 Quizás",
    3: "✅ Sí",
}

MAPA_TEXTO = {0: "", 1: "No", 2: "Quizás", 3: "Sí"}

MIN_PREGUNTAS_RESPONDIDAS = 10  # Mínimo requerido para guardar

# =====================================================
# FUNCIONES
# =====================================================

def normalizar_texto(txt):
    if pd.isna(txt):
        return ""
    txt = str(txt).strip().replace("´", "'").replace("'", "'")
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return txt.upper()


@st.cache_data(ttl=3600)
def cargar_catalogos():
    colegios   = pd.read_excel(CATALOGO_PATH, sheet_name="Colegios")
    profesores = pd.read_excel(CATALOGO_PATH, sheet_name="Profesores")
    colegios.columns   = colegios.columns.str.strip()
    profesores.columns = profesores.columns.str.strip()
    colegios["region_key"]    = colegios["Region"].apply(normalizar_texto)
    profesores["colegio_key"] = profesores["Colegio"].apply(normalizar_texto)
    return colegios, profesores


@st.cache_resource
def conectar_google_sheets():
    """Conexión reutilizable a Google Sheets (no se recrea en cada guardado)."""
    if "gcp_service_account" not in st.secrets:
        st.error("❌ No se encontró la configuración 'gcp_service_account' en los secrets.")
        st.stop()
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return gspread.authorize(creds)


def obtener_indice(opciones, valor):
    try:
        return opciones.index(valor)
    except ValueError:
        return 0


def limpiar_respuestas():
    for i in range(15):
        st.session_state[f"p{i+1}"] = 0
    st.session_state["comentario"] = ""
    st.session_state["genero"]     = "Sin respuesta"
    st.session_state["libros"]     = "Sin respuesta"
    st.session_state["computador"] = "Sin respuesta"


# =====================================================
# CARGA DE CATÁLOGOS
# =====================================================

try:
    colegios_df, profesores_df = cargar_catalogos()
except Exception as e:
    st.error("❌ Error cargando catalogos_7c.xlsx. Verifica que el archivo esté en la carpeta correcta.")
    st.exception(e)
    st.stop()

# =====================================================
# SESSION STATE
# =====================================================

if "contexto" not in st.session_state:
    st.session_state.contexto = {}

if "contador" not in st.session_state:
    st.session_state.contador = 0

if "historial" not in st.session_state:
    st.session_state.historial = []  # Para mostrar últimas encuestas guardadas

# =====================================================
# DATOS DEL CURSO
# =====================================================

with st.expander("📌 Datos del curso (se mantienen entre encuestas)", expanded=True):
    col1, col2 = st.columns(2)

    with col1:
        opciones_region = [""] + REGIONES
        region = st.selectbox(
            "Región *",
            opciones_region,
            index=obtener_indice(opciones_region, st.session_state.contexto.get("region", ""))
        )
        region_key = normalizar_texto(region)

        colegios_filtrados = (
            colegios_df.loc[colegios_df["region_key"] == region_key, "Colegio"]
            .dropna().drop_duplicates().sort_values().tolist()
        )
        opciones_colegio = [""] + colegios_filtrados
        colegio = st.selectbox(
            "Colegio *",
            opciones_colegio,
            index=obtener_indice(opciones_colegio, st.session_state.contexto.get("colegio", ""))
        )
        colegio_key = normalizar_texto(colegio)

        profesores_filtrados = (
            profesores_df.loc[profesores_df["colegio_key"] == colegio_key, "Profesor"]
            .dropna().drop_duplicates().sort_values().tolist()
        )
        opciones_profesor = [""] + profesores_filtrados
        profesor = st.selectbox(
            "Profesor/a *",
            opciones_profesor,
            index=obtener_indice(opciones_profesor, st.session_state.contexto.get("profesor", ""))
        )

    with col2:
        nivel = st.selectbox(
            "Nivel *",
            NIVELES,
            index=obtener_indice(NIVELES, st.session_state.contexto.get("nivel", ""))
        )
        letra = st.selectbox(
            "Letra curso *",
            LETRAS,
            index=obtener_indice(LETRAS, st.session_state.contexto.get("letra", ""))
        )
        digitador = st.text_input(
            "Nombre digitador *",
            value=st.session_state.contexto.get("digitador", ""),
            placeholder="Tu nombre completo"
        )

# =====================================================
# RESPUESTAS P1 A P15
# =====================================================

st.subheader("Respuestas P1 a P15")
st.caption("Selecciona la opción que marcó el estudiante en el papel. **Sin respuesta** si la pregunta está en blanco.")

respuestas = []
cols = st.columns(5)

for i in range(15):
    key = f"p{i+1}"
    if key not in st.session_state:
        st.session_state[key] = 0

    with cols[i % 5]:
        val = st.selectbox(
            f"P{i+1}",
            options=list(OPCIONES_RESPUESTA.keys()),
            format_func=lambda x: OPCIONES_RESPUESTA[x],
            key=key
        )
        respuestas.append(val)

# =====================================================
# CONTADOR EN TIEMPO REAL
# =====================================================

respondidas   = sum(1 for r in respuestas if r != 0)
sin_responder = 15 - respondidas

col_a, col_b, col_c = st.columns(3)
col_a.metric("✅ Preguntas respondidas", respondidas)
col_b.metric("⬜ Sin respuesta", sin_responder)
col_c.metric("💾 Guardadas esta sesión", st.session_state.contador)

if respondidas < MIN_PREGUNTAS_RESPONDIDAS and respondidas > 0:
    st.warning(f"⚠️ Esta encuesta tiene solo {respondidas} respuestas. Se requieren al menos {MIN_PREGUNTAS_RESPONDIDAS} para guardar.")

# =====================================================
# INFORMACIÓN ADICIONAL
# =====================================================

with st.expander("ℹ️ Información adicional (opcional)"):
    comentario = st.text_area(
        "16. Comentario, sugerencia o agradecimiento",
        key="comentario",
        placeholder="Transcribe el comentario escrito por el estudiante, si hay alguno."
    )

    col3, col4, col5 = st.columns(3)

    with col3:
        genero = st.selectbox(
            "17. Género",
            ["Sin respuesta", "Masculino", "Femenino", "Otro género",
             "No me identifico con ningún género", "Prefiero no contestar"],
            key="genero"
        )
    with col4:
        libros = st.selectbox(
            "18. Libros en el hogar",
            ["Sin respuesta", "De 0 a 5", "De 6 a 10", "De 11 a 25", "De 26 a 50", "De 51 a 100"],
            key="libros"
        )
    with col5:
        computador = st.selectbox(
            "19. Computador o tablet en casa",
            ["Sin respuesta", "No", "Sí, 1", "Sí, 2", "Sí, 3 o más"],
            key="computador"
        )

# =====================================================
# GUARDAR
# =====================================================

st.markdown("---")
col_btn1, col_btn2 = st.columns([2, 1])

with col_btn1:
    guardar = st.button("💾 Guardar respuesta", type="primary", use_container_width=True)

with col_btn2:
    limpiar = st.button("🗑️ Limpiar para nueva encuesta", use_container_width=True)

if guardar:
    errores = []
    if not region:     errores.append("Región")
    if not colegio:    errores.append("Colegio")
    if not profesor:   errores.append("Profesor/a")
    if not nivel:      errores.append("Nivel")
    if not letra:      errores.append("Letra curso")
    if not digitador:  errores.append("Digitador")
    if respondidas == 0:
        errores.append("Al menos una respuesta P1-P15")
    elif respondidas < MIN_PREGUNTAS_RESPONDIDAS:
        errores.append(f"Mínimo {MIN_PREGUNTAS_RESPONDIDAS} preguntas respondidas (hay {respondidas})")

    if errores:
        st.error("⚠️ Faltan campos obligatorios: " + ", ".join(errores))
    else:
        # Guardar contexto para próxima encuesta
        st.session_state.contexto = {
            "region": region, "colegio": colegio, "profesor": profesor,
            "nivel": nivel, "letra": letra, "digitador": digitador
        }

        respuestas_texto = [MAPA_TEXTO[r] for r in respuestas]
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        data = {
            "Submitted Date": timestamp,
            "Completion Time": "",
            "Completion Status": "Complete",
            "Selecciona tu región": region,
            "Selecciona tu colegio": colegio,
            "Selecciona el nombre de tu profesor o profesora": profesor,
            "Selecciona el nivel de tu curso": nivel,
            "Selecciona la letra de tu curso": letra,
            "E.1. ¿Cuál?": "",
            "1. Mi profesor(a) es amable conmigo cuando hago preguntas": respuestas_texto[0],
            "2. Me gusta la forma en que me trata mi profesor(a) cuando necesito ayuda": respuestas_texto[1],
            "3. Mi profesor(a) me pone atención cuando le hablo": respuestas_texto[2],
            "4. Me gustan las cosas que estamos aprendiendo en esta clase": respuestas_texto[3],
            "5. En esta clase, aprendemos a corregir nuestros errores": respuestas_texto[4],
            "6. Cuando nos enseña algo, mi profesor(a) nos pregunta si le entendemos": respuestas_texto[5],
            "7. Mi profesor(a) explica muy bien las cosas": respuestas_texto[6],
            "8. Mi profesor(a) dedica tiempo a ayudarnos a recordar lo que aprendemos": respuestas_texto[7],
            "9. Para ayudarnos a recordar, mi profesor(a) habla de cosas que ya hemos aprendido": respuestas_texto[8],
            "10. En esta clase aprendemos mucho casi todos los días": respuestas_texto[9],
            "11. Mi profesor(a) se asegura de que me esfuerzo al máximo": respuestas_texto[10],
            "12. Cuando algo es difícil para mí, mi profesor(a) igual me hace intentarlo": respuestas_texto[11],
            "13. Mi profesor(a) me pide que explique mis respuestas": respuestas_texto[12],
            "14. En esta clase, en general, nos mantenemos ocupados y no se pierde el tiempo": respuestas_texto[13],
            "15. Mis compañeros(as) actúan como mi profesor(a) espera": respuestas_texto[14],
            "16. Escribe cualquier comentario, sugerencia o agradecimiento que permita a tu profesor(a) hacer mejores clases para ti.": comentario,
            "17. ¿Cómo identificas tu género?": "" if genero == "Sin respuesta" else genero,
            "18. ¿Cúantos libros hay en tu hogar?": "" if libros == "Sin respuesta" else libros,
            "18. ¿Hay computador o tablet en tu casa? Si es así, ¿Cuántos?": "" if computador == "Sin respuesta" else computador,
            "Dado que la encuesta es anónima y para asegurarnos de que no se generen respuestas duplicadas, piensa y escribe un número del 1 al 100.": "",
            "Response Url": "", "Referrer": "", "Ip Address": "", "Unprotected File List": "",
            "Digitador": digitador  # ← Auditoría
        }

        fila = [data.get(col, "") for col in COLUMNAS]

        try:
            client  = conectar_google_sheets()
            sheet   = client.open_by_key(SHEET_ID)
            worksheet = sheet.worksheet(WORKSHEET_NAME)
            worksheet.append_row(fila, value_input_option="USER_ENTERED")

            st.session_state.contador += 1

            # Guardar en historial local de sesión
            st.session_state.historial.append({
                "N°": st.session_state.contador,
                "Hora": timestamp,
                "Profesor": profesor,
                "Curso": f"{nivel} {letra}",
                "Respondidas": respondidas,
                "Digitador": digitador
            })

            st.success(f"✅ Encuesta #{st.session_state.contador} guardada correctamente.")
            st.info("Los datos del curso se mantienen. Limpia las respuestas para ingresar la siguiente encuesta.")

        except Exception as e:
            st.error("❌ No se pudo guardar en Google Sheets.")
            st.exception(e)

if limpiar:
    limpiar_respuestas()
    st.rerun()

# =====================================================
# HISTORIAL DE SESIÓN
# =====================================================

if st.session_state.historial:
    st.markdown("---")
    with st.expander(f"📊 Historial de esta sesión ({len(st.session_state.historial)} encuestas guardadas)"):
        df_historial = pd.DataFrame(st.session_state.historial)
        st.dataframe(df_historial, use_container_width=True, hide_index=True)