import streamlit as st
import pandas as pd
from datetime import datetime
import unicodedata

st.set_page_config(page_title="Digitación Encuesta 7C", layout="wide")

st.title("Digitación Encuesta 7C")

ARCHIVO_SALIDA = "respuestas_7c_estructura_google_sheets.xlsx"
CATALOGO_PATH = "catalogos_7c.xlsx"

REGIONES = [
    "REGION DE TARAPACA",
    "REGION DE ANTOFAGASTA",
    "REGION DE VALPARAISO",
    "REGION METROPOLITANA",
    "REGIÓN DEL LIB. GRAL. BERNARDO O'HIGGINS",
    "REGION DEL MAULE",
    "REGION DEL BIOBIO",
    "REGION DE LA ARAUCANIA",
    "REGION DE LOS RIOS",
    "REGION DE LOS LAGOS",
    "REGION DE AYSEN",
    "REGION DE MAGALLANES"
]

NIVELES = [
    "",
    "1ro básico",
    "2do básico",
    "3ro básico",
    "4to básico",
    "5to básico",
    "6to básico"
]

LETRAS = [
    "",
    "A", "B", "C", "D", "E", "F", "G",
    "H", "I", "J", "K", "L", "M", "N"
]

COLUMNAS = [
    "Submitted Date",
    "Completion Time",
    "Completion Status",
    "Selecciona tu región",
    "Selecciona tu colegio",
    "Selecciona el nombre de tu profesor o profesora",
    "Selecciona el nivel de tu curso",
    "Selecciona la letra de tu curso",
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
    "Response Url",
    "Referrer",
    "Ip Address",
    "Unprotected File List"
]

def normalizar_texto(txt):
    if pd.isna(txt):
        return ""
    txt = str(txt).strip().replace("´", "'").replace("’", "'")
    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    return txt.upper()

@st.cache_data
def cargar_catalogos():
    colegios = pd.read_excel(CATALOGO_PATH, sheet_name="Colegios")
    profesores = pd.read_excel(CATALOGO_PATH, sheet_name="Profesores")

    colegios.columns = colegios.columns.str.strip()
    profesores.columns = profesores.columns.str.strip()

    colegios["region_key"] = colegios["Region"].apply(normalizar_texto)
    colegios["colegio_key"] = colegios["Colegio"].apply(normalizar_texto)
    profesores["colegio_key"] = profesores["Colegio"].apply(normalizar_texto)

    return colegios, profesores

def obtener_indice(opciones, valor):
    try:
        return opciones.index(valor)
    except ValueError:
        return 0

if "contexto" not in st.session_state:
    st.session_state.contexto = {}

if "contador" not in st.session_state:
    st.session_state.contador = 0

try:
    colegios_df, profesores_df = cargar_catalogos()
except FileNotFoundError:
    st.error("No se encontró catalogos_7c.xlsx. Debe estar en la misma carpeta que app.py.")
    st.stop()
except Exception as e:
    st.error("Error al cargar catalogos_7c.xlsx.")
    st.exception(e)
    st.stop()

with st.expander("Datos del curso", expanded=True):
    col1, col2 = st.columns(2)

    with col1:
        opciones_region = [""] + REGIONES

        region = st.selectbox(
            "Región",
            opciones_region,
            index=obtener_indice(
                opciones_region,
                st.session_state.contexto.get("region", "")
            )
        )

        region_key = normalizar_texto(region)

        colegios_filtrados = (
            colegios_df.loc[
                colegios_df["region_key"] == region_key,
                "Colegio"
            ]
            .dropna()
            .drop_duplicates()
            .sort_values()
            .tolist()
        )

        opciones_colegio = [""] + colegios_filtrados

        colegio = st.selectbox(
            "Colegio",
            opciones_colegio,
            index=obtener_indice(
                opciones_colegio,
                st.session_state.contexto.get("colegio", "")
            )
        )

        colegio_key = normalizar_texto(colegio)

        profesores_filtrados = (
            profesores_df.loc[
                profesores_df["colegio_key"] == colegio_key,
                "Profesor"
            ]
            .dropna()
            .drop_duplicates()
            .sort_values()
            .tolist()
        )

        opciones_profesor = [""] + profesores_filtrados

        profesor = st.selectbox(
            "Profesor/a",
            opciones_profesor,
            index=obtener_indice(
                opciones_profesor,
                st.session_state.contexto.get("profesor", "")
            )
        )

    with col2:
        nivel = st.selectbox(
            "Nivel",
            NIVELES,
            index=obtener_indice(
                NIVELES,
                st.session_state.contexto.get("nivel", "")
            )
        )

        letra = st.selectbox(
            "Letra curso",
            LETRAS,
            index=obtener_indice(
                LETRAS,
                st.session_state.contexto.get("letra", "")
            )
        )

        digitador = st.text_input(
            "Nombre digitador",
            value=st.session_state.contexto.get("digitador", "")
        )

st.subheader("Respuestas P1 a P15")
st.info("0 = Sin respuesta | 1 = No | 2 = Quizás | 3 = Sí")

respuestas = []
cols = st.columns(5)

for i in range(15):
    key = f"p{i+1}"

    if key not in st.session_state:
        st.session_state[key] = 0

    with cols[i % 5]:
        val = st.number_input(
            f"P{i+1}",
            min_value=0,
            max_value=3,
            step=1,
            key=key
        )
        respuestas.append(val)

st.subheader("Información adicional")

comentario = st.text_area(
    "16. Comentario, sugerencia o agradecimiento",
    key="comentario"
)

col3, col4, col5 = st.columns(3)

with col3:
    genero = st.selectbox(
        "17. Género",
        [
            "Sin respuesta",
            "Masculino",
            "Femenino",
            "Otro género",
            "No me identifico con ningún género",
            "Prefiero no contestar"
        ],
        key="genero"
    )

with col4:
    libros = st.selectbox(
        "18. Libros en el hogar",
        [
            "Sin respuesta",
            "De 0 a 5",
            "De 6 a 10",
            "De 11 a 25",
            "De 26 a 50",
            "De 51 a 100"
        ],
        key="libros"
    )

with col5:
    computador = st.selectbox(
        "19. Computador o tablet en casa",
        [
            "Sin respuesta",
            "No",
            "Sí, 1",
            "Sí, 2",
            "Sí, 3 o más"
        ],
        key="computador"
    )

mapa_respuestas = {
    0: "",
    1: "No",
    2: "Quizás",
    3: "Sí"
}

respuestas_texto = [mapa_respuestas[r] for r in respuestas]

respondidas = sum(1 for r in respuestas if r != 0)
sin_responder = 15 - respondidas

st.markdown("---")

col_a, col_b, col_c = st.columns(3)

col_a.metric("Preguntas respondidas", respondidas)
col_b.metric("Preguntas sin respuesta", sin_responder)
col_c.metric("Encuestas guardadas en esta sesión", st.session_state.contador)

if st.button("Guardar respuesta", type="primary"):

    errores = []

    if not region:
        errores.append("Región")
    if not colegio:
        errores.append("Colegio")
    if not profesor:
        errores.append("Profesor/a")
    if not nivel:
        errores.append("Nivel")
    if not letra:
        errores.append("Letra curso")
    if not digitador:
        errores.append("Digitador")
    if respondidas == 0:
        errores.append("Al menos una respuesta P1 a P15")

    if errores:
        st.error("Faltan campos obligatorios: " + ", ".join(errores))

    else:
        st.session_state.contexto = {
            "region": region,
            "colegio": colegio,
            "profesor": profesor,
            "nivel": nivel,
            "letra": letra,
            "digitador": digitador
        }

        data = {
            "Submitted Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
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
            "Response Url": "",
            "Referrer": "",
            "Ip Address": "",
            "Unprotected File List": ""
        }

        nueva_fila = pd.DataFrame([data], columns=COLUMNAS)

        try:
            existente = pd.read_excel(ARCHIVO_SALIDA)
            final = pd.concat([existente, nueva_fila], ignore_index=True)
        except FileNotFoundError:
            final = nueva_fila

        final.to_excel(ARCHIVO_SALIDA, index=False)

        st.session_state.contador += 1

        st.success("Respuesta guardada correctamente en archivo Excel local.")
        st.info(f"Archivo generado/actualizado: {ARCHIVO_SALIDA}")
        st.info("Los datos del curso se mantienen. Puedes ingresar la siguiente encuesta.")

st.markdown("---")

if st.button("Limpiar respuestas para nueva encuesta"):
    for i in range(15):
        st.session_state[f"p{i+1}"] = 0

    st.session_state["comentario"] = ""
    st.session_state["genero"] = "Sin respuesta"
    st.session_state["libros"] = "Sin respuesta"
    st.session_state["computador"] = "Sin respuesta"

    st.rerun()

st.subheader("Vista previa de respuestas guardadas")

try:
    df_preview = pd.read_excel(ARCHIVO_SALIDA)
    st.dataframe(df_preview.tail(10), use_container_width=True)
except FileNotFoundError:
    st.caption("Aún no hay respuestas guardadas.")