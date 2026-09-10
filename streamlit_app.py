"""
App de Registro de Parámetros de Calidad - Insumos de Cocina Dulce (B2B Starbucks)
=====================================================================================
Streamlit + Google Drive/Sheets (histórico mensual, una hoja por día+turno).

Estructura esperada del repo:
  streamlit_app_cocina.py
  requirements.txt
  data/MA-PL-019_PLAN_CALIDAD_COCINA.xlsx

Secrets necesarios en Streamlit Cloud (Manage app -> Settings -> Secrets):

[gcp_service_account]
... (mismo bloque que en las apps de PT y PI) ...

ROOT_FOLDER_ID = "1h46IR2nUS8TZUSIgl1lD-M4nCtp-p_wW"
"""

import io
import re
from datetime import date, timedelta

import pandas as pd
import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build as build_drive_service
    GSHEETS_DISPONIBLE = True
except ImportError:
    GSHEETS_DISPONIBLE = False

# ----------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Registro de Calidad - Insumos Cocina Dulce B2B Starbucks",
    page_icon="🍯",
    layout="wide",
)

EXCEL_PATH = "MA-PL-019_PLAN_CALIDAD_COCINA.xlsx"
SHEET_NAME = "COCINA"
CLIENTE_FIJO = "STARBUCKS"
AREA_FIJA = "COCINA"

SUBCARPETA_DRIVE = "Insumos Cocina"
PLANTILLA_NOMBRE = "Insumos cocina - Plantilla base"
TEMPLATE_SHEET_NAME = "Hoja 1"
FOOTER_MARCA = "V°B° Jefe de Calidad"
FILA_ENCABEZADO_PLANTILLA = 4  # fila con los títulos de columna

MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

EQUIPO_CALIDAD = [
    "Verónica Iriarte",
    "Cristina Merino",
    "Lisseth Aspíllaga",
    "Sandra Chavez",
    "Alejandro Herrera",
    "Katherin Hidalgo",
]

TURNOS = ["Día", "Tarde", "Madrugada"]

# (clave, etiqueta) — orden fijo de las columnas Conforme/No conforme del exportable
PARAM_DEFS = [
    ("brix", "Brix (°B)"),
    ("temperatura", "Temperatura (°C)"),
    ("tiempo", "Tiempo (min)"),
    ("textura", "Textura"),
    ("tamizado", "Tamizado"),
    ("apariencia_color", "Apariencia y Color"),
    ("olor", "Olor"),
    ("observaciones", "Observaciones"),
]

HEADERS_EXPORT = [
    "FECHA", "AREA", "CLIENTE", "N° de Muestra", "Insumo", "Lote (Juliano)",
    "Fecha de Elaboración", "Fecha de Vencimiento", "Conservación",
    "Brix (°B)", "Temperatura (°C)", "Tiempo (min)", "Textura", "Tamizado",
    "Apariencia y Color", "Olor", "Observaciones",
    "Conclusión (C/NC)", "Iniciales",
]


def iniciales(nombre_completo: str) -> str:
    partes = nombre_completo.strip().split()
    if len(partes) < 2:
        return nombre_completo[:2].upper()
    return (partes[0][0] + partes[-1][0]).upper()


def nombre_mes_es(fecha: date) -> str:
    return f"{MESES_ES[fecha.month]} {fecha.year}"


def calcular_juliano(fecha: date) -> str:
    return str(fecha.timetuple().tm_yday).zfill(3)


def campo_aplica(valor) -> bool:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return False
    v = str(valor).strip()
    return v not in ("", "-", "—", "nan", "None")


# ----------------------------------------------------------------------------
# TABLA MIL-STD-105E - Nivel de Inspección Especial S-2, Inspección Rigurosa
# (Tightened, Tabla II-B), AQL 4.0% — misma tabla que en PT y PI
# ----------------------------------------------------------------------------
SAMPLING_TABLE_S2 = [
    (2, 8, "A", 2, 0, 1),
    (9, 15, "A", 2, 0, 1),
    (16, 25, "B", 3, 0, 1),
    (26, 50, "B", 3, 0, 1),
    (51, 90, "B", 3, 0, 1),
    (91, 150, "C", 5, 0, 1),
    (151, 280, "C", 5, 0, 1),
    (281, 500, "C", 5, 0, 1),
    (501, 1200, "D", 8, 1, 2),
    (1201, 3200, "D", 8, 1, 2),
    (3201, 10000, "D", 8, 1, 2),
    (10001, 35000, "E", 13, 1, 2),
    (35001, 150000, "E", 13, 1, 2),
    (150001, 10_000_000, "E", 13, 1, 2),
]


def get_sample_size(lot_size: int):
    for low, high, letter, n, ac, re_ in SAMPLING_TABLE_S2:
        if low <= lot_size <= high:
            return letter, n, ac, re_
    if lot_size < 2:
        return "A", 2, 0, 1
    return "E", 13, 1, 2


# ----------------------------------------------------------------------------
# CARGA Y LIMPIEZA DE DATOS DEL EXCEL
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_specs(excel_bytes: bytes) -> pd.DataFrame:
    df_raw = pd.read_excel(
        io.BytesIO(excel_bytes),
        sheet_name=SHEET_NAME,
        header=None,
        skiprows=3,
    )

    cols = [
        "insumo", "brix", "temperatura", "tiempo", "textura", "tamizado",
        "apariencia_color", "olor", "vida_util", "conservacion",
        "observaciones", "responsable",
    ]
    df_raw = df_raw.iloc[:, : len(cols)]
    df_raw.columns = cols
    df_raw = df_raw[df_raw["insumo"].notna()].copy()

    def clean_txt(x):
        if pd.isna(x):
            return x
        return re.sub(r"\s+", " ", str(x)).strip()

    text_cols = [
        "insumo", "brix", "temperatura", "tiempo", "textura", "tamizado",
        "apariencia_color", "olor", "conservacion", "observaciones",
    ]
    for c in text_cols:
        df_raw[c] = df_raw[c].apply(clean_txt)

    df_raw["insumo"] = df_raw["insumo"].str.upper()
    df_raw["vida_util"] = pd.to_numeric(df_raw["vida_util"], errors="coerce")

    return df_raw.reset_index(drop=True)


def get_excel_bytes():
    try:
        with open(EXCEL_PATH, "rb") as f:
            return f.read()
    except FileNotFoundError:
        st.warning(f"No encontré el archivo `{EXCEL_PATH}` en el repositorio.")
        up = st.file_uploader("Sube el Excel de especificaciones (.xlsx)", type=["xlsx"])
        if up is not None:
            return up.read()
        st.stop()


# ----------------------------------------------------------------------------
# CONEXIÓN A GOOGLE DRIVE / SHEETS
# ----------------------------------------------------------------------------
def get_gsheet_client_and_drive():
    if not GSHEETS_DISPONIBLE:
        return None, None
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        gc = gspread.authorize(creds)
        drive = build_drive_service("drive", "v3", credentials=creds)
        return gc, drive
    except Exception:
        return None, None


def _buscar_archivo_en_carpeta(drive, nombre: str, carpeta_id: str, es_carpeta: bool = False):
    tipo_mime = "application/vnd.google-apps.folder" if es_carpeta else "application/vnd.google-apps.spreadsheet"
    query = (
        f"name = '{nombre}' and '{carpeta_id}' in parents "
        f"and mimeType = '{tipo_mime}' and trashed = false"
    )
    resultado = drive.files().list(q=query, fields="files(id, name)").execute()
    archivos = resultado.get("files", [])
    return archivos[0]["id"] if archivos else None


def get_or_create_month_spreadsheet_id(drive, root_folder_id: str, fecha: date) -> str:
    subcarpeta_id = _buscar_archivo_en_carpeta(drive, SUBCARPETA_DRIVE, root_folder_id, es_carpeta=True)
    if subcarpeta_id is None:
        raise RuntimeError(f"No encontré la subcarpeta '{SUBCARPETA_DRIVE}' dentro de la carpeta raíz.")

    nombre_mes = nombre_mes_es(fecha)
    archivo_mes_id = _buscar_archivo_en_carpeta(drive, nombre_mes, subcarpeta_id)
    if archivo_mes_id:
        return archivo_mes_id

    raise RuntimeError(
        f"Todavía no existe el archivo del mes '{nombre_mes}' dentro de "
        f"'{SUBCARPETA_DRIVE}'. Duplica manualmente '{PLANTILLA_NOMBRE}' en esa "
        f"carpeta de Drive y renómbralo exactamente '{nombre_mes}' (Google no "
        f"permite que la cuenta de servicio cree archivos nuevos automáticamente)."
    )


def _encontrar_fila_footer(ws) -> int | None:
    valores = ws.get_all_values()
    for idx, fila in enumerate(valores, start=1):
        for celda in fila:
            if celda.strip().upper().startswith(FOOTER_MARCA.upper()):
                return idx
    return None


def get_or_create_daily_worksheet(spreadsheet, fecha: date, turno: str):
    titulo_hoja = f"{fecha.strftime('%Y-%m-%d')} {turno}"
    try:
        return spreadsheet.worksheet(titulo_hoja)
    except gspread.exceptions.WorksheetNotFound:
        pass

    plantilla = spreadsheet.worksheet(TEMPLATE_SHEET_NAME)
    ws = spreadsheet.duplicate_sheet(source_sheet_id=plantilla.id, new_sheet_name=titulo_hoja)

    fila_footer = _encontrar_fila_footer(ws)
    primera_fila_datos = FILA_ENCABEZADO_PLANTILLA + 1
    if fila_footer and fila_footer > primera_fila_datos:
        ws.delete_rows(primera_fila_datos, fila_footer - 1)

    return ws


def guardar_en_google_sheets(filas: list) -> tuple:
    gc, drive = get_gsheet_client_and_drive()
    if gc is None or drive is None:
        return False, "No se pudo conectar a Google Drive/Sheets (revisa los Secrets configurados)."
    try:
        root_folder_id = st.secrets.get("ROOT_FOLDER_ID")
        if not root_folder_id:
            return False, "No se encontró ROOT_FOLDER_ID en los Secrets."

        fecha = st.session_state.fecha_elaboracion
        turno = st.session_state.turno

        spreadsheet_id = get_or_create_month_spreadsheet_id(drive, root_folder_id, fecha)
        spreadsheet = gc.open_by_key(spreadsheet_id)
        ws = get_or_create_daily_worksheet(spreadsheet, fecha, turno)

        fila_footer = _encontrar_fila_footer(ws)
        if fila_footer is None:
            ws.append_rows(filas)
        else:
            ws.insert_rows(filas, row=fila_footer)

        return True, f"Guardado en '{nombre_mes_es(fecha)}' → hoja '{fecha.strftime('%Y-%m-%d')} {turno}'."
    except Exception as e:
        return False, f"Error al guardar en Google Sheets: {e}"


# ----------------------------------------------------------------------------
# ESTADO / WIZARD
# ----------------------------------------------------------------------------
if "step" not in st.session_state:
    st.session_state.step = 1


def go_next():
    st.session_state.step += 1


def go_back():
    st.session_state.step -= 1


excel_bytes = get_excel_bytes()
specs_df = load_specs(excel_bytes)

# ============================================================================
# PASO 1 - PORTADA
# ============================================================================
if st.session_state.step == 1:
    st.title("🍯 Registro de Parámetros de Calidad")
    st.subheader("Insumos de Cocina Dulce - B2B Starbucks")
    st.markdown(
        """
        Esta aplicación permite al **equipo de calidad** registrar la inspección
        de insumos elaborados en cocina dulce (fudge, mermeladas, rellenos, etc.),
        siguiendo el plan de muestreo **MIL-STD-105E (Nivel Especial S-2,
        Inspección Rigurosa, AQL 4.0%)**.
        """
    )
    st.button("Comenzar registro ➜", on_click=go_next, type="primary")

# ============================================================================
# PASO 2 - EQUIPO DE CALIDAD, TURNO Y FECHA
# ============================================================================
elif st.session_state.step == 2:
    st.header("1️⃣ Datos del registro")

    responsable = st.selectbox(
        "Equipo de calidad (responsable del registro)",
        EQUIPO_CALIDAD,
        index=EQUIPO_CALIDAD.index(st.session_state.get("responsable", EQUIPO_CALIDAD[0]))
        if st.session_state.get("responsable") in EQUIPO_CALIDAD else 0,
    )

    turno = st.selectbox(
        "Turno",
        TURNOS,
        index=TURNOS.index(st.session_state.get("turno", TURNOS[0]))
        if st.session_state.get("turno") in TURNOS else 0,
    )

    fecha_elaboracion = st.date_input(
        "Fecha de elaboración (= fecha de registro)",
        value=st.session_state.get("fecha_elaboracion", None),
        format="DD/MM/YYYY",
    )
    if fecha_elaboracion is None:
        st.caption("⚠️ Selecciona la fecha para continuar (no se precarga sola).")

    st.info(f"**Cliente:** {CLIENTE_FIJO}  |  **Área:** {AREA_FIJA}")

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary", disabled=fecha_elaboracion is None):
            st.session_state.responsable = responsable
            st.session_state.turno = turno
            st.session_state.fecha_elaboracion = fecha_elaboracion
            go_next()
            st.rerun()

# ============================================================================
# PASO 3 - INSUMO
# ============================================================================
elif st.session_state.step == 3:
    st.header("2️⃣ Insumo")

    insumos = sorted(specs_df["insumo"].dropna().unique().tolist())
    insumo_sel = st.selectbox("Selecciona el insumo", insumos)
    fila = specs_df[specs_df["insumo"] == insumo_sel].iloc[0]

    with st.expander("📋 Ficha de referencia", expanded=True):
        for clave, label in PARAM_DEFS:
            valor = fila[clave]
            if campo_aplica(valor):
                st.markdown(f"**{label}:** {valor}")
        if campo_aplica(fila["vida_util"]):
            st.markdown(f"**Vida útil:** {int(fila['vida_util'])} días")
        st.markdown(f"**Conservación:** {fila['conservacion']}")

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.insumo = insumo_sel
            st.session_state.fila_insumo = fila.to_dict()
            go_next()
            st.rerun()

# ============================================================================
# PASO 4 - LOTE (CÓDIGO JULIANO)
# ============================================================================
elif st.session_state.step == 4:
    st.header("3️⃣ Lote (código Juliano)")

    modo_juliano = st.radio(
        "¿Cómo deseas obtener el código Juliano?",
        ["Calcular automáticamente (según fecha de elaboración)", "Ingresar manualmente"],
    )

    if modo_juliano.startswith("Calcular"):
        juliano = calcular_juliano(st.session_state.fecha_elaboracion)
        st.success(f"Código Juliano calculado: **{juliano}** "
                   f"(día N° {int(juliano)} del año {st.session_state.fecha_elaboracion.year})")
    else:
        juliano = st.text_input("Ingresa el código Juliano manualmente", value="")

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary", disabled=not str(juliano).strip()):
            st.session_state.lote_juliano = str(juliano).strip()
            go_next()
            st.rerun()

# ============================================================================
# PASO 5 - BATCH Y MUESTREO
# ============================================================================
elif st.session_state.step == 5:
    st.header("4️⃣ Tamaño del batch y muestreo (MIL-STD-105E, S-2, Rigurosa, AQL 4.0%)")

    batch_size = st.number_input(
        "¿Cuántas unidades tiene el batch?", min_value=1, step=1,
        value=st.session_state.get("batch_size", 300),
    )
    letra, n_muestras, ac, re_ = get_sample_size(int(batch_size))
    st.success(
        f"Para un batch de **{int(batch_size)}** unidades: letra código **{letra}** → "
        f"**n = {n_muestras}** muestras. Criterio: Aceptar con **{ac}** o menos no conformes, "
        f"Rechazar con **{re_}** o más."
    )

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.batch_size = int(batch_size)
            st.session_state.n_muestras = n_muestras
            st.session_state.letra_codigo = letra
            go_next()
            st.rerun()

# ============================================================================
# PASO 6 - REGISTRO DE PARÁMETROS POR MUESTRA
# ============================================================================
elif st.session_state.step == 6:
    st.header("5️⃣ Registro de parámetros por muestra")

    fila = st.session_state.fila_insumo
    n = st.session_state.n_muestras

    st.markdown(f"**Insumo:** {st.session_state.insumo} &nbsp;|&nbsp; "
                f"**N° de muestras a evaluar:** {n}")
    st.info(f"**Conservación de referencia (informativa, no se pregunta):** {fila['conservacion']}")

    parametros = []
    for clave, label in PARAM_DEFS:
        valor = fila[clave]
        if not campo_aplica(valor):
            continue
        pregunta = f"¿{label} cumple con lo especificado? ({valor})"
        parametros.append((clave, label, pregunta))

    st.caption("Completa cada muestra tocando 'Conforme' o 'No conforme'. Si algún parámetro "
               "tiene 1 o más 'No conforme', se habilitará un cuadro de comentario al final de esa pestaña.")

    if "respuestas" not in st.session_state:
        st.session_state.respuestas = {}
    if "comentarios_parametro" not in st.session_state:
        st.session_state.comentarios_parametro = {}

    tabs = st.tabs([label for _, label, _ in parametros])
    for (clave, label, pregunta), tab in zip(parametros, tabs):
        with tab:
            st.markdown(f"**{pregunta}**")
            hay_no_conforme = False
            for i in range(n):
                key = f"resp_{clave}_{i}"
                if key not in st.session_state.respuestas:
                    st.session_state.respuestas[key] = "Conforme"
                valor_actual = st.session_state.respuestas[key]
                idx_default = 0 if valor_actual == "Conforme" else 1
                st.session_state.respuestas[key] = st.radio(
                    f"Muestra {i + 1}",
                    ["Conforme", "No conforme"],
                    index=idx_default,
                    horizontal=True,
                    key=f"widget_{key}",
                )
                if st.session_state.respuestas[key] == "No conforme":
                    hay_no_conforme = True

            if hay_no_conforme:
                st.session_state.comentarios_parametro[clave] = st.text_area(
                    f"Comentario / corrección para '{label}' (hay al menos 1 muestra No conforme)",
                    value=st.session_state.comentarios_parametro.get(clave, ""),
                    key=f"comentario_{clave}",
                )
            else:
                st.session_state.comentarios_parametro[clave] = ""

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Siguiente ➜", type="primary"):
            st.session_state.parametros_muestra = parametros
            go_next()
            st.rerun()

# ============================================================================
# PASO 7 - CONCLUSIÓN DEL REGISTRO
# ============================================================================
elif st.session_state.step == 7:
    st.header("6️⃣ Conclusión del registro")

    conclusion = st.radio(
        "Conclusión",
        ["Conforme", "No conforme"],
        index=0 if st.session_state.get("conclusion", "Conforme") == "Conforme" else 1,
        horizontal=True,
    )

    col1, col2 = st.columns(2)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        if st.button("Generar resumen ➜", type="primary"):
            st.session_state.conclusion = conclusion
            go_next()
            st.rerun()

# ============================================================================
# PASO 8 - RESUMEN FINAL, GUARDADO EN SHEETS Y EXPORTACIÓN
# ============================================================================
elif st.session_state.step == 8:
    st.header("7️⃣ Resumen final")

    fila = st.session_state.fila_insumo
    n = st.session_state.n_muestras
    parametros = st.session_state.parametros_muestra
    claves_activas = [clave for clave, _, _ in parametros]

    if campo_aplica(fila["vida_util"]):
        vida_util_dias = int(fila["vida_util"])
        fecha_venc = st.session_state.fecha_elaboracion + timedelta(days=vida_util_dias)
        fecha_venc_txt = fecha_venc.strftime("%d/%m/%Y")
    else:
        fecha_venc_txt = "No aplica"

    st.subheader("Datos del registro")
    resumen_info = pd.DataFrame(
        {
            "Campo": [
                "Equipo de calidad", "Turno", "Fecha de elaboración", "Cliente", "Área",
                "Insumo", "Lote (Juliano)", "Fecha de vencimiento", "Conservación",
                "Tamaño de batch", "Letra código muestreo", "N° de muestras", "Conclusión",
            ],
            "Valor": [
                st.session_state.responsable, st.session_state.turno,
                st.session_state.fecha_elaboracion.strftime("%d/%m/%Y"),
                CLIENTE_FIJO, AREA_FIJA,
                st.session_state.insumo, st.session_state.lote_juliano, fecha_venc_txt,
                fila["conservacion"],
                st.session_state.batch_size, st.session_state.letra_codigo, n,
                st.session_state.conclusion,
            ],
        }
    )
    st.dataframe(resumen_info, use_container_width=True, hide_index=True)

    st.subheader("Resultados por parámetro (muestras)")
    conteo = {}
    for clave, label, _ in parametros:
        valores = [st.session_state.respuestas[f"resp_{clave}_{i}"] for i in range(n)]
        conteo[label] = {
            "Conforme": valores.count("Conforme"),
            "No conforme": valores.count("No conforme"),
        }
    st.dataframe(pd.DataFrame(conteo).T, use_container_width=True)

    # ------------------------------------------------------------------
    # Armado de las filas exportables (una fila por muestra)
    # ------------------------------------------------------------------
    valores_por_parametro = {
        clave: [st.session_state.respuestas[f"resp_{clave}_{i}"] for i in range(n)]
        for clave, _, _ in parametros
    }

    def valor_muestra(clave, i):
        return valores_por_parametro.get(clave, [""] * n)[i] if clave in valores_por_parametro else ""

    filas_export = []
    for i in range(n):
        filas_export.append([
            st.session_state.fecha_elaboracion.strftime("%d/%m/%Y"),
            AREA_FIJA,
            CLIENTE_FIJO,
            i + 1,
            st.session_state.insumo,
            st.session_state.lote_juliano,
            st.session_state.fecha_elaboracion.strftime("%d/%m/%Y"),
            fecha_venc_txt,
            fila["conservacion"],
            valor_muestra("brix", i),
            valor_muestra("temperatura", i),
            valor_muestra("tiempo", i),
            valor_muestra("textura", i),
            valor_muestra("tamizado", i),
            valor_muestra("apariencia_color", i),
            valor_muestra("olor", i),
            valor_muestra("observaciones", i),
            st.session_state.conclusion,
            iniciales(st.session_state.responsable),
        ])

    export_df = pd.DataFrame(filas_export, columns=HEADERS_EXPORT)

    with st.expander("Ver todas las filas que se guardarán/exportarán"):
        st.dataframe(export_df, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Guardar historial")
    if st.button("💾 Guardar en Google Sheets (historial)", type="primary"):
        exito, mensaje = guardar_en_google_sheets(filas_export)
        if exito:
            st.success(mensaje)
        else:
            st.error(mensaje)

    csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Registro")
    excel_buffer.seek(0)

    st.divider()
    col1, col2, col3 = st.columns(3)
    with col1:
        st.button("⬅ Atrás", on_click=go_back)
    with col2:
        st.download_button(
            "⬇ Descargar CSV",
            data=csv_bytes,
            file_name=f"registro_COCINA_{st.session_state.insumo}_{st.session_state.fecha_elaboracion}.csv",
            mime="text/csv",
        )
    with col3:
        st.download_button(
            "⬇ Descargar Excel",
            data=excel_buffer,
            file_name=f"registro_COCINA_{st.session_state.insumo}_{st.session_state.fecha_elaboracion}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.divider()
    if st.button("🔄 Nuevo registro"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
