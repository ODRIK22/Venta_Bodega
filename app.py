import streamlit as st
import pandas as pd
import unicodedata
import re
import os
import io

# Configuración inicial de la página Streamlit
st.set_page_config(
    page_title="Directorio de Prospectos - Radio de 20 km",
    page_icon="🏭",
    layout="wide"
)

st.title("Directorio de Prospectos - Radio de 20 km")
st.caption("Filtro e historial de prospección comercial B2B para renta de bodegas")

MUNICIPIOS_OBJETIVO_NORM = [
    "pachuca de soto",
    "mineral de la reforma",
    "zempoala",
    "san agustin tlaxiaca"
]

EMPRESAS_CONTACTADAS = [
    {"id": 1, "nombre": "TRANSTELL", "tel": "7717170909"},
    {"id": 2, "nombre": "Logística Inteligente Nacional", "tel": "5546084544"},
    {"id": 3, "nombre": "RMZ Logistics International Freight Forwarder", "tel": "7717119304"},
    {"id": 4, "nombre": "Logística, Transporte y Distribución de México", "tel": "55120912728"},
    {"id": 5, "nombre": "REFACCIONARIA APYMSA", "tel": "3332084420"},
    {"id": 6, "nombre": "Refaccionaria Reyes", "tel": "7717143791"},
    {"id": 7, "nombre": "Auto Volk's Pachuca Tulipanes", "tel": "7711986688"},
    {"id": 8, "nombre": "Distribuidora de Hidalgo", "tel": "7717143607"},
    {"id": 9, "nombre": "Distribuidora de Hidalgo Matilde", "tel": "7711009614"},
    {"id": 10, "nombre": "Distribuidora del Sureste", "tel": ""},
    {"id": 11, "nombre": "Distribuidora de Hidalgo (Carr. Pachuca-Actopan)", "tel": "7711483290"}
]

ARCHIVO_HISTORIAL = "seguimiento_prospectos.csv"

def normalizar_texto(texto):
    """Normaliza texto eliminando acentos y conservando alfanuméricos y espacios."""
    if not isinstance(texto, str):
        return ""
    s = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8').lower()
    return re.sub(r'[^a-z0-9\s]', '', s).strip()

def resolver_columna(columnas_df, nombres_posibles):
    """Resuelve dinámicamente nombres de columnas aceptando claves DENUE o encabezados completos."""
    columnas_norm = [normalizar_texto(c) for c in columnas_df]
    
    for posible in nombres_posibles:
        posible_norm = normalizar_texto(posible)
        for i, col_norm in enumerate(columnas_norm):
            if col_norm == posible_norm:
                return columnas_df[i]
                
    for posible in nombres_posibles:
        posible_norm = normalizar_texto(posible)
        for i, col_norm in enumerate(columnas_norm):
            if posible_norm in col_norm and 'clave' not in col_norm:
                return columnas_df[i]
                
    return None

def cargar_historial_guardado():
    """Carga las marcas de verificación de empresas contactadas desde el archivo de persistencia."""
    if os.path.exists(ARCHIVO_HISTORIAL):
        try:
            return pd.read_csv(ARCHIVO_HISTORIAL, dtype=str)
        except Exception:
            return pd.DataFrame(columns=["Nombre_Norm", "Contactado", "Notas"])
    return pd.DataFrame(columns=["Nombre_Norm", "Contactado", "Notas"])

def guardar_historial(df_actual):
    """Guarda las marcas de verificación y notas en el archivo CSV persistente."""
    try:
        registros_guardar = []
        for _, row in df_actual.iterrows():
            if row.get("Contactado", False) or (row.get("Notas") and str(row.get("Notas")).strip() not in ["", "No disponible"]):
                registros_guardar.append({
                    "Nombre_Norm": normalizar_texto(row["Nombre"]),
                    "Nombre": row["Nombre"],
                    "Contactado": "True" if row.get("Contactado", False) else "False",
                    "Notas": str(row.get("Notas", ""))
                })
        
        df_save = pd.DataFrame(registros_guardar)
        df_save.to_csv(ARCHIVO_HISTORIAL, index=False, encoding="utf-8")
    except Exception as e:
        st.error(f"Error al guardar el historial: {e}")

def es_empresa_contactada_previa(row):
    """Evalúa si la empresa pertenece a la lista inicial de 11 empresas trabajadas por la compañera."""
    nom_denue = normalizar_texto(str(row.get("Nombre", "")))
    raz_denue = normalizar_texto(str(row.get("Razón Social", "")))
    tel_denue = re.sub(r'\D', '', str(row.get("Teléfono", "")))

    for emp in EMPRESAS_CONTACTADAS:
        nom_comp = normalizar_texto(emp["nombre"])
        tel_comp = re.sub(r'\D', '', emp["tel"])

        if len(tel_comp) >= 7 and len(tel_denue) >= 7 and (tel_comp[-7:] in tel_denue or tel_denue[-7:] in tel_comp):
            return True

        if nom_comp and len(nom_comp) >= 5 and (nom_comp in nom_denue or nom_comp in raz_denue):
            return True

        palabras_distintivas = ["apymsa", "transtell", "rmz", "volks"]
        for p in palabras_distintivas:
            if p in nom_comp and (p in nom_denue or p in raz_denue):
                return True

    return False

def generar_excel_formateado(df):
    """Genera un archivo ejecutable nativo de Excel (.xlsx) con columnas anchas y encabezados elegantes."""
    output = io.BytesIO()
    df_export = df.copy()
    
    # Convertir el booleano 'Contactado' a texto estructurado "Sí" / "No"
    if "Contactado" in df_export.columns:
        df_export["Contactado"] = df_export["Contactado"].map({True: "Sí", False: "No"})

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name="Prospectos_Bodegas")
        worksheet = writer.sheets["Prospectos_Bodegas"]
        
        # Estilos para encabezados (Fondo azul oscuro, texto blanco en negrita)
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        for col_idx, col in enumerate(worksheet.columns, start=1):
            max_len = 0
            for cell in col:
                # Estilos de encabezado
                if cell.row == 1:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = center_align
                else:
                    cell.border = thin_border
                
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)

            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 4, 14)

    return output.getvalue()

def cargar_y_procesar_datos():
    file_path = "INEGI_DENUE_06082026.csv"
    try:
        df = pd.read_csv(file_path, encoding="latin1", dtype=str, low_memory=False)
    except FileNotFoundError:
        st.error(f"No se encontró el archivo '{file_path}' en la raíz del proyecto.")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error al leer el archivo CSV: {e}")
        return pd.DataFrame()

    reglas_columnas = {
        "Nombre": ["nom_estab", "nombre de la unidad economica", "nombre comercial"],
        "Razón Social": ["raz_social", "razon social"],
        "Giro": ["nombre_act", "nombre de clase de la actividad", "giro", "actividad"],
        "Municipio": ["municipio"],
        "Teléfono": ["telefono", "numero de telefono"],
        "Correo": ["correoelec", "correo electronico", "correo"],
        "Sitio Web": ["www", "sitio en internet", "sitio web", "pagina web"]
    }

    columnas_mapeadas = {}
    for nombre_final, alternativas in reglas_columnas.items():
        col_encontrada = resolver_columna(df.columns.tolist(), alternativas)
        if col_encontrada:
            columnas_mapeadas[nombre_final] = col_encontrada

    if "Municipio" not in columnas_mapeadas:
        st.error("No se encontró la columna de Municipio en el archivo CSV.")
        return pd.DataFrame()

    df_filtrado = df[list(columnas_mapeadas.values())].copy()
    inverso_mapeo = {v: k for k, v in columnas_mapeadas.items()}
    df_filtrado = df_filtrado.rename(columns=inverso_mapeo)

    # Filtro Geográfico (20 km)
    df_filtrado["muni_norm"] = df_filtrado["Municipio"].apply(normalizar_texto)
    df_filtrado = df_filtrado[df_filtrado["muni_norm"].isin(MUNICIPIOS_OBJETIVO_NORM)].drop(columns=["muni_norm"])

    # Formatear la columna 'Sitio Web'
    if "Sitio Web" in df_filtrado.columns:
        def formatear_url(url):
            if pd.isna(url) or not str(url).strip() or str(url).strip().lower() in ["no disponible", "nan", "none", "null", "0"]:
                return None
            url_str = str(url).strip()
            if not (url_str.startswith("http://") or url_str.startswith("https://")):
                return f"https://{url_str}"
            return url_str

        df_filtrado["Sitio Web"] = df_filtrado["Sitio Web"].apply(formatear_url)

    # Limpiar nulos
    columnas_texto = [c for c in df_filtrado.columns if c != "Sitio Web"]
    for col in columnas_texto:
        df_filtrado[col] = df_filtrado[col].astype(str).replace(["nan", "None", "null", "<NA>", "", "0"], pd.NA).fillna("No disponible")

    # Cargar historial guardado de marcas previas
    df_historial = cargar_historial_guardado()
    dict_contactados = {}
    dict_notas = {}
    if not df_historial.empty:
        for _, hrow in df_historial.iterrows():
            key = hrow["Nombre_Norm"]
            dict_contactados[key] = (hrow["Contactado"] == "True")
            dict_notas[key] = hrow.get("Notas", "")

    # Determinar estado de Checkbox y Estatus CRM
    contactado_check_list = []
    notas_list = []
    estatus_crm_list = []

    for _, row in df_filtrado.iterrows():
        nombre_norm = normalizar_texto(row["Nombre"])
        es_previo = es_empresa_contactada_previa(row)
        
        esta_marcado = dict_contactados.get(nombre_norm, es_previo)
        nota_guardada = dict_notas.get(nombre_norm, "")

        contactado_check_list.append(esta_marcado)
        notas_list.append(nota_guardada)
        estatus_crm_list.append("Ya Contactado" if esta_marcado else "Nuevo Prospecto")

    # Insertar columna Checkbox en primer lugar
    df_filtrado.insert(0, "Contactado", contactado_check_list)
    df_filtrado["Notas"] = notas_list
    df_filtrado["Estatus CRM"] = estatus_crm_list

    return df_filtrado

# Carga de datos
df_todos = cargar_y_procesar_datos()

if not df_todos.empty:
    # Sidebar: Controles
    st.sidebar.header("⚙️ Filtros de Prospección")
    
    excluir_contactados = st.sidebar.checkbox(
        "Excluir prospectos ya contactados",
        value=True,
        help="Oculta las empresas marcadas con check o trabajadas previamente."
    )

    municipios_disponibles = sorted(df_todos["Municipio"].unique().tolist())
    municipios_seleccionados = st.sidebar.multiselect(
        "Filtrar por Municipio",
        options=municipios_disponibles,
        default=municipios_disponibles
    )

    busqueda_empresa = st.sidebar.text_input(
        "Buscar por Nombre o Giro",
        placeholder="Ej. Logística, Refaccionaria, Almacén..."
    )

    # Aplicar filtros a la vista
    df_display = df_todos.copy()

    if excluir_contactados:
        df_display = df_display[~df_display["Contactado"]]

    if municipios_seleccionados:
        df_display = df_display[df_display["Municipio"].isin(municipios_seleccionados)]

    if busqueda_empresa:
        query = normalizar_texto(busqueda_empresa)
        mask = (
            df_display["Nombre"].apply(normalizar_texto).str.contains(query, na=False) |
            df_display["Razón Social"].apply(normalizar_texto).str.contains(query, na=False) |
            df_display["Giro"].apply(normalizar_texto).str.contains(query, na=False)
        )
        df_display = df_display[mask]

    # Métricas de Resumen
    c1, c2, c3 = st.columns(3)
    total_en_denue = len(df_todos)
    total_contactados = len(df_todos[df_todos["Contactado"]])
    total_nuevos = len(df_todos[~df_todos["Contactado"]])

    with c1:
        st.metric(label="Total en DENUE (20 km)", value=f"{total_en_denue} empresas")
    with c2:
        st.metric(label="Prospectos Nuevos Disponibles", value=f"{total_nuevos} disponibles")
    with c3:
        st.metric(label="Empresas Contactadas / Excluidas", value=f"{total_contactados} marcas")

    st.markdown("---")

    # Cabecera de controles superiores (Boton de descarga de Excel)
    top_col1, top_col2 = st.columns([3, 1])
    with top_col1:
        st.info("💡 **Tip:** Puedes marcar la casilla **Contactado** o escribir **Notas** directamente en la tabla. Se guardarán automáticamente.")
    with top_col2:
        excel_bytes = generar_excel_formateado(df_display)
        st.download_button(
            label="📊 Exportar a Excel (.xlsx)",
            data=excel_bytes,
            file_name="Prospectos_Bodegas_Radio20km.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Descarga la lista actual formateada en filas y columnas nativas de Microsoft Excel."
        )

    # Configuración de columnas interactivas con st.data_editor
    column_config = {
        "Contactado": st.column_config.CheckboxColumn(
            "Contactado",
            help="Marca para indicar que ya fue contactada",
            default=False
        ),
        "Sitio Web": st.column_config.LinkColumn(
            "Sitio Web",
            help="Haz clic para visitar el sitio web oficial",
            display_text="Visitar sitio"
        ),
        "Notas": st.column_config.TextColumn(
            "Notas de Seguimiento",
            help="Escribe comentarios o estatus de llamada",
            width="medium"
        ),
        "Nombre": st.column_config.TextColumn("Nombre", disabled=True),
        "Razón Social": st.column_config.TextColumn("Razón Social", disabled=True),
        "Giro": st.column_config.TextColumn("Giro", disabled=True),
        "Municipio": st.column_config.TextColumn("Municipio", disabled=True),
        "Teléfono": st.column_config.TextColumn("Teléfono", disabled=True),
        "Correo": st.column_config.TextColumn("Correo", disabled=True),
        "Estatus CRM": st.column_config.TextColumn("Estatus CRM", disabled=True)
    }

    # Renderizar la tabla interactiva
    df_editado = st.data_editor(
        df_display,
        column_config=column_config,
        hide_index=True,
        use_container_width=True,
        key="editor_prospectos_persistente"
    )

    # Guardar automáticamente la persistencia en CSV
    guardar_historial(df_editado)

else:
    st.warning("No se encontraron empresas en el archivo.")
