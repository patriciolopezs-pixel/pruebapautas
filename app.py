import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
import io
import datetime
import re

# Configuración de la página web
st.set_page_config(page_title="Portal de Pautas de Supervisión - RedSalud", layout="centered")

# --- ESTILOS CSS CORPORATIVOS Y TARJETAS ---
st.markdown("""
<style>
    /* Fondo global */
    .stApp {
        background-color: #F4F7F6 !important;
    }
    
    /* Encabezados */
    h1, h2, h3, h4 {
        color: #00205B !important;
        font-family: 'Segoe UI', Tahoma, sans-serif !important;
        font-weight: 700 !important;
    }
    
    .stMarkdown, label, p, span {
        color: #00205B !important;
    }

    /* Campos de entrada */
    input, select, textarea, div[data-baseweb="select"] {
        background-color: #FFFFFF !important;
        color: #00205B !important;
        -webkit-text-fill-color: #00205B !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
        text-transform: uppercase !important;
    }

    /* Botón Principal (Turquesa RedSalud) */
    button[kind="primary"] {
        background-color: #00828A !important;
        border: none !important;
        border-radius: 8px !important;
    }
    button[kind="primary"] p {
        color: #FFFFFF !important;
        font-weight: bold !important;
    }

    /* Botón Secundario / Descarga / Volver */
    button[kind="secondary"], .stDownloadButton button {
        background-color: #00205B !important;
        border: none !important;
        border-radius: 8px !important;
    }
    button[kind="secondary"] p, .stDownloadButton button p {
        color: #FFFFFF !important;
        font-weight: bold !important;
    }

    /* Estilo para las Tarjetas de la Pantalla de Inicio */
    .card-pauta {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-top: 6px solid #00828A;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 12px rgba(0, 32, 91, 0.08);
    }

    .card-pauta-future {
        background-color: #FAFAFA;
        border: 1px dashed #CBD5E1;
        border-top: 6px solid #B0BEC5;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 20px;
        opacity: 0.8;
    }
</style>
""", unsafe_allow_html=True)


# --- 1. ESTADO DE LA SESIÓN Y NAVEGACIÓN ---
if 'pagina_activa' not in st.session_state:
    st.session_state.pagina_activa = "inicio"

if 'pautas_data' not in st.session_state:
    st.session_state.pautas_data = {i: None for i in range(1, 19)}

if 'pauta_actual' not in st.session_state:
    st.session_state.pauta_actual = 1

if 'ultimo_centro' not in st.session_state:
    st.session_state.ultimo_centro = ""


def parse_fecha(fecha_str):
    if fecha_str:
        try:
            return datetime.datetime.strptime(str(fecha_str).strip(), "%d-%m-%Y").date()
        except Exception:
            try:
                return pd.to_datetime(fecha_str).date()
            except Exception:
                pass
    return datetime.date.today()


# --- 2. LÓGICA DE CARGA MASIVA ---
def procesar_df_masivo(df, fecha_sup_default):
    df.columns = [str(c).strip().lower() for c in df.columns]
    
    def get_val(row, candidates, default=""):
        for cand in candidates:
            for col in df.columns:
                if cand in col:
                    val = str(row[col]).strip()
                    return val if val != 'nan' and val != 'None' else default
        return default

    count = 0
    for i, row in df.iterrows():
        if count >= 18:
            break
        num_pauta = count + 1
        
        centro_val = re.sub(r'[^a-zA-ZáéíóúÁÉÍÓÚñÑ\s]', '', get_val(row, ['centro'])).strip().upper()
        rut_val = re.sub(r'[^0-9kK\-]', '', get_val(row, ['rut paciente', 'rut_paciente'])).strip().upper()
        
        nom_prof = get_val(row, ['nombre profesional', 'nombre realizador', 'nombre persona'])
        ape_prof = get_val(row, ['apellidos profesional', 'apellido realizador', 'apellidos persona', 'apellido persona'])
        
        nombre_clean = re.sub(r'[^a-zA-ZáéíóúÁÉÍÓÚñÑ\s]', '', nom_prof).strip().upper()
        apellido_clean = re.sub(r'[^a-zA-ZáéíóúÁÉÍÓÚñÑ\s]', '', ape_prof).strip().upper()
        
        fecha_atencion_raw = get_val(row, ['fecha de ejecución', 'fecha atencion', 'fecha_atencion', 'fecha'])
        try:
            fecha_atencion_clean = pd.to_datetime(fecha_atencion_raw).strftime("%d-%m-%Y")
        except Exception:
            fecha_atencion_clean = fecha_sup_default.strftime("%d-%m-%Y")
            
        serv_raw = get_val(row, ['servicio']).lower()
        if 'pabell' in serv_raw or 'pd' in serv_raw:
            serv_clean = "Pabellón de Cirugía menor Dental (PD)"
        elif 'imagen' in serv_raw or 'rx' in serv_raw:
            serv_clean = "Imagenología Dental (RX)"
        else:
            serv_clean = "Sala de Procedimiento Dental (BD)"
            
        exo_raw = get_val(row, ['exodoncia', 'prestación corresponde']).upper()
        exo_clean = "SI" if "SI" in exo_raw else "NO"

        cumple_existente = "SI"
        if st.session_state.pautas_data[num_pauta] is not None:
            cumple_existente = st.session_state.pautas_data[num_pauta].get("cumple", "SI")

        st.session_state.pautas_data[num_pauta] = {
            "centro": centro_val if centro_val else st.session_state.ultimo_centro,
            "fecha_sup": fecha_sup_default.strftime("%d-%m-%Y"),
            "nombre": nombre_clean,
            "apellido": apellido_clean,
            "rut": rut_val,
            "fecha_atencion": fecha_atencion_clean,
            "servicio": serv_clean,
            "exodoncia": exo_clean,
            "cumple": cumple_existente
        }
        if centro_val:
            st.session_state.ultimo_centro = centro_val
        count += 1

    return count


# --- 3. GENERADOR EXCEL CONSOLIDADO ---
def generar_excel_consolidado(pautas_dict):
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidado Pautas"

    thin_border = Border(
        left=Side(style='thin', color='000000'), 
        right=Side(style='thin', color='000000'), 
        top=Side(style='thin', color='000000'), 
        bottom=Side(style='thin', color='000000')
    )
    center_aligned_text = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_aligned_text = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    bold_font_white = Font(bold=True, color="FFFFFF")
    bold_font_navy = Font(bold=True, color="00205B")
    
    navy_header_fill = PatternFill(start_color="00205B", end_color="00205B", fill_type="solid")
    teal_sub_fill = PatternFill(start_color="00828A", end_color="00828A", fill_type="solid")
    soft_teal_fill = PatternFill(start_color="E6F7F5", end_color="E6F7F5", fill_type="solid")

    ws.merge_cells('A1:AK1')
    ws['A1'] = "PAUTA DE SUPERVISIÓN CUMPLIMIENTO DE PAUSA DE SEGURIDAD DENTAL EN BOX DENTAL, PABELLÓN DE CIRUGÍA MENOR DENTAL E IMAGENOLOGÍA DENTAL (GCL 2.1 AO)"
    ws['A1'].font = bold_font_white
    ws['A1'].fill = navy_header_fill
    ws['A1'].alignment = center_aligned_text

    ws['A2'] = "Indicaciones llenado pauta"
    ws.merge_cells('B2:AK2')
    ws['B2'] = "Marque √ SI cumple, Marque X NO cumple, o No Aplica (si corresponde). En ítem cumple registre SI o NO. Registre en observaciones motivo incumplimiento."
    ws['A2'].font = bold_font_navy
    ws['B2'].alignment = center_aligned_text

    etiquetas = [
        "Centro", 
        "Fecha de Supervisión", 
        "Nombre de la persona supervisada", 
        "Apellido(s) de la persona supervisada", 
        "RUT del paciente", 
        "Fecha de Atención supervisada", 
        "Servicio Clínico donde se realizó el procedimiento, ya sea Sala de Procedimiento Dental (BD), Pabellón de Cirugía menor Dental (PD), e Imagenología Dental (RX)",
        "Procedimiento corresponde a Exodoncia (SI, NO)"
    ]

    for i, etiqueta in enumerate(etiquetas, start=3):
        ws.cell(row=i, column=1, value=etiqueta).font = bold_font_navy
        ws.cell(row=i, column=1).alignment = left_aligned_text
        ws.cell(row=i, column=1).fill = soft_teal_fill

    ws.column_dimensions['A'].width = 50

    ws.cell(row=11, column=1, value="N° DE PAUTA").font = bold_font_white
    ws.cell(row=11, column=1).fill = teal_sub_fill
    
    ws.cell(row=12, column=1, value="CRITERIOS A EVALUAR").font = bold_font_white
    ws.cell(row=12, column=1).fill = teal_sub_fill
    
    ws.cell(row=13, column=1, value="Se constata Pausa de Seguridad Dental realizada y registrada en Ficha Clínica.")
    ws.cell(row=13, column=1).alignment = left_aligned_text
    
    ws.cell(row=14, column=1, value="Cumple (SI/NO)").font = bold_font_navy
    ws.cell(row=14, column=1).fill = soft_teal_fill
    
    ws.cell(row=15, column=1, value="Total Cumple").font = bold_font_navy
    ws.cell(row=15, column=1).fill = soft_teal_fill

    total_cumple = 0
    total_no_cumple = 0

    for idx in range(18):
        num_pauta = idx + 1
        col_start = 2 + (idx * 2)
        col_end = col_start + 1

        pauta_data = pautas_dict.get(num_pauta) or {}

        campos = ["centro", "fecha_sup", "nombre", "apellido", "rut", "fecha_atencion", "servicio", "exodoncia"]
        for row_idx, campo in enumerate(campos, start=3):
            ws.merge_cells(start_row=row_idx, start_column=col_start, end_row=row_idx, end_column=col_end)
            ws.cell(row=row_idx, column=col_start, value=pauta_data.get(campo, ""))
            ws.cell(row=row_idx, column=col_start).alignment = center_aligned_text

        ws.merge_cells(start_row=11, start_column=col_start, end_row=11, end_column=col_end)
        ws.cell(row=11, column=col_start, value=num_pauta).alignment = center_aligned_text
        ws.cell(row=11, column=col_start).font = bold_font_navy
        ws.cell(row=11, column=col_start).fill = soft_teal_fill

        ws.cell(row=12, column=col_start, value="SI").alignment = center_aligned_text
        ws.cell(row=12, column=col_start).font = bold_font_navy
        ws.cell(row=12, column=col_end, value="NO").alignment = center_aligned_text
        ws.cell(row=12, column=col_end).font = bold_font_navy

        cumple = pauta_data.get("cumple", "")
        if cumple == "SI":
            ws.cell(row=13, column=col_start, value="√").alignment = center_aligned_text
            ws.merge_cells(start_row=14, start_column=col_start, end_row=14, end_column=col_end)
            ws.cell(row=14, column=col_start, value="SI").alignment = center_aligned_text
            total_cumple += 1
        elif cumple == "NO":
            ws.cell(row=13, column=col_end, value="X").alignment = center_aligned_text
            ws.merge_cells(start_row=14, start_column=col_start, end_row=14, end_column=col_end)
            ws.cell(row=14, column=col_start, value="NO").alignment = center_aligned_text
            total_no_cumple += 1
        else:
            ws.merge_cells(start_row=14, start_column=col_start, end_row=14, end_column=col_end)

    ws.merge_cells('B15:F15')
    ws['B15'] = total_cumple
    ws['B15'].alignment = center_aligned_text

    ws.merge_cells('G15:J15')
    ws['G15'] = "Total No Cumple"
    ws['G15'].font = bold_font_navy
    ws['G15'].alignment = center_aligned_text

    ws.merge_cells('K15:N15')
    ws['K15'] = total_no_cumple
    ws['K15'].alignment = center_aligned_text

    ws.merge_cells('O15:R15')
    ws['O15'] = "% Cumplimiento"
    ws['O15'].font = bold_font_navy
    ws['O15'].alignment = center_aligned_text

    completadas = sum(1 for v in pautas_dict.values() if v is not None)
    porcentaje = f"{(total_cumple/18)*100:.1f}%" if completadas == 18 else f"{(total_cumple/completadas)*100:.1f}%" if completadas > 0 else "-"
    ws.merge_cells('S15:V15')
    ws['S15'] = porcentaje
    ws['S15'].alignment = center_aligned_text

    ws.merge_cells('A16:Q20')
    ws['A16'] = "Observaciones:"
    ws['A16'].font = bold_font_navy
    ws['A16'].alignment = Alignment(horizontal="left", vertical="top")

    ws.merge_cells('R16:U18')

    ws.merge_cells('R19:U20')
    ws['R19'] = "Nombre o Timbre\ndel responsable de\naplicar la pauta"
    ws['R19'].font = bold_font_navy
    ws['R19'].alignment = center_aligned_text

    ws.row_dimensions[19].height = 20
    ws.row_dimensions[20].height = 20

    for r in range(1, 16):
        for c in range(1, 38):
            ws.cell(row=r, column=c).border = thin_border

    for r in range(16, 21):
        for c in range(1, 22):
            ws.cell(row=r, column=c).border = thin_border

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ==============================================================================
# --- VISTA 1: PANTALLA DE INICIO (PORTAL DE PAUTAS) ---
# ==============================================================================
if st.session_state.pagina_activa == "inicio":
    st.title("RedSalud | Portal de Pautas de Supervisión")
    st.write("Bienvenido al sistema de consolidación de pautas de supervisión clínica. Seleccione la pauta que desea evaluar:")

    st.markdown("---")

    # Tarjeta 1: Pausa de Seguridad Dental
    st.markdown("""
    <div class="card-pauta">
        <h3 style="margin-top:0; color:#00205B;">🦷 Pausa de Seguridad Dental (GCL 2.1 AO)</h3>
        <p>Evaluación de cumplimiento de Pausa de Seguridad Dental en Box Dental, Pabellón de Cirugía Menor e Imagenología Dental.</p>
        <p><b>Capacidad:</b> 18 Pautas de Supervisión por Consolidado.</p>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("🚀 Ingresar a Pauta de Seguridad Dental", type="primary", use_container_width=True):
        st.session_state.pagina_activa = "pauta_dental"
        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # Tarjeta 2: Próximas Pautas (Placeholder)
    st.markdown("""
    <div class="card-pauta-future">
        <h3 style="margin-top:0; color:#00205B;">📋 Próximas Pautas de Supervisión</h3>
        <p>Espacio reservado para incorporar nuevas pautas institucionales (Higiene de Manos, Seguridad Quirúrgica, EPP, etc.).</p>
        <p><i>Estado: En desarrollo...</i></p>
    </div>
    """, unsafe_allow_html=True)


# ==============================================================================
# --- VISTA 2: APLICACIÓN PAUSA DE SEGURIDAD DENTAL ---
# ==============================================================================
elif st.session_state.pagina_activa == "pauta_dental":
    
    # Botón para volver al inicio
    if st.button("⬅️ Volver al Portal de Inicio", use_container_width=True):
        st.session_state.pagina_activa = "inicio"
        st.rerun()

    st.markdown("---")
    st.title("RedSalud | Pausa de Seguridad Dental")

    # MÓDULO DE CARGA MASIVA
    with st.expander("📥 **Carga Masiva Mensual (Copiar y Pegar desde Excel / Archivo)**", expanded=False):
        st.write("Copia la tabla desde Excel y pégala abajo, o sube el archivo directamente:")
        
        fecha_sup_masiva = st.date_input("Fecha de Supervisión para el lote:", value=datetime.date.today())
        
        tab1, tab2 = st.tabs(["📋 Pegar Texto desde Excel", "📁 Subir Archivo Excel/CSV"])
        
        with tab1:
            texto_pegado = st.text_area("Pega la tabla copiada desde Excel aquí:", height=150, placeholder="Pega aquí las filas copiadas directamente de tu Excel...")
            if st.button("⚡ Procesar Texto Pegado", type="primary"):
                if texto_pegado.strip():
                    try:
                        df_pasted = pd.read_csv(io.StringIO(texto_pegado), sep='\t')
                        cargadas = procesar_df_masivo(df_pasted, fecha_sup_masiva)
                        st.success(f"✅ ¡Se cargaron {cargadas} pautas automáticamente!")
                        st.rerun()
                    except Exception as e:
                        st.error("Ocurrió un error al leer el texto. Asegúrate de copiar las columnas completas desde Excel.")
                else:
                    st.warning("Por favor pega algún texto antes de procesar.")
                    
        with tab2:
            archivo_subido = st.file_uploader("Selecciona archivo Excel (.xlsx) o CSV:", type=["xlsx", "csv"])
            if st.button("⚡ Procesar Archivo Subido", type="primary"):
                if archivo_subido is not None:
                    try:
                        if archivo_subido.name.endswith('.csv'):
                            df_file = pd.read_csv(archivo_subido)
                        else:
                            df_file = pd.read_excel(archivo_subido)
                        cargadas = procesar_df_masivo(df_file, fecha_sup_masiva)
                        st.success(f"✅ ¡Se cargaron {cargadas} pautas automáticamente!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Error al leer el archivo: {e}")
                else:
                    st.warning("Selecciona un archivo antes de procesar.")

    st.markdown("---")

    completadas = sum(1 for v in st.session_state.pautas_data.values() if v is not None)

    st.progress(completadas / 18)
    st.caption(f"Progreso global: **{completadas} de 18 pautas guardadas**")

    def formato_opcion(num):
        estado = "✅ Guardada" if st.session_state.pautas_data[num] is not None else "⏳ Pendiente"
        return f"Pauta N° {num} ({estado})"

    pauta_seleccionada = st.selectbox(
        "Selecciona la pauta a ingresar o revisar:",
        options=list(range(1, 19)),
        index=st.session_state.pauta_actual - 1,
        format_func=formato_opcion
    )

    st.session_state.pauta_actual = pauta_seleccionada
    p_num = st.session_state.pauta_actual

    datos_existentes = st.session_state.pautas_data[p_num] or {}

    servicios = [
        "Sala de Procedimiento Dental (BD)", 
        "Pabellón de Cirugía menor Dental (PD)", 
        "Imagenología Dental (RX)"
    ]
    idx_serv = servicios.index(datos_existentes.get('servicio')) if datos_existentes.get('servicio') in servicios else 0
    idx_exo = 0 if datos_existentes.get('exodoncia') != "NO" else 1
    idx_cumple = 0 if datos_existentes.get('cumple') != "NO" else 1

    # Formulario libre
    st.markdown('<div class="card-pauta">', unsafe_allow_html=True)
    st.subheader(f"Formulario Pauta N° {p_num}")

    centro = st.text_input("Centro", value=datos_existentes.get('centro', st.session_state.ultimo_centro), key=f"c_{p_num}")
    fecha_sup = st.date_input("Fecha de Supervisión", value=parse_fecha(datos_existentes.get('fecha_sup')), key=f"fs_{p_num}")

    nombre = st.text_input("Nombre de la persona supervisada", value=datos_existentes.get('nombre', ''), key=f"n_{p_num}", help="Solo letras permitidas")
    apellido = st.text_input("Apellido(s) de la persona supervisada", value=datos_existentes.get('apellido', ''), key=f"a_{p_num}", help="Solo letras permitidas")

    rut = st.text_input("RUT del paciente", value=datos_existentes.get('rut', ''), key=f"r_{p_num}", help="Solo números, guion y K. Ejemplo: 12345678-K")
    fecha_atencion = st.date_input("Fecha de Atención supervisada", value=parse_fecha(datos_existentes.get('fecha_atencion')), key=f"fa_{p_num}")

    servicio = st.selectbox("Servicio Clínico", servicios, index=idx_serv, key=f"s_{p_num}")
    exodoncia = st.radio("¿Procedimiento corresponde a Exodoncia?", ["SI", "NO"], index=idx_exo, key=f"e_{p_num}")

    st.markdown("---")
    st.markdown("**CRITERIO A EVALUAR**")
    cumple = st.radio("¿Se constata Pausa de Seguridad Dental realizada y registrada en Ficha Clínica?", ["SI", "NO"], index=idx_cumple, key=f"cu_{p_num}")

    btn_guardar = st.button(f"💾 Guardar Pauta N° {p_num}", type="primary", use_container_width=True, key=f"btn_{p_num}")

    if btn_guardar:
        nombre_clean = re.sub(r'[^a-zA-ZáéíóúÁÉÍÓÚñÑ\s]', '', nombre).strip().upper()
        apellido_clean = re.sub(r'[^a-zA-ZáéíóúÁÉÍÓÚñÑ\s]', '', apellido).strip().upper()
        centro_clean = centro.strip().upper()
        rut_clean = re.sub(r'[^0-9kK\-]', '', rut).strip().upper()

        st.session_state.pautas_data[p_num] = {
            "centro": centro_clean,
            "fecha_sup": fecha_sup.strftime("%d-%m-%Y"),
            "nombre": nombre_clean,
            "apellido": apellido_clean,
            "rut": rut_clean,
            "fecha_atencion": fecha_atencion.strftime("%d-%m-%Y"),
            "servicio": servicio,
            "exodoncia": exodoncia,
            "cumple": cumple
        }
        st.session_state.ultimo_centro = centro_clean
        
        if p_num < 18:
            st.session_state.pauta_actual = p_num + 1
        
        st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    # Resumen y descarga
    if completadas == 18:
        st.success("🎉 ¡Has completado las 18 pautas exitosamente!")
    else:
        st.info(f"Faltan **{18 - completadas} pautas** por completar para finalizar el proceso.")

    excel_file = generar_excel_consolidado(st.session_state.pautas_data)

    st.download_button(
        label="📥 Descargar Excel Consolidado",
        data=excel_file,
        file_name="Consolidado_Pausa_Seguridad_Dental.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    if st.button("🔄 Reiniciar todo y borrar pautas", use_container_width=True):
        st.session_state.pautas_data = {i: None for i in range(1, 19)}
        st.session_state.ultimo_centro = ""
        st.session_state.pauta_actual = 1
        st.rerun()

    pautas_list = [v for v in st.session_state.pautas_data.values() if v is not None]
    if len(pautas_list) > 0:
        with st.expander(f"📋 Ver resumen de pautas guardadas ({len(pautas_list)}/18)"):
            st.dataframe(pd.DataFrame(pautas_list), use_container_width=True)
