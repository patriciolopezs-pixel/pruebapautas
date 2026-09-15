import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
import io
import datetime
import re  # Módulo para filtrado de texto

# Configuración de la página web
st.set_page_config(page_title="Pausa de Seguridad Dental - RedSalud", layout="centered")

# --- ESTILOS CSS CON MAYÚSCULAS AUTOMÁTICAS ---
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
    
    /* Textos y Etiquetas de preguntas */
    .stMarkdown, label, p {
        color: #00205B !important;
    }

    /* Campos de entrada: Forzar visualización en MAYÚSCULAS */
    input, select, textarea, div[data-baseweb="select"] {
        background-color: #FFFFFF !important;
        color: #00205B !important;
        -webkit-text-fill-color: #00205B !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
        text-transform: uppercase !important; /* Escribe en Mayúsculas visualmente */
    }

    /* Contenedor del Formulario (Tarjeta Blanca) */
    [data-testid="stForm"] {
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-top: 6px solid #00828A !important;
        border-radius: 12px !important;
        padding: 24px !important;
        box-shadow: 0 4px 12px rgba(0, 32, 91, 0.08) !important;
    }

    /* Botón Guardar (Turquesa RedSalud) */
    button[kind="primary"] {
        background-color: #00828A !important;
        border: none !important;
        border-radius: 8px !important;
    }
    button[kind="primary"] p {
        color: #FFFFFF !important;
        font-weight: bold !important;
    }

    /* Botones Secundarios y Descargar (Azul Marino RedSalud) */
    button[kind="secondary"], .stDownloadButton button {
        background-color: #00205B !important;
        border: none !important;
        border-radius: 8px !important;
    }
    button[kind="secondary"] p, .stDownloadButton button p {
        color: #FFFFFF !important;
        font-weight: bold !important;
    }

    /* Cuadro de Resumen Desplegable */
    div[data-testid="stExpander"] {
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        border-radius: 8px !important;
    }
</style>
""", unsafe_allow_html=True)


# --- 1. ESTADO DE LA SESIÓN ---
if 'pautas_data' not in st.session_state:
    st.session_state.pautas_data = {i: None for i in range(1, 19)}

if 'pauta_actual' not in st.session_state:
    st.session_state.pauta_actual = 1

if 'ultimo_centro' not in st.session_state:
    st.session_state.ultimo_centro = ""


def parse_fecha(fecha_str):
    if fecha_str:
        try:
            return datetime.datetime.strptime(fecha_str, "%d-%m-%Y").date()
        except Exception:
            pass
    return datetime.date.today()


# --- 2. GENERADOR EXCEL CONSOLIDADO ---
def generar_excel_consolidado(pautas_dict):
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidado Pautas"

    thin_border = Border(
        left=Side(style='thin', color='B0BEC5'), 
        right=Side(style='thin', color='B0BEC5'), 
        top=Side(style='thin', color='B0BEC5'), 
        bottom=Side(style='thin', color='B0BEC5')
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
        "Centro", "Fecha de Supervisión", "Nombre de la persona supervisada", 
        "Apellido(s) de la persona supervisada", "RUT del paciente", 
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
        ws.cell(row=11, column=col_start).fill = soft_teal_fill

        ws.cell(row=12, column=col_start, value="SI").alignment = center_aligned_text
        ws.cell(row=12, column=col_end, value="NO").alignment = center_aligned_text

        cumple = pauta_data.get("cumple", "")
        if cumple == "SI":
            ws.cell(row=14, column=col_start, value="X").alignment = center_aligned_text
            total_cumple += 1
        elif cumple == "NO":
            ws.cell(row=14, column=col_end, value="X").alignment = center_aligned_text
            total_no_cumple += 1

    ws.merge_cells('B15:C15')
    ws['B15'] = total_cumple
    ws['B15'].alignment = center_aligned_text
    
    ws.merge_cells('D15:E15')
    ws['D15'] = "Total No Cumple"
    ws['D15'].font = bold_font_navy
    
    ws.merge_cells('F15:G15')
    ws['F15'] = total_no_cumple
    ws['F15'].alignment = center_aligned_text

    ws.merge_cells('H15:J15')
    ws['H15'] = "% Cumplimiento"
    ws['H15'].font = bold_font_navy

    completadas = sum(1 for v in pautas_dict.values() if v is not None)
    porcentaje = f"{(total_cumple/18)*100:.1f}%" if completadas == 18 else "-"
    ws.merge_cells('K15:L15')
    ws['K15'] = porcentaje
    ws['K15'].alignment = center_aligned_text

    for row in ws.iter_rows(min_row=1, max_row=16, min_col=1, max_col=37):
        for cell in row:
            cell.border = thin_border

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- 3. INTERFAZ DE USUARIO ---
st.title("RedSalud | Pausa de Seguridad Dental")

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

# --- FORMULARIO DE PAUTA ---
with st.form(key=f"form_pauta_numero_{p_num}"):
    st.subheader(f"Formulario Pauta N° {p_num}")
    
    centro = st.text_input("Centro", value=datos_existentes.get('centro', st.session_state.ultimo_centro))
    fecha_sup = st.date_input("Fecha de Supervisión", value=parse_fecha(datos_existentes.get('fecha_sup')))
    nombre = st.text_input("Nombre de la persona supervisada", value=datos_existentes.get('nombre', ''))
    apellido = st.text_input("Apellido(s) de la persona supervisada", value=datos_existentes.get('apellido', ''))
    rut = st.text_input("RUT del paciente", value=datos_existentes.get('rut', ''), help="Formato: 12345678-K (sin puntos)")
    fecha_atencion = st.date_input("Fecha de Atención supervisada", value=parse_fecha(datos_existentes.get('fecha_atencion')))
    servicio = st.selectbox("Servicio Clínico", servicios, index=idx_serv)
    exodoncia = st.radio("¿Procedimiento corresponde a Exodoncia?", ["SI", "NO"], index=idx_exo)

    st.markdown("---")
    st.markdown("**CRITERIO A EVALUAR**")
    cumple = st.radio("¿Se constata Pausa de Seguridad Dental realizada y registrada en Ficha Clínica?", ["SI", "NO"], index=idx_cumple)

    btn_guardar = st.form_submit_button(f"💾 Guardar Pauta N° {p_num}", type="primary", use_container_width=True)

    if btn_guardar:
        # --- LIMPIEZA Y VALIDACIÓN DE DATOS ---
        # 1. Nombre y Apellido: Solo letras, espacios, acentos y Ñ, convertidos a MAYÚSCULAS
        nombre_clean = re.sub(r'[^a-zA-ZáéíóúÁÉÍÓÚñÑ\s]', '', nombre).strip().upper()
        apellido_clean = re.sub(r'[^a-zA-ZáéíóúÁÉÍÓÚñÑ\s]', '', apellido).strip().upper()
        centro_clean = centro.strip().upper()
        
        # 2. RUT: Solo números, guion (-) y letra K/k (en mayúscula)
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

st.markdown("---")

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
