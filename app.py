import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
import io
import datetime

# Configuración de la página web
st.set_page_config(page_title="Pautas de Supervisión Dental", layout="centered")

# --- 1. INICIALIZACIÓN DE ESTADO ---
if 'pautas' not in st.session_state:
    st.session_state.pautas = []
if 'ultimo_centro' not in st.session_state:
    st.session_state.ultimo_centro = ""
if 'mensaje_exito' not in st.session_state:
    st.session_state.mensaje_exito = ""

# Valores por defecto para los campos
if 'input_centro' not in st.session_state:
    st.session_state.input_centro = st.session_state.ultimo_centro
if 'input_nombre' not in st.session_state:
    st.session_state.input_nombre = ""
if 'input_apellido' not in st.session_state:
    st.session_state.input_apellido = ""
if 'input_rut' not in st.session_state:
    st.session_state.input_rut = ""


# --- 2. FUNCIONES CALLBACK (PROCESAMIENTO DIRECTO) ---
def guardar_pauta():
    # Registrar pauta actual
    nueva_pauta = {
        "centro": st.session_state.input_centro,
        "fecha_sup": st.session_state.input_fecha_sup.strftime("%d-%m-%Y"),
        "nombre": st.session_state.input_nombre,
        "apellido": st.session_state.input_apellido,
        "rut": st.session_state.input_rut,
        "fecha_atencion": st.session_state.input_fecha_atencion.strftime("%d-%m-%Y"),
        "servicio": st.session_state.input_servicio,
        "exodoncia": st.session_state.input_exodoncia,
        "cumple": st.session_state.input_cumple
    }
    st.session_state.pautas.append(nueva_pauta)
    
    # Mantener el centro ingresado
    st.session_state.ultimo_centro = st.session_state.input_centro
    
    # Confirmación
    num = len(st.session_state.pautas)
    st.session_state.mensaje_exito = f"✅ ¡Pauta N° {num} guardada con éxito!"

    # Limpiar campos personales para la siguiente pauta
    st.session_state.input_nombre = ""
    st.session_state.input_apellido = ""
    st.session_state.input_rut = ""

def borrar_ultima():
    if len(st.session_state.pautas) > 0:
        st.session_state.pautas.pop()
        st.session_state.mensaje_exito = "↩️ Se eliminó la última pauta registrada."

def reiniciar_todo():
    st.session_state.pautas = []
    st.session_state.ultimo_centro = ""
    st.session_state.mensaje_exito = ""
    st.session_state.input_centro = ""
    st.session_state.input_nombre = ""
    st.session_state.input_apellido = ""
    st.session_state.input_rut = ""


# --- 3. GENERADOR DE EXCEL CONSOLIDADO ---
def generar_excel_consolidado(pautas):
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidado Pautas"

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center_aligned_text = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_aligned_text = Alignment(horizontal="left", vertical="center", wrap_text=True)
    bold_font = Font(bold=True)
    blue_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    ws.merge_cells('A1:AK1')
    ws['A1'] = "PAUTA DE SUPERVISIÓN CUMPLIMIENTO DE PAUSA DE SEGURIDAD DENTAL EN BOX DENTAL, PABELLÓN DE CIRUGÍA MENOR DENTAL E IMAGENOLOGÍA DENTAL (GCL 2.1 AO)"
    ws['A1'].font = bold_font
    ws['A1'].alignment = center_aligned_text

    ws['A2'] = "Indicaciones llenado pauta"
    ws.merge_cells('B2:AK2')
    ws['B2'] = "Marque √ SI cumple, Marque X NO cumple, o No Aplica (si corresponde). En ítem cumple registre SI o NO. Registre en observaciones motivo incumplimiento."
    ws['A2'].font = bold_font
    ws['B2'].alignment = center_aligned_text

    etiquetas = [
        "Centro", "Fecha de Supervisión", "Nombre de la persona supervisada", 
        "Apellido(s) de la persona supervisada", "RUT del paciente", 
        "Fecha de Atención supervisada", 
        "Servicio Clínico donde se realizó el procedimiento, ya sea Sala de Procedimiento Dental (BD), Pabellón de Cirugía menor Dental (PD), e Imagenología Dental (RX)",
        "Procedimiento corresponde a Exodoncia (SI, NO)"
    ]

    for i, etiqueta in enumerate(etiquetas, start=3):
        ws.cell(row=i, column=1, value=etiqueta).font = bold_font
        ws.cell(row=i, column=1).alignment = left_aligned_text
        ws.cell(row=i, column=1).fill = blue_fill

    ws.column_dimensions['A'].width = 50

    ws.cell(row=11, column=1, value="N° DE PAUTA").font = bold_font
    ws.cell(row=12, column=1, value="CRITERIOS A EVALUAR").font = bold_font
    ws.cell(row=13, column=1, value="Se constata Pausa de Seguridad Dental realizada y registrada en Ficha Clínica.")
    ws.cell(row=14, column=1, value="Cumple (SI/NO)").font = bold_font
    ws.cell(row=15, column=1, value="Total Cumple").font = bold_font
    ws.cell(row=11, column=1).fill = blue_fill
    ws.cell(row=12, column=1).fill = blue_fill
    ws.cell(row=14, column=1).fill = blue_fill
    ws.cell(row=15, column=1).fill = blue_fill

    total_cumple = 0
    total_no_cumple = 0

    for idx in range(18):
        col_start = 2 + (idx * 2)
        col_end = col_start + 1

        pauta_data = pautas[idx] if idx < len(pautas) else {}

        campos = ["centro", "fecha_sup", "nombre", "apellido", "rut", "fecha_atencion", "servicio", "exodoncia"]
        for row_idx, campo in enumerate(campos, start=3):
            ws.merge_cells(start_row=row_idx, start_column=col_start, end_row=row_idx, end_column=col_end)
            ws.cell(row=row_idx, column=col_start, value=pauta_data.get(campo, ""))
            ws.cell(row=row_idx, column=col_start).alignment = center_aligned_text

        ws.merge_cells(start_row=11, start_column=col_start, end_row=11, end_column=col_end)
        ws.cell(row=11, column=col_start, value=idx + 1).alignment = center_aligned_text

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
    ws['D15'].font = bold_font
    
    ws.merge_cells('F15:G15')
    ws['F15'] = total_no_cumple
    ws['F15'].alignment = center_aligned_text

    ws.merge_cells('H15:J15')
    ws['H15'] = "% Cumplimiento"
    ws['H15'].font = bold_font

    porcentaje = f"{(total_cumple/18)*100:.1f}%" if len(pautas) == 18 else "-"
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


# --- 4. INTERFAZ DE USUARIO ---
st.title("🦷 Pausa de Seguridad Dental")

pautas_ingresadas = len(st.session_state.pautas)

st.progress(pautas_ingresadas / 18)
st.caption(f"Pautas guardadas: **{pautas_ingresadas} de 18**")

if st.session_state.mensaje_exito:
    st.success(st.session_state.mensaje_exito)

if pautas_ingresadas < 18:
    st.subheader(f"Ingresando Pauta N° {pautas_ingresadas + 1}")
    
    # Entradas de texto enlazadas directamente al estado
    st.text_input("Centro", key="input_centro")
    st.date_input("Fecha de Supervisión", key="input_fecha_sup")
    st.text_input("Nombre de la persona supervisada", key="input_nombre")
    st.text_input("Apellido(s) de la persona supervisada", key="input_apellido")
    st.text_input("RUT del paciente", key="input_rut")
    st.date_input("Fecha de Atención supervisada", key="input_fecha_atencion")
    st.selectbox("Servicio Clínico", [
        "Sala de Procedimiento Dental (BD)", 
        "Pabellón de Cirugía menor Dental (PD)", 
        "Imagenología Dental (RX)"
    ], key="input_servicio")
    st.radio("¿Procedimiento corresponde a Exodoncia?", ["SI", "NO"], key="input_exodoncia")

    st.markdown("---")
    st.markdown("**CRITERIO A EVALUAR**")
    st.radio("¿Se constata Pausa de Seguridad Dental realizada y registrada en Ficha Clínica?", ["SI", "NO"], key="input_cumple")

    # Botón directo con ejecución de callback
    st.button(
        f"💾 Guardar Pauta N° {pautas_ingresadas + 1}", 
        type="primary", 
        use_container_width=True,
        on_click=guardar_pauta
    )

    if pautas_ingresadas > 0:
        st.button("⏪ Borrar última pauta ingresada", use_container_width=True, on_click=borrar_ultima)

else:
    st.success("🎉 ¡Has completado exitosamente las 18 pautas de supervisión!")
    
    excel_file = generar_excel_consolidado(st.session_state.pautas)
    
    st.download_button(
        label="📥 Descargar Excel Consolidado (18 Pautas)",
        data=excel_file,
        file_name="Consolidado_Pausa_Seguridad_Dental.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.button("🔄 Reiniciar y crear nuevo consolidado", use_container_width=True, on_click=reiniciar_todo)

# Vista previa
if pautas_ingresadas > 0:
    with st.expander(f"📋 Ver lista de pautas ingresadas ({pautas_ingresadas}/18)"):
        st.dataframe(pd.DataFrame(st.session_state.pautas), use_container_width=True)
