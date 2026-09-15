import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
import io

# Configuración de la página web
st.set_page_config(page_title="Pautas de Supervisión Dental", layout="centered")

# Inicializar variables de sesión
if 'pautas' not in st.session_state:
    st.session_state.pautas = []
if 'ultimo_centro' not in st.session_state:
    st.session_state.ultimo_centro = ""
if 'mensaje_exito' not in st.session_state:
    st.session_state.mensaje_exito = ""

def generar_excel_consolidado(pautas):
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidado Pautas"

    # Estilos
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    center_aligned_text = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_aligned_text = Alignment(horizontal="left", vertical="center", wrap_text=True)
    bold_font = Font(bold=True)
    blue_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    # 1. Encabezados
    ws.merge_cells('A1:AK1')
    ws['A1'] = "PAUTA DE SUPERVISIÓN CUMPLIMIENTO DE PAUSA DE SEGURIDAD DENTAL EN BOX DENTAL, PABELLÓN DE CIRUGÍA MENOR DENTAL E IMAGENOLOGÍA DENTAL (GCL 2.1 AO)"
    ws['A1'].font = bold_font
    ws['A1'].alignment = center_aligned_text

    ws['A2'] = "Indicaciones llenado pauta"
    ws.merge_cells('B2:AK2')
    ws['B2'] = "Marque √ SI cumple, Marque X NO cumple, o No Aplica (si corresponde). En ítem cumple registre SI o NO. Registre en observaciones motivo incumplimiento."
    ws['A2'].font = bold_font
    ws['B2'].alignment = center_aligned_text

    # 2. Filas de datos
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

    # 3. Criterios
    ws.cell(row=11, column=1, value="N° DE PAUTA").font = bold_font
    ws.cell(row=12, column=1, value="CRITERIOS A EVALUAR").font = bold_font
    ws.cell(row=13, column=1, value="Se constata Pausa de Seguridad Dental realizada y registrada en Ficha Clínica.")
    ws.cell(row=14, column=1, value="Cumple (SI/NO)").font = bold_font
    ws.cell(row=15, column=1, value="Total Cumple").font = bold_font
    ws.cell(row=11, column=1).fill = blue_fill
    ws.cell(row=12, column=1).fill = blue_fill
    ws.cell(row=14, column=1).fill = blue_fill
    ws.cell(row=15, column=1).fill = blue_fill

    # 4. Cargar datos de las 18 pautas
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

    # 5. Totales e Indicadores
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


# --- INTERFAZ DE APLICACIÓN ---
st.title("🦷 Pausa de Seguridad Dental")

pautas_ingresadas = len(st.session_state.pautas)

# Barra de avance
st.progress(pautas_ingresadas / 18)
st.caption(f"Pautas guardadas: **{pautas_ingresadas} de 18**")

# Aviso de confirmación
if st.session_state.mensaje_exito:
    st.success(st.session_state.mensaje_exito)

if pautas_ingresadas < 18:
    i = pautas_ingresadas
    st.subheader(f"Ingresando Pauta N° {i + 1}")
    
    # Formulario con llaves (keys) únicas por cada paso
    with st.form(f"form_pauta_step_{i}"):
        centro = st.text_input("Centro", value=st.session_state.ultimo_centro, key=f"centro_{i}")
        fecha_sup = st.date_input("Fecha de Supervisión", key=f"fecha_sup_{i}")
        nombre = st.text_input("Nombre de la persona supervisada", key=f"nombre_{i}")
        apellido = st.text_input("Apellido(s) de la persona supervisada", key=f"apellido_{i}")
        rut = st.text_input("RUT del paciente", key=f"rut_{i}")
        fecha_atencion = st.date_input("Fecha de Atención supervisada", key=f"fecha_atencion_{i}")
        servicio = st.selectbox("Servicio Clínico", [
            "Sala de Procedimiento Dental (BD)", 
            "Pabellón de Cirugía menor Dental (PD)", 
            "Imagenología Dental (RX)"
        ], key=f"servicio_{i}")
        exodoncia = st.radio("¿Procedimiento corresponde a Exodoncia?", ["SI", "NO"], key=f"exodoncia_{i}")

        st.markdown("---")
        st.markdown("**CRITERIO A EVALUAR**")
        cumple = st.radio("¿Se constata Pausa de Seguridad Dental realizada y registrada en Ficha Clínica?", ["SI", "NO"], key=f"cumple_{i}")

        btn_guardar = st.form_submit_button(f"Guardar Pauta N° {i + 1}", type="primary", use_container_width=True)

        if btn_guardar:
            st.session_state.pautas.append({
                "centro": centro,
                "fecha_sup": fecha_sup.strftime("%d-%m-%Y"),
                "nombre": nombre,
                "apellido": apellido,
                "rut": rut,
                "fecha_atencion": fecha_atencion.strftime("%d-%m-%Y"),
                "servicio": servicio,
                "exodoncia": exodoncia,
                "cumple": cumple
            })
            st.session_state.ultimo_centro = centro
            st.session_state.mensaje_exito = f"✅ ¡Pauta N° {i + 1} guardada con éxito! Ahora estás completando la Pauta N° {i + 2}."
            st.rerun()

    # Opción para deshacer el último ingreso
    if pautas_ingresadas > 0:
        if st.button("⏪ Borrar última pauta ingresada", use_container_width=True):
            st.session_state.pautas.pop()
            st.session_state.mensaje_exito = "↩️ Se eliminó la última pauta registrada."
            st.rerun()

else:
    st.success("🎉 ¡Has completado las 18 pautas de supervisión!")
    
    excel_file = generar_excel_consolidado(st.session_state.pautas)
    
    st.download_button(
        label="📥 Descargar Excel Consolidado (18 Pautas)",
        data=excel_file,
        file_name="Consolidado_Pausa_Seguridad_Dental.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    if st.button("Reiniciar y crear nuevo consolidado", use_container_width=True):
        st.session_state.pautas = []
        st.session_state.ultimo_centro = ""
        st.session_state.mensaje_exito = ""
        st.rerun()

# Vista previa desplegable
if pautas_ingresadas > 0:
    with st.expander(f"📋 Ver pautas ingresadas ({pautas_ingresadas}/18)"):
        st.dataframe(pd.DataFrame(st.session_state.pautas), use_container_width=True)
