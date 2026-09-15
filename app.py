import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
import io

# Configuración de la página web
st.set_page_config(page_title="Pautas de Supervisión Dental", layout="wide")

# Inicializar variables en sesión para guardar las 18 pautas
if 'pautas' not in st.session_state:
    st.session_state.pautas = []

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
    blue_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # Color similar al de la imagen

    # 1. Títulos y Encabezados
    ws.merge_cells('A1:AK1')
    ws['A1'] = "PAUTA DE SUPERVISIÓN CUMPLIMIENTO DE PAUSA DE SEGURIDAD DENTAL EN BOX DENTAL, PABELLÓN DE CIRUGÍA MENOR DENTAL E IMAGENOLOGÍA DENTAL (GCL 2.1 AO)"
    ws['A1'].font = bold_font
    ws['A1'].alignment = center_aligned_text

    ws['A2'] = "Indicaciones llenado pauta"
    ws.merge_cells('B2:AK2')
    ws['B2'] = "Marque √ SI cumple, Marque X NO cumple, o No Aplica (si corresponde). En ítem cumple registre SI o NO. Registre en observaciones motivo incumplimiento."
    ws['A2'].font = bold_font
    ws['B2'].alignment = center_aligned_text

    # 2. Etiquetas de las filas (Columna A)
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

    # Ancho de la columna A para que quepa el texto
    ws.column_dimensions['A'].width = 50

    # 3. Filas de Criterios (Columna A)
    ws.cell(row=11, column=1, value="N° DE PAUTA").font = bold_font
    ws.cell(row=12, column=1, value="CRITERIOS A EVALUAR").font = bold_font
    ws.cell(row=13, column=1, value="Se constata Pausa de Seguridad Dental realizada y registrada en Ficha Clínica.")
    ws.cell(row=14, column=1, value="Cumple (SI/NO)").font = bold_font
    ws.cell(row=15, column=1, value="Total Cumple").font = bold_font
    ws.cell(row=11, column=1).fill = blue_fill
    ws.cell(row=12, column=1).fill = blue_fill
    ws.cell(row=14, column=1).fill = blue_fill
    ws.cell(row=15, column=1).fill = blue_fill

    # 4. Poblar datos de las 18 Pautas
    total_cumple = 0
    total_no_cumple = 0

    for idx in range(18):
        col_start = 2 + (idx * 2) # Pauta 1 empieza en B(2), Pauta 2 en D(4), etc.
        col_end = col_start + 1

        # Combinar celdas superiores e insertar datos (si la pauta fue ingresada)
        pauta_data = pautas[idx] if idx < len(pautas) else {}

        campos = ["centro", "fecha_sup", "nombre", "apellido", "rut", "fecha_atencion", "servicio", "exodoncia"]
        for row_idx, campo in enumerate(campos, start=3):
            ws.merge_cells(start_row=row_idx, start_column=col_start, end_row=row_idx, end_column=col_end)
            ws.cell(row=row_idx, column=col_start, value=pauta_data.get(campo, ""))
            ws.cell(row=row_idx, column=col_start).alignment = center_aligned_text

        # N° de Pauta
        ws.merge_cells(start_row=11, start_column=col_start, end_row=11, end_column=col_end)
        ws.cell(row=11, column=col_start, value=idx + 1).alignment = center_aligned_text

        # SI / NO
        ws.cell(row=12, column=col_start, value="SI").alignment = center_aligned_text
        ws.cell(row=12, column=col_end, value="NO").alignment = center_aligned_text

        # Marcado de Cumplimiento
        cumple = pauta_data.get("cumple", "")
        if cumple == "SI":
            ws.cell(row=14, column=col_start, value="X").alignment = center_aligned_text
            total_cumple += 1
        elif cumple == "NO":
            ws.cell(row=14, column=col_end, value="X").alignment = center_aligned_text
            total_no_cumple += 1

    # 5. Totales e indicadores finales (emulando la fila 15 de la imagen)
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

    # Aplicar bordes a todo el rango utilizado
    for row in ws.iter_rows(min_row=1, max_row=16, min_col=1, max_col=37):
        for cell in row:
            cell.border = thin_border

    # Guardar a memoria para la descarga
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- INTERFAZ WEB DE STREAMLIT ---
st.title("🦷 Plataforma de Consolidación: Pausa de Seguridad Dental")
st.write("Complete el formulario para las 18 pautas. Al finalizar, podrá descargar el Excel consolidado.")

pautas_ingresadas = len(st.session_state.pautas)

# Mostrar progreso
st.progress(pautas_ingresadas / 18)
st.write(f"**Pautas ingresadas:** {pautas_ingresadas} de 18")

if pautas_ingresadas < 18:
    st.subheader(f"Ingresando Pauta N° {pautas_ingresadas + 1}")
    
    with st.form("form_pauta", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            centro = st.text_input("Centro")
            nombre = st.text_input("Nombre de la persona supervisada")
            rut = st.text_input("RUT del paciente")
            servicio = st.selectbox("Servicio Clínico", [
                "Sala de Procedimiento Dental (BD)", 
                "Pabellón de Cirugía menor Dental (PD)", 
                "Imagenología Dental (RX)"
            ])
        
        with col2:
            fecha_sup = st.date_input("Fecha de Supervisión")
            apellido = st.text_input("Apellido(s) de la persona supervisada")
            fecha_atencion = st.date_input("Fecha de Atención supervisada")
            exodoncia = st.radio("Procedimiento corresponde a Exodoncia", ["SI", "NO"], horizontal=True)

        st.markdown("---")
        st.markdown("### CRITERIOS A EVALUAR")
        cumple = st.radio("¿Se constata Pausa de Seguridad Dental realizada y registrada en Ficha Clínica?", ["SI", "NO"], horizontal=True)

        btn_guardar = st.form_submit_button("Guardar Pauta", type="primary")

        if btn_guardar:
            # Guardamos los datos en el estado de la sesión
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
            st.rerun() # Recarga la página para mostrar la siguiente pauta
else:
    st.success("✅ ¡Se han ingresado las 18 pautas satisfactoriamente!")
    
    excel_file = generar_excel_consolidado(st.session_state.pautas)
    
    st.download_button(
        label="📥 Descargar Documento Consolidado (Excel)",
        data=excel_file,
        file_name="Consolidado_Pausa_Seguridad_Dental.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if st.button("Reiniciar / Crear un nuevo consolidado"):
        st.session_state.pautas = []
        st.rerun()
