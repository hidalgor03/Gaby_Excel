import streamlit as st
import pandas as pd
import uuid
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo  # ✅ NUEVO

st.set_page_config(page_title="Generador de Acta Corporativa", layout="wide")

st.title("📄 Generador de Acta Corporativa")

uploaded_file = st.file_uploader("Subir Excel Forms", type=["xlsx"])
logo_file = st.file_uploader("Subir Logo Corporativo", type=["png", "jpg", "jpeg"])
firma_file = st.file_uploader("Subir Firma Instructor (PNG)", type=["png"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    st.subheader("Vista previa")
    st.dataframe(df.head())

    # 🔹 Sidebar configuración
    st.sidebar.header("Configuración General")
    curso = st.sidebar.text_input("Nombre del Curso")
    instructor = st.sidebar.text_input("Nombre del Instructor")
    puntaje_maximo = st.sidebar.number_input("Puntaje Máximo", min_value=1.0, value=10.0)
    puntaje_minimo = st.sidebar.number_input("Porcentaje mínimo (%)", min_value=0.0, max_value=100.0, value=100.0)

    st.sidebar.header("Tamaño Logo")
    logo_width = st.sidebar.number_input("Ancho Logo (pulgadas)", min_value=0.5, value=2.5)
    logo_height = st.sidebar.number_input("Alto Logo (pulgadas)", min_value=0.5, value=2.0)

    st.sidebar.header("Tamaño Firma")
    firma_width = st.sidebar.number_input("Ancho Firma (pulgadas)", min_value=0.5, value=2.5)
    firma_height = st.sidebar.number_input("Alto Firma (pulgadas)", min_value=0.5, value=1.25)

    if st.button("Generar Acta PDF"):

        required_columns = ["Nombre", "Hora de finalización", "Total de puntos"]

        if not all(col in df.columns for col in required_columns):
            st.error("El Excel no contiene las columnas necesarias.")
        else:
            try:

                codigo_acta = str(uuid.uuid4())[:8].upper()

                df_resultado = df[required_columns].copy()

                df_resultado["Total de puntos"] = (
                    df_resultado["Total de puntos"]
                    .astype(str)
                    .str.replace(",", ".")
                    .str.replace("%", "")
                    .str.strip()
                )

                df_resultado["Total de puntos"] = pd.to_numeric(
                    df_resultado["Total de puntos"],
                    errors="coerce"
                )

                df_resultado = df_resultado.dropna(subset=["Total de puntos"])

                df_resultado["Porcentaje"] = (
                    df_resultado["Total de puntos"] / puntaje_maximo
                ) * 100

                df_resultado["Porcentaje"] = df_resultado["Porcentaje"].round(2)

                df_resultado["Fecha_dt"] = pd.to_datetime(
                    df_resultado["Hora de finalización"],
                    dayfirst=True,
                    errors="coerce"
                )

                df_resultado["Fecha"] = df_resultado["Fecha_dt"].dt.strftime("%d-%m-%Y")

                df_resultado["Resultado"] = df_resultado["Porcentaje"].apply(
                    lambda x: "Aprobó" if x >= puntaje_minimo else "No Aprobó"
                )

                df_resultado["Vigencia"] = df_resultado.apply(
                    lambda row: (
                        (row["Fecha_dt"] + pd.DateOffset(years=2)).strftime("%d-%m-%Y")
                        if row["Resultado"] == "Aprobó"
                        else "—"
                    ),
                    axis=1
                )

                df_final = df_resultado[
                    ["Fecha", "Nombre", "Total de puntos", "Porcentaje", "Resultado", "Vigencia"]
                ]

                # -------- PDF --------
                buffer = BytesIO()
                doc = SimpleDocTemplate(buffer, pagesize=A4)
                elements = []
                styles = getSampleStyleSheet()

                # Logo
                if logo_file:
                    logo = Image(
                        logo_file,
                        width=logo_width * inch,
                        height=logo_height * inch
                    )
                    elements.append(logo)
                    elements.append(Spacer(1, 0.3 * inch))

                elements.append(Paragraph("<b>ACTA DE EVALUACIÓN</b>", styles["Title"]))
                elements.append(Spacer(1, 0.2 * inch))
                elements.append(Paragraph(f"Código Acta: {codigo_acta}", styles["Normal"]))
                elements.append(Paragraph(f"Curso: {curso}", styles["Normal"]))
                elements.append(Paragraph(f"Instructor: {instructor}", styles["Normal"]))
                elements.append(Paragraph(f"Criterio aprobación: ≥ {puntaje_minimo}%", styles["Normal"]))

                # ✅ FECHA CON ZONA HORARIA DE CHILE
                fecha_chile = datetime.now(ZoneInfo("America/Santiago")).strftime("%d-%m-%Y %H:%M")
                elements.append(Paragraph(f"Fecha generación: {fecha_chile}", styles["Normal"]))

                elements.append(Spacer(1, 0.5 * inch))

                data = [df_final.columns.tolist()] + df_final.values.tolist()
                table = Table(data, repeatRows=1)

                style = TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                    ('ALIGN', (2, 1), (-1, -1), 'CENTER'),
                ])

                for i, row in enumerate(df_final["Resultado"], start=1):
                    if row == "Aprobó":
                        style.add('BACKGROUND', (4, i), (4, i), colors.lightgreen)
                    else:
                        style.add('BACKGROUND', (4, i), (4, i), colors.salmon)

                table.setStyle(style)
                elements.append(table)

                elements.append(Spacer(1, 0.1 * inch))

                # Firma
                if firma_file:
                    elements.append(Paragraph("Firma Instructor:", styles["Normal"]))
                    firma = Image(
                        firma_file,
                        width=firma_width * inch,
                        height=firma_height * inch
                    )
                    elements.append(firma)
                else:
                    elements.append(Paragraph("Firma Instructor: ___________________________", styles["Normal"]))

                doc.build(elements)
                buffer.seek(0)

                st.success(f"Acta generada correctamente | Código: {codigo_acta}")

                st.download_button(
                    label="Descargar Acta PDF",
                    data=buffer,
                    file_name=f"ACTA_{codigo_acta}.pdf",
                    mime="application/pdf"
                )

            except Exception as e:
                st.error(f"Error al procesar: {e}")