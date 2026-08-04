import streamlit as st
import pandas as pd
import tempfile
import os
import re
from PyPDF2 import PdfReader, PdfWriter
from openpyxl import load_workbook
import pdfplumber
from PIL import Image
import pytesseract

st.set_page_config(page_title="Conciliador Bancario", layout="wide")
st.title("🏦 Conciliador Bancario")

# --- Carga de archivos
pdf_file = st.file_uploader("📄 PDF del banco (Bancolombia)", type=["pdf"])
aux_file = st.file_uploader("📊 Archivo auxiliar contable (Excel)", type=["xlsx"])

if pdf_file and aux_file:
    st.info("Procesando archivos...")

    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path_original = os.path.join(tmpdir, "BANCO.pdf")
        pdf_path_unlocked = os.path.join(tmpdir, "BANCO_UNLOCKED.pdf")
        ruta_aux = os.path.join(tmpdir, "AUXILIAR.xlsx")
        ruta_final = os.path.join(tmpdir, "CONCILIADO.xlsx")
        ruta_extracto = os.path.join(tmpdir, "EXTRACTO_COMPLETO.xlsx")

        # Guardar archivos
        with open(pdf_path_original, "wb") as f:
            f.write(pdf_file.read())
        with open(ruta_aux, "wb") as f:
            f.write(aux_file.read())

        # --- Desbloquear PDF
        reader = PdfReader(pdf_path_original)
        if reader.is_encrypted:
            st.warning("El PDF está protegido con contraseña.")
            password_input = st.text_input("🔑 Ingresa la contraseña del PDF:", type="password")
            if password_input:
                try:
                    result = reader.decrypt(password_input)
                    if result == 0:
                        st.error("❌ Contraseña incorrecta.")
                        st.stop()
                    else:
                        writer = PdfWriter()
                        for page in reader.pages:
                            writer.add_page(page)
                        with open(pdf_path_unlocked, "wb") as f:
                            writer.write(f)
                        st.success("PDF desbloqueado correctamente.")
                except Exception as e:
                    st.error(f"Error al desbloquear el PDF: {e}")
                    st.stop()
            else:
                st.stop()
        else:
            pdf_path_unlocked = pdf_path_original
            st.info("El PDF no está cifrado.")

        # --- Extracción del PDF con OCR opcional
        filas = []
        with pdfplumber.open(pdf_path_unlocked) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if not texto:
                    im = page.to_image(resolution=300)
                    texto = pytesseract.image_to_string(im.original)
                lineas = texto.split("\n")

                patron = (
                    r"(\d{1,2}/\d{1,2})\s+"                # FECHA
                    r"(.*?)\s+"                            # DESCRIPCIÓN
                    r"(?:(CANAL\s+\w+|CENTRO\s+\w+|AGUAZUL|PUERTO\s+\w+)\s+)?"  # SUCURSAL opcional
                    r"(?:(\d{1,6})\s+)?"                   # DCTO. opcional
                    r"(-?\d{1,3}(?:,\d{3})*\.\d{2})\s+"    # VALOR
                    r"(-?\d{1,3}(?:,\d{3})*\.\d{2})"       # SALDO
                )

                for linea in lineas:
                    m = re.match(patron, linea)
                    if m:
                        filas.append([
                            m.group(1),
                            m.group(2).strip(),
                            m.group(3) if m.group(3) else "",
                            m.group(4) if m.group(4) else "",
                            m.group(5),
                            m.group(6)
                        ])

        df_extracto = pd.DataFrame(filas, columns=["FECHA", "DESCRIPCIÓN", "SUCURSAL", "DCTO.", "VALOR", "SALDO"])
        df_extracto["VALOR"] = df_extracto["VALOR"].str.replace(",", "").astype(float)
        df_extracto["SALDO"] = df_extracto["SALDO"].str.replace(",", "").astype(float)
        df_extracto.to_excel(ruta_extracto, index=False)

        # --- Procesar auxiliar contable
        wb = load_workbook(ruta_aux)
        ws = wb.active
        tabla = []
        for fila in ws.iter_rows(values_only=True):
            if all(c is None for c in fila):
                continue
            tabla.append(list(fila))

        df_aux = pd.DataFrame(tabla[1:], columns=tabla[0])
        df_aux.columns = [str(c).strip().lower() for c in df_aux.columns]

        # Buscar posibles nombres de columnas (más flexibles)
        col_debito = next((c for c in df_aux.columns if any(x in c for x in ["deb", "debe", "cargo"])), None)
        col_credito = next((c for c in df_aux.columns if any(x in c for x in ["cred", "haber", "abono"])), None)
        col_saldo = next((c for c in df_aux.columns if any(x in c for x in ["saldo", "balance", "importe", "valor"])), None)

        # Convertir a numérico si existen
        for col in [col_debito, col_credito, col_saldo]:
            if col:
                df_aux[col] = pd.to_numeric(df_aux[col], errors="coerce").fillna(0)

        # Calcular saldos positivos
        if col_debito and col_credito:
            df_aux["saldos_positivos"] = df_aux[col_debito].abs() + df_aux[col_credito].abs()
        elif col_saldo:
            df_aux["saldos_positivos"] = df_aux[col_saldo].abs()
        else:
            st.error("No se encontraron columnas de Débito, Crédito o Saldo (ni equivalentes) en el auxiliar.")
            st.stop()

        st.success("Columnas detectadas correctamente en el auxiliar.")

        # --- Cruce de datos
        df_extracto["FECHA_AUX"] = None
        df_extracto["DESCRIPCION_AUX"] = None
        df_extracto["DOCNUM_AUX"] = None
        df_extracto["TIPO_COINCIDENCIA"] = None
        df_aux_temp = df_aux.copy()

        conceptos_bancarios = [
            "COBRO IVA PAGOS AUTOMATICOS",
            "CUOTA PLAN CANAL NEGOCIOS",
            "IMPTO GOBIERNO 4X1000",
            "IVA CUOTA PLAN CANAL NEGOCIOS",
            "SERVICIO PAGO A PROVEEDORES",
            "SERVICIO PAGO DE NOMINA"
        ]

        for i, valor in enumerate(df_extracto["VALOR"]):
            descripcion_extracto = str(df_extracto.at[i, "DESCRIPCIÓN"]).upper()
            exacto = df_aux_temp[df_aux_temp["saldos_positivos"] == abs(valor)]
            if not exacto.empty:
                fila = exacto.iloc[0]
                df_extracto.at[i, "FECHA_AUX"] = fila.get("fecha")
                df_extracto.at[i, "DESCRIPCION_AUX"] = fila.get("nota")
                df_extracto.at[i, "DOCNUM_AUX"] = fila.get("doc num")
                df_extracto.at[i, "TIPO_COINCIDENCIA"] = "COINCIDENCIA EXACTA"
                df_aux_temp = df_aux_temp.drop(fila.name)
            else:
                df_extracto.at[i, "TIPO_COINCIDENCIA"] = "NO EXISTE"

            for concepto in conceptos_bancarios:
                if concepto in descripcion_extracto:
                    df_extracto.at[i, "DESCRIPCION_AUX"] = "GASTOS BANCARIOS"
                    df_extracto.at[i, "TIPO_COINCIDENCIA"] = "GASTOS BANCARIOS"
                    break

        df_extracto.to_excel(ruta_final, index=False)
        st.success("✅ Archivo CONCILIADO.xlsx generado correctamente.")

        with open(ruta_final, "rb") as f:
            conciliado_bytes = f.read()

        st.download_button(
            label="⬇️ Descargar archivo conciliado",
            data=conciliado_bytes,
            file_name="CONCILIADO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
