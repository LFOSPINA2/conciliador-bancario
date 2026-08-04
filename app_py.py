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
                    r"(\d{1,2}/\d{1,2})\s+"                
                    r"(.*?)\s+"                            
                    r"(?:(CANAL\s+\w+|CENTRO\s+\w+|AGUAZUL|PUERTO\s+\w+)\s+)?"  
                    r"(?:(\d{1,6})\s+)?"                   
                    r"(-?\d{1,3}(?:,\d{3})*\.\d{2})\s+"    
                    r"(-?\d{1,3}(?:,\d{3})*\.\d{2})"       
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

        # --- Limpieza avanzada del auxiliar contable
        wb = load_workbook(ruta_aux)
        ws = wb.active
        tabla = []
        for fila in ws.iter_rows(values_only=True):
            if all(c is None for c in fila):
                continue
            tabla.append(list(fila))

        encabezado_idx = None
        for i, fila in enumerate(tabla):
            fila_lower = [str(c).lower().strip() for c in fila]
            if any(x in fila_lower for x in ["debito", "debitos", "crédito", "creditos", "saldo"]):
                encabezado_idx = i
                break

        if encabezado_idx is None:
            encabezado_idx = 0

        df_aux = pd.DataFrame(tabla[encabezado_idx + 1:], columns=tabla[encabezado_idx])
        df_aux.columns = [str(c).lower().strip() for c in df_aux.columns]

        df_aux = df_aux.dropna(how="all")
        df_aux = df_aux[df_aux.apply(lambda x: any(pd.notna(x)), axis=1)]

        col_debito = next((c for c in df_aux.columns if any(x in c for x in ["deb", "debe", "cargo"])), None)
        col_credito = next((c for c in df_aux.columns if any(x in c for x in ["cred", "haber", "abono"])), None)
        col_saldo = next((c for c in df_aux.columns if any(x in c for x in ["saldo", "balance", "importe", "valor"])), None)

        for col in [col_debito, col_credito, col_saldo]:
            if col:
                df_aux[col] = pd.to_numeric(df_aux[col], errors="coerce").fillna(0)

        if col_debito and col_credito:
            df_aux["saldos_positivos"] = df_aux[col_debito].abs() + df_aux[col_credito].abs()
        elif col_saldo:
            df_aux["saldos_positivos"] = df_aux[col_saldo].abs()
        else:
            st.error("No se encontraron columnas de Débito, Crédito o Saldo.")
            st.stop()

        st.success("Columnas detectadas y base auxiliar limpiada correctamente.")

        # --- Cruce corregido con tolerancia
        st.info("Iniciando cruce de datos entre extracto y auxiliar...")

        df_extracto["FECHA_AUX"] = None
        df_extracto["DESCRIPCION_AUX"] = None
        df_extracto["DOCNUM_AUX"] = None
        df_extracto["TIPO_COINCIDENCIA"] = None
        df_aux_temp = df_aux.copy()

        # Lista ampliada de gastos bancarios
        conceptos_bancarios = [
            "COBRO IVA PAGOS AUTOMATICOS",
            "CUOTA PLAN",
            "CUOTA PLAN CANAL NEGOCIOS",
            "IVA CUOTA PLAN",
            "IVA CUOTA PLAN CANAL NEGOCIOS",
            "IMPTO GOBIERNO 4X1000",
            "SERVICIO PAGO A PROVEEDORES",
            "SERVICIO PAGO DE NOMINA",
            "COMISION",
            "COMISIONES",
            "GASTOS BANCARIOS"
        ]

        progress_bar = st.progress(0)
        total = len(df_extracto)

        for i, valor in enumerate(df_extracto["VALOR"]):
            descripcion_extracto = str(df_extracto.at[i, "DESCRIPCIÓN"]).upper()

            # Coincidencia exacta
            exacto = df_aux_temp[df_aux_temp["saldos_positivos"] == abs(valor)]

            # Coincidencia por tolerancia ±1 peso
            tolerancia = df_aux_temp[(df_aux_temp["saldos_positivos"] >= abs(valor) - 1) &
                                     (df_aux_temp["saldos_positivos"] <= abs(valor) + 1)]

            # Coincidencia por diferencia mínima flotante
            dif_min = df_aux_temp[abs(df_aux_temp["saldos_positivos"] - abs(valor)) < 0.01]

            fila = None
            tipo = None

            if not exacto.empty:
                fila = exacto.iloc[0]
                tipo = "COINCIDENCIA EXACTA"
            elif not tolerancia.empty:
                fila = tolerancia.iloc[0]
                tipo = "TOLERANCIA ±1"
            elif not dif_min.empty:
                fila = dif_min.iloc[0]
                tipo = "REDONDEO"

            if fila is not None:
                df_extracto.at[i, "FECHA_AUX"] = fila.get("fecha")
                df_extracto.at[i, "DESCRIPCION_AUX"] = fila.get("nota")
                df_extracto.at[i, "DOCNUM_AUX"] = fila.get("doc num")
                df_extracto.at[i, "TIPO_COINCIDENCIA"] = tipo
                df_aux_temp = df_aux_temp.drop(fila.name)
            else:
                df_extracto.at[i, "TIPO_COINCIDENCIA"] = "NO EXISTE"

            # Clasificación de gastos bancarios
            for concepto in conceptos_bancarios:
                if concepto in descripcion_extracto:
                    df_extracto.at[i, "DESCRIPCION_AUX"] = "GASTOS BANCARIOS"
                    df_extracto.at[i, "TIPO_COINCIDENCIA"] = "GASTOS BANCARIOS"
                    break

            progress_bar.progress((i + 1) / total)

        st.success("Cruce de datos completado correctamente.")

        # --- Generar archivo final
        df_extracto.to_excel(ruta_final, index=False)

        with open(ruta_final, "rb") as f:
            conciliado_bytes = f.read()

        st.download_button(
            label="⬇️ Descargar archivo conciliado",
            data=conciliado_bytes,
            file_name="CONCILIADO.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
