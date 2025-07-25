import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
from io import BytesIO

st.set_page_config(page_title="Rekening Koran Mandiri Extractor", layout="wide")

st.title("📄 Extractor Rekening Koran Mandiri ke Excel")

uploaded_file = st.file_uploader("Upload PDF Rekening Koran Mandiri", type="pdf")

if uploaded_file:
    pdf_bytes = uploaded_file.read()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    text = "\n".join(page.get_text() for page in doc)

    norekening = re.search(r'Account No\.\s*(\d+)', text)
    currency = re.search(r'Currency\s+([A-Z]+)', text)
    no_rekening = norekening.group(1) if norekening else "-"
    mata_uang = currency.group(1) if currency else "-"

    transaksi_pattern = re.findall(
        r"(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}:\d{2}(.*?)(?=\d{2}/\d{2}/\d{4}|\Z)",
        text,
        re.DOTALL,
    )

    rows = []
    for tgl, blok in transaksi_pattern:
        lines = [line.strip() for line in blok.strip().splitlines() if line.strip()]
        if len(lines) < 2:
            continue

        keterangan = " ".join(lines[:-1])
        angka = re.findall(r"[-\d,.]+", lines[-1])
        if len(angka) < 3:
            continue

        debit, kredit, saldo = angka[-3:]

        def parse_amount(val):
            val = val.replace(",", "").replace(".", "")
            return float(val) if val not in ["-", "", None] else 0.0

        rows.append({
            "Nomor Rekening": no_rekening,
            "Tanggal": pd.to_datetime(tgl, format="%d/%m/%Y").strftime("%d/%m/%Y"),
            "Keterangan": keterangan,
            "Debit": parse_amount(debit) if debit != "0.00" else 0.0,
            "Kredit": parse_amount(kredit) if kredit != "0.00" else 0.0,
            "Saldo": parse_amount(saldo),
            "Currency": mata_uang,
        })

    df = pd.DataFrame(rows)

    if not df.empty:
        df["Saldo Awal"] = df.iloc[0]["Saldo"] - df.iloc[0]["Kredit"] + df.iloc[0]["Debit"]
        df = df[[
            "Nomor Rekening", "Tanggal", "Keterangan", "Debit", "Kredit", "Saldo", "Currency", "Saldo Awal"
        ]]
        st.success("Berikut data hasil ekstraksi:")
        st.dataframe(df, use_container_width=True)

        output = BytesIO()
        df.to_excel(output, index=False)
        st.download_button(
            label="📥 Download Excel",
            data=output.getvalue(),
            file_name=f"Rekening_{no_rekening}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Tidak ditemukan transaksi dalam PDF.")
