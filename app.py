
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

    # Ekstrak nomor rekening dan mata uang
    norekening = re.search(r'Account No\.\s*(\d+)', text)
    currency = re.search(r'Currency\s+([A-Z]+)', text)
    no_rekening = norekening.group(1) if norekening else "-"
    mata_uang = currency.group(1) if currency else "-"

    # Ekstrak transaksi
    transaksi_pattern = re.findall(
        r"(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}:\d{2}(.*?)\s+([\-\d,.]+)\s+([\-\d,.]+)\s+([\-\d,.]+)",
        text,
        re.DOTALL,
    )

    rows = []
    for tgl, ket, debit, kredit, saldo in transaksi_pattern:
        # Gabungkan dan bersihkan keterangan
        keterangan = " ".join(ket.strip().splitlines()).strip()

        def parse_amount(val):
            val = val.replace(",", "")
            return float(val) if val not in ["-", ""] else 0.0

        rows.append({
            "Nomor Rekening": no_rekening,
            "Tanggal": pd.to_datetime(tgl, format="%d/%m/%Y").strftime("%d/%m/%Y"),
            "Keterangan": keterangan,
            "Debit": parse_amount(debit),
            "Kredit": parse_amount(kredit),
            "Saldo": parse_amount(saldo),
            "Currency": mata_uang,
        })

    df = pd.DataFrame(rows)

    if not df.empty:
        df["Saldo Awal"] = df["Saldo"].iloc[0] - df["Kredit"].iloc[0] + df["Debit"].iloc[0]
        df = df[[
            "Nomor Rekening", "Tanggal", "Keterangan", "Debit", "Kredit", "Saldo", "Currency", "Saldo Awal"
        ]]
        st.success("Berikut data hasil ekstraksi:")
        st.dataframe(df, use_container_width=True)

        # Download button
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
