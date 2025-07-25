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
    text = "\n".join(page.get_text().strip() for page in doc)

    st.subheader("🔍 Isi Mentah PDF:")
    with st.expander("Lihat isi teks hasil ekstraksi PDF"):
        st.code(text)

    lines = text.splitlines()
    norekening = re.search(r'Account No\.\s*(\d+)', text)
    currency = re.search(r'Currency\s+([A-Z]+)', text)
    no_rekening = norekening.group(1) if norekening else "-"
    mata_uang = currency.group(1) if currency else "-"

    rows = []
    current = {}
    buffer = []

    def extract_angka_and_ket(buf):
        for i in [-1, -2]:
            if abs(i) <= len(buf):
                nums = re.findall(r"[-\d,.]+", buf[i])
                if len(nums) >= 3:
                    keterangan = " ".join(buf[:i]) if i != -len(buf) else buf[0]
                    return nums[-3:], keterangan
        return None, None

    for line in lines:
        line = line.strip()
        if re.match(r"\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}", line):
            if buffer and "tanggal" in current:
                angka, ket = extract_angka_and_ket(buffer)
                if angka:
                    rows.append({
                        "Nomor Rekening": no_rekening,
                        "Tanggal": current["tanggal"].split()[0],
                        "Keterangan": ket.strip(),
                        "Debit": float(angka[0].replace(",", "")) if angka[0] != "0.00" else 0.0,
                        "Kredit": float(angka[1].replace(",", "")) if angka[1] != "0.00" else 0.0,
                        "Saldo": float(angka[2].replace(",", "")),
                        "Currency": mata_uang
                    })
            current = {"tanggal": line.strip()}
            buffer = []
        else:
            buffer.append(line)

    if buffer and "tanggal" in current:
        angka, ket = extract_angka_and_ket(buffer)
        if angka:
            rows.append({
                "Nomor Rekening": no_rekening,
                "Tanggal": current["tanggal"].split()[0],
                "Keterangan": ket.strip(),
                "Debit": float(angka[0].replace(",", "")) if angka[0] != "0.00" else 0.0,
                "Kredit": float(angka[1].replace(",", "")) if angka[1] != "0.00" else 0.0,
                "Saldo": float(angka[2].replace(",", "")),
                "Currency": mata_uang
            })

    df = pd.DataFrame(rows)

    if not df.empty:
        df["Tanggal"] = pd.to_datetime(df["Tanggal"], format="%d/%m/%Y").dt.strftime("%d/%m/%Y")
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
