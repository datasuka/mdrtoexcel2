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

    for line in lines:
        line = line.strip()
        if re.match(r"\d{2}/\d{2}/\d{4}", line):
            if buffer and "tanggal" in current:
                nums = re.findall(r"[-\d,.]+", buffer[-1])[-3:]
                ket = " ".join(buffer[:-1]) if len(buffer) > 1 else buffer[0]
                if len(nums) == 3:
                    rows.append({
                        "Nomor Rekening": no_rekening,
                        "Tanggal": current["tanggal"],
                        "Keterangan": ket.strip(),
                        "Debit": float(nums[0].replace(",", "")) if nums[0] != "0.00" else 0.0,
                        "Kredit": float(nums[1].replace(",", "")) if nums[1] != "0.00" else 0.0,
                        "Saldo": float(nums[2].replace(",", "")),
                        "Currency": mata_uang
                    })
            current = {"tanggal": line.strip()}
            buffer = []
        else:
            buffer.append(line)

    if buffer and "tanggal" in current:
        nums = re.findall(r"[-\d,.]+", buffer[-1])[-3:]
        ket = " ".join(buffer[:-1]) if len(buffer) > 1 else buffer[0]
        if len(nums) == 3:
            rows.append({
                "Nomor Rekening": no_rekening,
                "Tanggal": current["tanggal"],
                "Keterangan": ket.strip(),
                "Debit": float(nums[0].replace(",", "")) if nums[0] != "0.00" else 0.0,
                "Kredit": float(nums[1].replace(",", "")) if nums[1] != "0.00" else 0.0,
                "Saldo": float(nums[2].replace(",", "")),
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
