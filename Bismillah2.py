import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P

st.title("📊 Gabung, Filter, dan Visualisasi Multi-Sheet")

# Upload file (boleh banyak & format campur)
uploaded_files = st.file_uploader("Unggah file Excel/CSV/ODS", type=["xlsx","csv","ods"], accept_multiple_files=True)

header_rows = {}
all_sheets = {}

if uploaded_files:
    for file in uploaded_files:
        file_name = file.name

        if file_name.endswith(".csv"):
            df = pd.read_csv(file, header=None)
            header_opt = st.number_input(f"Pilih baris header untuk {file_name}", min_value=0, max_value=len(df)-1, value=0)
            df = pd.read_csv(file, header=header_opt)
            all_sheets[file_name] = {"Sheet1": df}

        elif file_name.endswith(".xlsx"):
            xls = pd.ExcelFile(file)
            all_sheets[file_name] = {}
            for sheet in xls.sheet_names:
                df_temp = pd.read_excel(file, sheet_name=sheet, header=None)
                header_opt = st.number_input(f"Pilih baris header untuk {file_name} - {sheet}", min_value=0, max_value=len(df_temp)-1, value=0)
                df = pd.read_excel(file, sheet_name=sheet, header=header_opt)
                all_sheets[file_name][sheet] = df

        elif file_name.endswith(".ods"):
            xls = pd.ExcelFile(file, engine="odf")
            all_sheets[file_name] = {}
            for sheet in xls.sheet_names:
                df_temp = pd.read_excel(file, sheet_name=sheet, header=None, engine="odf")
                header_opt = st.number_input(f"Pilih baris header untuk {file_name} - {sheet}", min_value=0, max_value=len(df_temp)-1, value=0)
                df = pd.read_excel(file, sheet_name=sheet, header=header_opt, engine="odf")
                all_sheets[file_name][sheet] = df

# 🔹 Gabungkan semua sheet
gabungan = []
for fname, sheets in all_sheets.items():
    for sname, df in sheets.items():
        df["Sumber_File"] = fname
        df["Sumber_Sheet"] = sname
        gabungan.append(df)

if gabungan:
    data_gabungan = pd.concat(gabungan, ignore_index=True)

    # isi kosong → kategori: "Tidak Ada", numerik: 0
    for col in data_gabungan.columns:
        if data_gabungan[col].dtype == "object":
            data_gabungan[col] = data_gabungan[col].fillna("Tidak Ada")
        else:
            data_gabungan[col] = pd.to_numeric(data_gabungan[col], errors="coerce").fillna(0)

    st.subheader("📌 Data Gabungan")
    st.dataframe(data_gabungan.head())

    # Pilih kolom filter
    st.subheader("🔎 Penyaringan Data")
    kolom_filter = st.multiselect("Pilih kolom untuk filter", data_gabungan.columns.tolist())
    data_penyaringan = data_gabungan.copy()
    for col in kolom_filter:
        val = st.multiselect(f"Pilih nilai untuk {col}", data_gabungan[col].unique().tolist())
        if val:
            data_penyaringan = data_penyaringan[data_penyaringan[col].isin(val)]

    st.dataframe(data_penyaringan)

    # Pilihan X Y
    st.subheader("📈 Visualisasi Data")
    kolom_x = st.selectbox("Pilih kolom X", data_penyaringan.columns)
    kolom_y = st.selectbox("Pilih kolom Y", data_penyaringan.columns)

    if kolom_x and kolom_y:
        chart = alt.Chart(data_penyaringan).mark_bar().encode(
            x=alt.X(kolom_x, type="nominal" if data_penyaringan[kolom_x].dtype=="object" else "quantitative"),
            y=alt.Y(kolom_y, type="nominal" if data_penyaringan[kolom_y].dtype=="object" else "quantitative")
        ).interactive()

        st.altair_chart(chart, use_container_width=True)

    # Unduh Excel
    excel_buf = BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        data_gabungan.to_excel(writer, index=False, sheet_name="Gabungan")
    st.download_button("⬇️ Unduh Excel", data=excel_buf.getvalue(),
                       file_name="data_gabungan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Unduh ODS
    ods_doc = OpenDocumentSpreadsheet()
    table = Table(name="Gabungan")
    for _, row in data_gabungan.iterrows():
        tr = TableRow()
        for val in row:
            tc = TableCell()
            tc.addElement(P(text=str(val)))
            tr.addElement(tc)
        table.addElement(tr)
    ods_doc.spreadsheet.addElement(table)

    ods_buf = BytesIO()
    ods_doc.save(ods_buf)
    st.download_button("⬇️ Unduh ODS", data=ods_buf.getvalue(),
                       file_name="data_gabungan.ods", mime="application/vnd.oasis.opendocument.spreadsheet")
    
