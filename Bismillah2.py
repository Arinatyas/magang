import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P

st.set_page_config(page_title="Data Gabungan & Visualisasi", layout="wide")

st.title("📊 Aplikasi Data Gabungan & Visualisasi")

# ======================
# Upload file
# ======================
uploaded_files = st.file_uploader(
    "Unggah file (Excel .xlsx, ODS .ods, CSV)", 
    type=["xlsx", "ods", "csv"], 
    accept_multiple_files=True
)

all_data = []

if uploaded_files:
    for file in uploaded_files:
        if file.name.endswith(".xlsx"):
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet)
                df["Source_File"] = file.name
                df["Source_Sheet"] = sheet
                all_data.append(df)

        elif file.name.endswith(".ods"):
            try:
                import pyexcel_ods3
                data_ods = pyexcel_ods3.get_data(file)
                for sheet, values in data_ods.items():
                    df = pd.DataFrame(values[1:], columns=values[0])
                    df["Source_File"] = file.name
                    df["Source_Sheet"] = sheet
                    all_data.append(df)
            except Exception as e:
                st.error(f"Gagal membaca ODS {file.name}: {e}")

        elif file.name.endswith(".csv"):
            df = pd.read_csv(file)
            df["Source_File"] = file.name
            df["Source_Sheet"] = "CSV"
            all_data.append(df)

    # Gabungkan semua data
    if all_data:
        data_gabungan = pd.concat(all_data, ignore_index=True)

        # Kolom kosong dinamai Unnamed_x
        data_gabungan.columns = [
            str(c) if c != "" else f"Unnamed_{i}"
            for i, c in enumerate(data_gabungan.columns)
        ]

        st.subheader("📑 Data Gabungan")
        st.dataframe(data_gabungan.head(50), use_container_width=True)

        # ======================
        # Pilihan kolom X dan Y
        # ======================
        st.subheader("⚙️ Pilih Kolom untuk Visualisasi")

        kolom_x = st.selectbox("Kolom X", data_gabungan.columns)
        kolom_y = st.selectbox("Kolom Y", data_gabungan.columns)

        # Deteksi tipe data
        tipe_x = "numerik" if pd.api.types.is_numeric_dtype(data_gabungan[kolom_x]) else "kategori"
        tipe_y = "numerik" if pd.api.types.is_numeric_dtype(data_gabungan[kolom_y]) else "kategori"

        st.write(f"🔹 Kolom X = {kolom_x} ({tipe_x})")
        st.write(f"🔹 Kolom Y = {kolom_y} ({tipe_y})")

        # ======================
        # Visualisasi
        # ======================
        st.subheader("📊 Visualisasi")

        chart = None
        if tipe_x == "kategori" and tipe_y == "numerik":
            chart = alt.Chart(data_gabungan).mark_bar().encode(
                x=alt.X(kolom_x, type="nominal"),
                y=alt.Y(kolom_y, type="quantitative"),
                tooltip=[kolom_x, kolom_y]
            )
        elif tipe_x == "numerik" and tipe_y == "numerik":
            chart = alt.Chart(data_gabungan).mark_circle(size=60).encode(
                x=alt.X(kolom_x, type="quantitative"),
                y=alt.Y(kolom_y, type="quantitative"),
                tooltip=[kolom_x, kolom_y]
            )
        elif tipe_x == "kategori" and tipe_y == "kategori":
            chart = alt.Chart(data_gabungan).mark_bar().encode(
                x=alt.X(kolom_x, type="nominal"),
                y=alt.Y(kolom_y, type="nominal"),
                tooltip=[kolom_x, kolom_y]
            )
        else:
            chart = alt.Chart(data_gabungan).mark_line(point=True).encode(
                x=alt.X(kolom_x, type="ordinal"),
                y=alt.Y(kolom_y, type="quantitative"),
                tooltip=[kolom_x, kolom_y]
            )

        if chart:
            st.altair_chart(chart, use_container_width=True)

        # ======================
        # Unduh hasil gabungan
        # ======================
        st.subheader("📥 Unduh Data Gabungan")

        # Simpan Excel
        buffer_xlsx = BytesIO()
        with pd.ExcelWriter(buffer_xlsx, engine="openpyxl") as writer:
            data_gabungan.to_excel(writer, index=False, sheet_name="Gabungan")
        buffer_xlsx.seek(0)

        st.download_button(
            label="Unduh Excel (.xlsx)",
            data=buffer_xlsx,
            file_name="data_gabungan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Simpan ODS
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
        buffer_ods = BytesIO()
        ods_doc.save(buffer_ods)
        buffer_ods.seek(0)

        st.download_button(
            label="Unduh ODS (.ods)",
            data=buffer_ods,
            file_name="data_gabungan.ods",
            mime="application/vnd.oasis.opendocument.spreadsheet"
        )

else:
    st.info("📂 Silakan upload minimal satu file untuk mulai.")
                
