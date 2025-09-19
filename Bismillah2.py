import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P

st.set_page_config(page_title="Data Gabungan & Visualisasi", layout="wide")

st.title("📊 Aplikasi Data Gabungan & Visualisasi dengan Filter")

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

        # Isi NaN → 0 (numerik) atau "Tidak Ada" (lainnya)
        for col in data_gabungan.columns:
            if pd.api.types.is_numeric_dtype(data_gabungan[col]):
                data_gabungan[col] = data_gabungan[col].fillna(0)
            else:
                data_gabungan[col] = data_gabungan[col].fillna("Tidak Ada")

        st.subheader("📑 Data Gabungan")
        st.dataframe(data_gabungan.head(50), use_container_width=True)

        # ======================
        # Filter Data
        # ======================
        st.subheader("🔎 Filter Data")

        filter_col = st.selectbox("Pilih kolom untuk filter", ["(Tidak ada)"] + list(data_gabungan.columns))

        if filter_col != "(Tidak ada)":
            if pd.api.types.is_numeric_dtype(data_gabungan[filter_col]):
                min_val, max_val = float(data_gabungan[filter_col].min()), float(data_gabungan[filter_col].max())
                range_min, range_max = st.slider(
                    f"Pilih rentang nilai untuk {filter_col}",
                    min_val, max_val, (min_val, max_val)
                )
                data_filtered = data_gabungan[
                    (data_gabungan[filter_col] >= range_min) & (data_gabungan[filter_col] <= range_max)
                ]
            else:
                unique_vals = data_gabungan[filter_col].unique().tolist()
                selected_vals = st.multiselect(f"Pilih nilai untuk {filter_col}", unique_vals, default=unique_vals)
                data_filtered = data_gabungan[data_gabungan[filter_col].isin(selected_vals)]
        else:
            data_filtered = data_gabungan

        st.write("📌 Jumlah data setelah filter:", len(data_filtered))
        st.dataframe(data_filtered.head(50), use_container_width=True)

        # ======================
        # Pilihan kolom X dan Y
        # ======================
        st.subheader("⚙️ Pilih Kolom untuk Visualisasi")

        kolom_x = st.selectbox("Kolom X", data_filtered.columns)
        kolom_y = st.selectbox("Kolom Y", data_filtered.columns)

        # ======================
        # Pilihan Jenis Grafik
        # ======================
        chart_type = st.radio(
            "Pilih jenis grafik",
            ["Bar", "Line", "Scatter", "Pie"]
        )

        st.subheader("📊 Visualisasi")

        chart = None

        if chart_type == "Bar":
            chart = alt.Chart(data_filtered).mark_bar().encode(
                x=kolom_x,
                y=kolom_y,
                tooltip=[kolom_x, kolom_y]
            )
        elif chart_type == "Line":
            chart = alt.Chart(data_filtered).mark_line(point=True).encode(
                x=kolom_x,
                y=kolom_y,
                tooltip=[kolom_x, kolom_y]
            )
        elif chart_type == "Scatter":
            chart = alt.Chart(data_filtered).mark_circle(size=60).encode(
                x=kolom_x,
                y=kolom_y,
                tooltip=[kolom_x, kolom_y]
            )
        elif chart_type == "Pie":
            chart = alt.Chart(data_filtered).mark_arc().encode(
                theta=kolom_y,
                color=kolom_x,
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
            data_filtered.to_excel(writer, index=False, sheet_name="Gabungan")
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
        # Tambahkan header
        header_row = TableRow()
        for col in data_filtered.columns:
            tc = TableCell()
            tc.addElement(P(text=str(col)))
            header_row.addElement(tc)
        table.addElement(header_row)

        # Tambahkan isi data
        for _, row in data_filtered.iterrows():
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
                            
