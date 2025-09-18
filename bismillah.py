import streamlit as st
import pandas as pd
import altair as alt

# Judul Aplikasi
st.title("📊 Aplikasi Filter & Visualisasi Data")
st.write("Upload file (CSV, Excel, ODS), pilih kolom yang ingin difilter, lalu tampilkan hasilnya dalam tabel dan grafik.")

# Upload File
uploaded_files = st.file_uploader(
    "Unggah satu atau lebih file",
    type=["csv", "xlsx", "ods"],
    accept_multiple_files=True
)

if uploaded_files:
    dataframes = []
    for uploaded_file in uploaded_files:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, sheet_name=None)  # semua sheet
            df = pd.concat(df.values(), ignore_index=True)
        elif uploaded_file.name.endswith(".ods"):
            df = pd.read_excel(uploaded_file, engine="odf", sheet_name=None)
            df = pd.concat(df.values(), ignore_index=True)
        dataframes.append(df)

    # Gabungkan semua file
    df = pd.concat(dataframes, ignore_index=True)

    st.subheader("📑 Data Asli")
    st.dataframe(df.head(20), use_container_width=True)

    # =====================
    # Filter Data
    # =====================
    st.subheader("🔎 Filter Data")

    all_cols = df.columns.tolist()
    filter_columns = st.multiselect("Pilih kolom yang ingin ditampilkan", all_cols)

    if filter_columns:
        filtered_df = df[filter_columns]  # hanya kolom yang dipilih
        st.dataframe(filtered_df, use_container_width=True)

        # =====================
        # Visualisasi
        # =====================
        st.subheader("📊 Visualisasi Data")

        x_axis = st.selectbox("Pilih sumbu X", filter_columns)
        y_axis = st.selectbox("Pilih sumbu Y", filter_columns)

        chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

        # Tentukan tipe data
        def col_type(col):
            if pd.api.types.is_numeric_dtype(filtered_df[col]):
                return "quantitative"
            elif pd.api.types.is_datetime64_any_dtype(filtered_df[col]):
                return "temporal"
            else:
                return "nominal"

        x_type = col_type(x_axis)
        y_type = col_type(y_axis)

        # Pilih grafik
        if chart_type == "Diagram Batang":
            chart = alt.Chart(filtered_df).mark_bar().encode(
                x=alt.X(x_axis, type=x_type),
                y=alt.Y(y_axis, type=y_type),
                tooltip=filter_columns
            )
        elif chart_type == "Diagram Garis":
            chart = alt.Chart(filtered_df).mark_line(point=True).encode(
                x=alt.X(x_axis, type=x_type),
                y=alt.Y(y_axis, type=y_type),
                tooltip=filter_columns
            )
        else:  # Sebar
            chart = alt.Chart(filtered_df).mark_circle(size=60).encode(
                x=alt.X(x_axis, type=x_type),
                y=alt.Y(y_axis, type=y_type),
                tooltip=filter_columns
            )

        st.altair_chart(chart, use_container_width=True)

        # =====================
        # Download Hasil
        # =====================
        st.subheader("💾 Unduh Hasil Filter")

        pilihan_download = st.radio("Pilih format unduhan", ["Excel (.xlsx)", "ODS (.ods)"])

        if pilihan_download == "Excel (.xlsx)":
            from io import BytesIO
            import xlsxwriter

            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="Hasil Filter")
            st.download_button(
                label="⬇️ Download Excel",
                data=output.getvalue(),
                file_name="hasil_filter.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        elif pilihan_download == "ODS (.ods)":
            from io import BytesIO
            import odf.opendocument
            from odf.table import Table, TableRow, TableCell
            from odf.text import P

            ods_doc = odf.opendocument.OpenDocumentSpreadsheet()
            table = Table(name="Hasil Filter")

            for _, row in filtered_df.iterrows():
                tr = TableRow()
                for cell in row:
                    tc = TableCell()
                    tc.addElement(P(text=str(cell)))
                    tr.addElement(tc)
                table.addElement(tr)

            ods_doc.spreadsheet.addElement(table)

            output = BytesIO()
            ods_doc.save(output)
            st.download_button(
                label="⬇️ Download ODS",
                data=output.getvalue(),
                file_name="hasil_filter.ods",
                mime="application/vnd.oasis.opendocument.spreadsheet"
            )
    else:
        st.info("Silakan pilih kolom yang ingin ditampilkan.")
