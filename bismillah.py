import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO

st.title("📊 Upload, Filter, dan Visualisasi Multi File")

# Upload beberapa file
uploaded_files = st.file_uploader(
    "Unggah beberapa file Excel/CSV/ODS",
    type=["xlsx", "csv", "ods"],
    accept_multiple_files=True
)

if uploaded_files:
    dfs = []

    for file in uploaded_files:
        try:
            if file.name.endswith(".csv"):
                df = pd.read_csv(file)
            elif file.name.endswith(".xlsx"):
                df = pd.read_excel(file, engine="openpyxl")
            elif file.name.endswith(".ods"):
                df = pd.read_excel(file, engine="odf")
            else:
                st.warning(f"Format file {file.name} tidak didukung, dilewati.")
                continue

            df["Sumber_File"] = file.name
            dfs.append(df)

        except Exception as e:
            st.error(f"Gagal membaca {file.name}: {e}")

    if dfs:
        all_data = pd.concat(dfs, ignore_index=True)

        st.subheader("🔎 Data Gabungan")
        st.dataframe(all_data.head())

        # Filter dinamis
        filter_columns = st.multiselect("Pilih kolom untuk filter", all_data.columns)
        filtered_df = all_data.copy()

        if filter_columns:
            for col in filter_columns:
                unique_vals = filtered_df[col].dropna().unique().tolist()
                selected_vals = st.multiselect(f"Filter {col}", unique_vals)
                if selected_vals:
                    filtered_df = filtered_df[filtered_df[col].isin(selected_vals)]

        st.subheader("📌 Hasil Filter")
        st.dataframe(filtered_df)

        # =====================
        # Visualisasi Grafik
        # =====================
        st.subheader("📈 Visualisasi Data")

        # Bersihkan nama kolom
        all_cols = [c.strip() for c in filtered_df.columns.tolist()]
        filtered_df.columns = all_cols

        # Deteksi tipe kolom
        numeric_cols = filtered_df.select_dtypes(include=['int64', 'float64']).columns.tolist()
        datetime_cols = filtered_df.select_dtypes(include=['datetime64[ns]']).columns.tolist()

        if len(numeric_cols) > 0 and len(all_cols) > 1:
            x_axis = st.selectbox("Pilih kolom X (kategori/tanggal)", all_cols)
            y_axis = st.selectbox("Pilih kolom Y (numerik)", numeric_cols)

            chart_type = st.radio("Jenis Grafik", ["Bar Chart", "Line Chart", "Scatter Plot"])

            # Tentukan tipe data otomatis
            if x_axis in numeric_cols:
                x_type = "quantitative"
            elif x_axis in datetime_cols:
                x_type = "temporal"
            else:
                x_type = "nominal"

            if y_axis in numeric_cols:
                y_type = "quantitative"
            elif y_axis in datetime_cols:
                y_type = "temporal"
            else:
                y_type = "nominal"

            # Pilih jenis chart
            if chart_type == "Bar Chart":
                chart = alt.Chart(filtered_df).mark_bar().encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=all_cols
                )
            elif chart_type == "Line Chart":
                chart = alt.Chart(filtered_df).mark_line(point=True).encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=all_cols
                )
            else:  # Scatter
                chart = alt.Chart(filtered_df).mark_circle(size=60).encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=all_cols
                )

            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("Tidak ada kolom numerik yang bisa dibuat grafik.")

        # =====================
        # Download Hasil
        # =====================
        st.subheader("💾 Download Hasil")

        # CSV
        csv = filtered_df.to_csv(index=False).encode("utf-8")
        st.download_button("Download CSV", csv, "hasil_filter.csv", "text/csv")

        # Excel
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
        st.download_button("Download Excel (.xlsx)", excel_buffer.getvalue(),
                           "hasil_filter.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # ODS
        ods_buffer = BytesIO()
        with pd.ExcelWriter(ods_buffer, engine="odf") as writer:
            filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
        st.download_button("Download ODS (.ods)", ods_buffer.getvalue(),
                           "hasil_filter.ods", "application/vnd.oasis.opendocument.spreadsheet")
