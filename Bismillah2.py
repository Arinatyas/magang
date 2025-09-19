import streamlit as st
import pandas as pd
import os
import altair as alt

st.title("📊 Aplikasi Penggabungan & Visualisasi Data")

# ======================
# Upload File
# ======================
uploaded_files = st.file_uploader(
    "Upload file Excel/ODS (bisa lebih dari satu)", 
    type=["xlsx", "xls", "ods"], 
    accept_multiple_files=True
)

data_frames = []
if uploaded_files:
    header_row = st.number_input("Pilih baris header (mulai dari 0)", min_value=0, value=0)

    for file in uploaded_files:
        try:
            df = pd.read_excel(file, header=header_row, engine="openpyxl")
        except:
            df = pd.read_excel(file, header=header_row, engine="odf")
        data_frames.append(df)

    if data_frames:
        data_gabungan = pd.concat(data_frames, ignore_index=True)

        st.subheader("📄 Data Gabungan")
        st.dataframe(data_gabungan)

        # ======================
        # Penyaringan
        # ======================
        st.subheader("🔍 Penyaringan Data")
        kolom_filter = st.multiselect("Pilih kolom untuk filter", data_gabungan.columns)

        data_penyaringan = data_gabungan.copy()
        for kol in kolom_filter:
            unique_vals = data_gabungan[kol].dropna().unique().tolist()
            pilihan = st.multiselect(f"Pilih nilai untuk {kol}", unique_vals)
            if pilihan:
                data_penyaringan = data_penyaringan[data_penyaringan[kol].isin(pilihan)]

        st.write("### Data Setelah Penyaringan")
        st.dataframe(data_penyaringan)

        # ======================
        # Unduh Data
        # ======================
        st.subheader("💾 Unduh Data")

        out_excel = "data_gabungan.xlsx"
        out_ods = "data_gabungan.ods"
        data_penyaringan.to_excel(out_excel, index=False, engine="openpyxl")
        data_penyaringan.to_excel(out_ods, index=False, engine="odf")

        with open(out_excel, "rb") as f:
            st.download_button("📥 Unduh Excel (.xlsx)", f, file_name=out_excel)

        with open(out_ods, "rb") as f:
            st.download_button("📥 Unduh ODS (.ods)", f, file_name=out_ods)

        # ======================
        # Visualisasi
        # ======================
        st.subheader("📈 Visualisasi Data")

        if len(data_penyaringan.columns) >= 2:
            x_axis = st.selectbox("Pilih Kolom X", data_penyaringan.columns)
            y_axis = st.selectbox("Pilih Kolom Y", data_penyaringan.columns)
            chart_type = st.selectbox("Pilih Jenis Grafik", ["Bar", "Line", "Scatter"])

            # Ganti NaN dengan 0
            data_penyaringan = data_penyaringan.fillna(0)

            # Deteksi tipe data otomatis
            def detect_type(series):
                try:
                    pd.to_numeric(series)
                    return "quantitative"
                except:
                    return "ordinal"

            x_type = detect_type(data_penyaringan[x_axis])
            y_type = detect_type(data_penyaringan[y_axis])

            if chart_type == "Bar":
                chart = alt.Chart(data_penyaringan).mark_bar().encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=list(data_penyaringan.columns)
                )
            elif chart_type == "Line":
                chart = alt.Chart(data_penyaringan).mark_line(point=True).encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=list(data_penyaringan.columns)
                )
            else:  # Scatter
                chart = alt.Chart(data_penyaringan).mark_circle(size=60).encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=list(data_penyaringan.columns)
                )

            st.altair_chart(chart, use_container_width=True)
            
