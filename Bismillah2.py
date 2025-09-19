import streamlit as st
import pandas as pd
import altair as alt
import io

st.set_page_config(page_title="📊 Aplikasi Gabung, Filter, dan Visualisasi Data", layout="wide")

st.title("📊 Aplikasi Gabung, Filter, dan Visualisasi Data Multi-Sheet")

# --- Upload File ---
uploaded_files = st.file_uploader(
    "Unggah beberapa file (Excel/CSV/ODS) sekaligus",
    type=["xlsx", "csv", "ods"],
    accept_multiple_files=True
)

# Pilih baris header
header_row = st.number_input("Pilih baris header (mulai dari 0, misalnya baris ke-6 berarti 5)", 
                             min_value=0, value=0, step=1)

all_data = []

if uploaded_files:
    for file in uploaded_files:
        if file.name.endswith(".csv"):
            df = pd.read_csv(file, header=header_row)
        elif file.name.endswith(".xlsx"):
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet, header=header_row)
                df["Sumber_Sheet"] = sheet
                df["Sumber_File"] = file.name
                all_data.append(df)
        elif file.name.endswith(".ods"):
            xls = pd.ExcelFile(file, engine="odf")
            for sheet in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet, header=header_row, engine="odf")
                df["Sumber_Sheet"] = sheet
                df["Sumber_File"] = file.name
                all_data.append(df)

    if all_data:
        data_gabungan = pd.concat(all_data, ignore_index=True)

        # Pastikan kolom unik
        data_gabungan.columns = pd.io.parsers.ParserBase(
            {"names": data_gabungan.columns}
        )._maybe_dedup_names(data_gabungan.columns)

        st.subheader("🔎 Data Gabungan (setelah diproses)")
        st.dataframe(data_gabungan.head(50))

        # --- FILTER DATA ---
        st.subheader("📌 Hasil Penyaringan")

        kolom_pilihan = st.multiselect("Pilih kolom yang ingin difilter & ditampilkan", data_gabungan.columns)

        if kolom_pilihan:
            data_penyaringan = data_gabungan[kolom_pilihan].copy()

            for col in kolom_pilihan:
                opsi = st.multiselect(f"Pilih nilai untuk kolom {col}", sorted(data_gabungan[col].dropna().unique()))
                if opsi:
                    data_penyaringan = data_penyaringan[data_penyaringan[col].isin(opsi)]

            # Pastikan kolom unik
            data_penyaringan.columns = pd.io.parsers.ParserBase(
                {"names": data_penyaringan.columns}
            )._maybe_dedup_names(data_penyaringan.columns)

            st.dataframe(data_penyaringan.head(50))

            # --- DOWNLOAD DATA ---
            with io.BytesIO() as buffer:
                data_penyaringan.to_excel(buffer, index=False, engine="openpyxl")
                st.download_button(
                    label="💾 Unduh Hasil Penyaringan (Excel)",
                    data=buffer.getvalue(),
                    file_name="hasil_penyaringan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with io.BytesIO() as buffer:
                data_penyaringan.to_csv(buffer, index=False)
                st.download_button(
                    label="💾 Unduh Hasil Penyaringan (CSV)",
                    data=buffer.getvalue(),
                    file_name="hasil_penyaringan.csv",
                    mime="text/csv"
                )

            # --- VISUALISASI ---
            st.subheader("📈 Visualisasi Data")
            if len(kolom_pilihan) >= 2:
                x_axis = st.selectbox("Pilih kolom sumbu X", kolom_pilihan)
                y_axis = st.selectbox("Pilih kolom sumbu Y", kolom_pilihan)

                chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

                if x_axis and y_axis:
                    if chart_type == "Diagram Batang":
                        chart = alt.Chart(data_penyaringan).mark_bar().encode(x=x_axis, y=y_axis)
                    elif chart_type == "Diagram Garis":
                        chart = alt.Chart(data_penyaringan).mark_line().encode(x=x_axis, y=y_axis)
                    else:
                        chart = alt.Chart(data_penyaringan).mark_point().encode(x=x_axis, y=y_axis)

                    st.altair_chart(chart, use_container_width=True)
            else:
                st.info("Pilih minimal 2 kolom untuk membuat visualisasi.")
        else:
            st.info("Silakan pilih kolom untuk melakukan penyaringan.")
    else:
        st.warning("Tidak ada data yang berhasil digabung.")
else:
    st.info("Silakan unggah minimal 1 file untuk mulai.")
    
