import streamlit as st
import pandas as pd
import altair as alt
import io

st.title("📊 Aplikasi Gabungan Data Excel/ODS")

# Upload banyak file
uploaded_files = st.file_uploader("Upload file Excel/ODS", type=["xlsx", "xls", "ods"], accept_multiple_files=True)

# Pilih baris header
header_row = st.number_input("Pilih baris header (mulai dari 0)", min_value=0, value=0)

data_frames = []

def perbaiki_header(df, header_row):
    # Ambil 2 baris untuk header (kalau ada)
    if header_row + 1 < len(df):
        headers = df.iloc[header_row:header_row+2].fillna(method="ffill").astype(str).values
        new_cols = []
        for col in zip(*headers):
            col_name = "_".join([c for c in col if c.strip() != ""])
            new_cols.append(col_name if col_name else "Unnamed")
        df = df.iloc[header_row+2:].reset_index(drop=True)
        df.columns = new_cols
    else:
        df = df.iloc[header_row+1:].reset_index(drop=True)
        df.columns = df.iloc[0]
        df = df[1:]
    return df

if uploaded_files:
    for file in uploaded_files:
        if file.name.endswith(".ods"):
            xls = pd.ExcelFile(file, engine="odf")
        else:
            xls = pd.ExcelFile(file)

        for sheet in xls.sheet_names:
            df = xls.parse(sheet, header=None)
            df = perbaiki_header(df, header_row)
            df["Sumber_File"] = file.name
            df["Sumber_Sheet"] = sheet
            data_frames.append(df)

    # Gabungan semua sheet & file
    data_gabungan = pd.concat(data_frames, ignore_index=True)

    st.subheader("📑 Data Gabungan")
    st.dataframe(data_gabungan)

    # Pilih kolom filter
    kolom_pilihan = st.multiselect("Pilih kolom untuk filter", data_gabungan.columns.tolist())

    if kolom_pilihan:
        data_penyaringan = data_gabungan[kolom_pilihan + ["Sumber_File", "Sumber_Sheet"]]
        st.subheader("🔎 Data Hasil Penyaringan")
        st.dataframe(data_penyaringan)

        # Visualisasi jika ada minimal 2 kolom numerik
        kolom_numerik = data_penyaringan.select_dtypes(include="number").columns.tolist()
        if len(kolom_numerik) >= 2:
            chart = alt.Chart(data_penyaringan).mark_bar().encode(
                x=alt.X(kolom_numerik[0], type="quantitative"),
                y=alt.Y(kolom_numerik[1], type="quantitative"),
                color="Sumber_File"
            ).interactive()
            st.altair_chart(chart, use_container_width=True)

            # Unduh visualisasi sebagai PNG
            from altair_saver import save
            buf = io.BytesIO()
            save(chart, buf, format="png")
            st.download_button("⬇️ Unduh Visualisasi PNG", buf, file_name="visualisasi.png")

        # Unduh data penyaringan
        output_xlsx = io.BytesIO()
        with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
            data_penyaringan.to_excel(writer, index=False)
        st.download_button("⬇️ Unduh Hasil Penyaringan (Excel)", output_xlsx.getvalue(), file_name="penyaringan.xlsx")

        output_ods = io.BytesIO()
        with pd.ExcelWriter(output_ods, engine="odf") as writer:
            data_penyaringan.to_excel(writer, index=False)
        st.download_button("⬇️ Unduh Hasil Penyaringan (ODS)", output_ods.getvalue(), file_name="penyaringan.ods")

    # Unduh data gabungan
    output_gabungan = io.BytesIO()
    with pd.ExcelWriter(output_gabungan, engine="openpyxl") as writer:
        data_gabungan.to_excel(writer, index=False)
    st.download_button("⬇️ Unduh Data Gabungan (Excel)", output_gabungan.getvalue(), file_name="gabungan.xlsx")

    # === VISUALISASI ===
    st.subheader("📈 Visualisasi Data")

    kolom_x = st.selectbox("Pilih kolom X", data_penyaringan.columns)
    kolom_y = st.selectbox("Pilih kolom Y", data_penyaringan.columns)
    jenis_chart = st.selectbox("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

    if kolom_x and kolom_y:
        # Deteksi tipe data
        def tipe_data(kolom):
            if pd.api.types.is_numeric_dtype(data_penyaringan[kolom]):
                return "quantitative"
            else:
                return "nominal"

        x_tipe = tipe_data(kolom_x)
        y_tipe = tipe_data(kolom_y)

        if jenis_chart == "Diagram Batang":
            chart = alt.Chart(data_penyaringan).mark_bar().encode(
                x=alt.X(kolom_x, type=x_tipe),
                y=alt.Y(kolom_y, type=y_tipe),
                color="Sumber_File"
            ).interactive()

        elif jenis_chart == "Diagram Garis":
            chart = alt.Chart(data_penyaringan).mark_line(point=True).encode(
                x=alt.X(kolom_x, type=x_tipe),
                y=alt.Y(kolom_y, type=y_tipe),
                color="Sumber_File"
            ).interactive()

        elif jenis_chart == "Diagram Sebar":
            chart = alt.Chart(data_penyaringan).mark_circle(size=80).encode(
                x=alt.X(kolom_x, type=x_tipe),
                y=alt.Y(kolom_y, type=y_tipe),
                color="Sumber_File",
                tooltip=list(data_penyaringan.columns)
            ).interactive()

        st.altair_chart(chart, use_container_width=True)

        
