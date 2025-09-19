import streamlit as st
import pandas as pd
import altair as alt
import os

# === Fungsi bantu ===
def read_file(filepath, header_row):
    ext = os.path.splitext(filepath)[-1].lower()
    try:
        if ext in [".xls", ".xlsx"]:
            xls = pd.ExcelFile(filepath)
            sheets = []
            for sheet in xls.sheet_names:
                df = pd.read_excel(filepath, sheet_name=sheet, header=header_row)
                # isi header kosong dengan header sebelumnya
                df.columns = [
                    col if col and not str(col).startswith("Unnamed") else f"{df.columns[i-1]}_{i}"
                    if i > 0 else f"col_{i}"
                    for i, col in enumerate(df.columns)
                ]
                sheets.append(df)
            return pd.concat(sheets, ignore_index=True)
        elif ext == ".ods":
            xls = pd.ExcelFile(filepath, engine="odf")
            sheets = []
            for sheet in xls.sheet_names:
                df = pd.read_excel(filepath, sheet_name=sheet, header=header_row, engine="odf")
                df.columns = [
                    col if col and not str(col).startswith("Unnamed") else f"{df.columns[i-1]}_{i}"
                    if i > 0 else f"col_{i}"
                    for i, col in enumerate(df.columns)
                ]
                sheets.append(df)
            return pd.concat(sheets, ignore_index=True)
        else:
            return pd.DataFrame()
    except Exception as e:
        st.warning(f"Gagal baca {filepath}: {e}")
        return pd.DataFrame()

# === Streamlit App ===
st.title("📊 Penggabungan Multi Sheet & Multi File (ODS/Excel)")

uploaded_files = st.file_uploader(
    "Upload file Excel/ODS (boleh lebih dari satu)", 
    type=["xls", "xlsx", "ods"], 
    accept_multiple_files=True
)

header_row = st.number_input("Pilih baris header (mulai dari 0)", min_value=0, value=0, step=1)

if uploaded_files:
    all_data = []
    for file in uploaded_files:
        df = read_file(file, header_row)
        if not df.empty:
            all_data.append(df)

    if all_data:
        data_gabungan = pd.concat(all_data, ignore_index=True)
        st.subheader("📌 Data Gabungan")
        st.dataframe(data_gabungan)

        # Filter kolom
        kolom_pilihan = st.multiselect("Pilih kolom untuk penyaringan", options=data_gabungan.columns)
        if kolom_pilihan:
            data_penyaringan = data_gabungan[kolom_pilihan]
            st.subheader("📌 Data Penyaringan")
            st.dataframe(data_penyaringan)

            # Visualisasi
            if len(kolom_pilihan) >= 2:
                chart = alt.Chart(data_penyaringan).mark_point().encode(
                    x=kolom_pilihan[0],
                    y=kolom_pilihan[1],
                    tooltip=kolom_pilihan
                )
                st.altair_chart(chart, use_container_width=True)

            # Unduh hasil
            ext_download = st.radio("Pilih format unduhan", ["Excel (.xlsx)", "ODS (.ods)"])
            if st.button("Unduh Hasil Penyaringan"):
                if ext_download == "Excel (.xlsx)":
                    out_path = "hasil_penyaringan.xlsx"
                    data_penyaringan.to_excel(out_path, index=False)
                else:
                    out_path = "hasil_penyaringan.ods"
                    data_penyaringan.to_excel(out_path, index=False, engine="odf")

                with open(out_path, "rb") as f:
                    st.download_button("📥 Klik untuk unduh", f, file_name=out_path)
    else:
        st.info("Tidak ada data valid dari file yang diupload.")
else:
    st.info("Silakan upload file ODS/Excel terlebih dahulu.")

# === VISUALISASI ===
if len(kolom_pilihan) >= 2:
    col1, col2 = kolom_pilihan[:2]

    # cek tipe data
    if pd.api.types.is_numeric_dtype(data_penyaringan[col1]) and pd.api.types.is_numeric_dtype(data_penyaringan[col2]):
        chart = alt.Chart(data_penyaringan).mark_circle(size=60).encode(
            x=alt.X(col1, type="quantitative"),
            y=alt.Y(col2, type="quantitative"),
            tooltip=kolom_pilihan
        ).interactive()
        st.altair_chart(chart, use_container_width=True)

    elif pd.api.types.is_numeric_dtype(data_penyaringan[col1]) and not pd.api.types.is_numeric_dtype(data_penyaringan[col2]):
        chart = alt.Chart(data_penyaringan).mark_bar().encode(
            x=alt.X(col2, type="nominal"),
            y=alt.Y(col1, type="quantitative"),
            tooltip=kolom_pilihan
        )
        st.altair_chart(chart, use_container_width=True)

    elif not pd.api.types.is_numeric_dtype(data_penyaringan[col1]) and pd.api.types.is_numeric_dtype(data_penyaringan[col2]):
        chart = alt.Chart(data_penyaringan).mark_bar().encode(
            x=alt.X(col1, type="nominal"),
            y=alt.Y(col2, type="quantitative"),
            tooltip=kolom_pilihan
        )
        st.altair_chart(chart, use_container_width=True)

    else:
        chart = alt.Chart(data_penyaringan).mark_bar().encode(
            x=alt.X(col1, type="nominal"),
            y=alt.Y(col2, type="nominal"),
            tooltip=kolom_pilihan
        )
        st.altair_chart(chart, use_container_width=True)
        
