import streamlit as st
import pandas as pd
import altair as alt

st.title("📊 Aplikasi Gabung, Filter, dan Visualisasi Data Multi-Sheet")
st.write("Unggah beberapa file (Excel/CSV/ODS) sekaligus")

# ====================
# Upload file
# ====================
uploaded_files = st.file_uploader(
    "Upload file (Excel/CSV/ODS)", 
    type=["xlsx", "csv", "ods"], 
    accept_multiple_files=True
)

header_row = st.number_input(
    "Pilih baris header (mulai dari 0, misalnya baris ke-6 berarti 5)", 
    min_value=0, 
    value=0
)

semua_data = []

def baca_file(file, header_row):
    ext = file.name.split(".")[-1].lower()
    dfs = []

    if ext in ["xlsx", "xls"]:
        xls = pd.ExcelFile(file)
        for sheet in xls.sheet_names:
            raw_df = pd.read_excel(file, sheet_name=sheet, header=None, engine="openpyxl")
            # Isi baris kosong sebelum header dengan baris sebelumnya
            raw_df.iloc[:header_row] = raw_df.iloc[:header_row].ffill()
            # Baca ulang dengan header yang benar
            df = pd.read_excel(file, sheet_name=sheet, header=header_row, engine="openpyxl")
            dfs.append(df)

    elif ext == "ods":
        xls = pd.ExcelFile(file, engine="odf")
        for sheet in xls.sheet_names:
            raw_df = pd.read_excel(file, sheet_name=sheet, header=None, engine="odf")
            raw_df.iloc[:header_row] = raw_df.iloc[:header_row].ffill()
            df = pd.read_excel(file, sheet_name=sheet, header=header_row, engine="odf")
            dfs.append(df)

    elif ext == "csv":
        raw_df = pd.read_csv(file, header=None)
        raw_df.iloc[:header_row] = raw_df.iloc[:header_row].ffill()
        df = pd.read_csv(file, header=header_row)
        dfs.append(df)

    return dfs

if uploaded_files:
    for file in uploaded_files:
        semua_data.extend(baca_file(file, header_row))

if semua_data:
    all_data = pd.concat(semua_data, ignore_index=True)

    # ========================
    # Isi data kosong setelah header
    # ========================
    for col in all_data.columns:
        if all_data[col].dtype == "object":
            all_data[col] = all_data[col].fillna("Kosong")
        else:
            all_data[col] = all_data[col].fillna(0)

    # ========================
    # Tampilkan data gabungan
    # ========================
    st.subheader("🔎 Data Gabungan (setelah diproses)")
    st.dataframe(all_data.head())

    # ========================
    # Filter kolom
    # ========================
    st.subheader("Pilih kolom yang ingin difilter & ditampilkan")
    filter_columns = st.multiselect("Pilih kolom:", all_data.columns)

    filtered_data = all_data.copy()
    for col in filter_columns:
        unique_vals = filtered_data[col].unique().tolist()
        selected_vals = st.multiselect(f"Pilih nilai untuk kolom {col}", unique_vals)
        if selected_vals:
            filtered_data = filtered_data[filtered_data[col].isin(selected_vals)]

    st.subheader("📌 Hasil Penyaringan")
    st.dataframe(filtered_data.head())

    # ========================
    # Visualisasi
    # ========================
    if not filtered_data.empty:
        st.subheader("📈 Visualisasi Data")
        x_axis = st.selectbox("Pilih kolom sumbu X", all_data.columns)
        y_axis = st.selectbox("Pilih kolom sumbu Y", all_data.columns)
        chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

        chart = None
        if chart_type == "Diagram Batang":
            chart = alt.Chart(filtered_data).mark_bar().encode(
                x=x_axis, y=y_axis, tooltip=list(filtered_data.columns)
            )
        elif chart_type == "Diagram Garis":
            chart = alt.Chart(filtered_data).mark_line(point=True).encode(
                x=x_axis, y=y_axis, tooltip=list(filtered_data.columns)
            )
        elif chart_type == "Diagram Sebar":
            chart = alt.Chart(filtered_data).mark_circle(size=60).encode(
                x=x_axis, y=y_axis, tooltip=list(filtered_data.columns)
            )

        if chart:
            st.altair_chart(chart, use_container_width=True)
            
