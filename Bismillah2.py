import streamlit as st
import pandas as pd
import altair as alt
import io

st.set_page_config(page_title="📊 Gabung & Visualisasi Multi-Sheet", layout="wide")

st.title("📊 Aplikasi Gabung, Filter, dan Visualisasi Data Multi-Sheet")

# =========================
# 1. Upload file
# =========================
uploaded_files = st.file_uploader(
    "Unggah beberapa file (Excel/CSV/ODS) sekaligus",
    type=["xlsx", "csv", "ods"],
    accept_multiple_files=True
)

header_row = st.number_input(
    "Pilih baris header (mulai dari 0, misalnya baris ke-6 berarti 5)", 
    min_value=0, value=0
)

all_data = []

# =========================
# 2. Baca & gabungkan data
# =========================
if uploaded_files:
    for file in uploaded_files:
        file_name = file.name

        if file_name.endswith(".csv"):
            df = pd.read_csv(file, header=header_row)
            df = df.ffill()  # baris kosong sebelum header → isi baris sebelumnya

        elif file_name.endswith((".xlsx", ".ods")):
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet, header=header_row)
                df = df.ffill()  # baris kosong sebelum header → isi baris sebelumnya
                df["source_file"] = file_name
                df["source_sheet"] = sheet
                all_data.append(df)

        else:
            st.warning(f"Format file {file_name} tidak didukung")

        if file_name.endswith(".csv"):
            df["source_file"] = file_name
            df["source_sheet"] = "csv_file"
            all_data.append(df)

    if all_data:
        data_gabungan = pd.concat(all_data, ignore_index=True)

        # Ganti NaN setelah header → 0 (numerik), "kosong" (kategori)
        for col in data_gabungan.columns:
            if pd.api.types.is_numeric_dtype(data_gabungan[col]):
                data_gabungan[col] = data_gabungan[col].fillna(0)
            else:
                data_gabungan[col] = data_gabungan[col].fillna("kosong")

        st.subheader("2️⃣ Data Gabungan (setelah diproses)")
        st.dataframe(data_gabungan.head(50), use_container_width=True)
    else:
        st.warning("❌ Tidak ada data yang berhasil digabung.")
else:
    st.info("⬆️ Silakan unggah file terlebih dahulu untuk melihat data.")

# =========================
# 3. Filter Data
# =========================
if uploaded_files and all_data:
    st.subheader("3️⃣ Filter Data")

    filter_cols = st.multiselect("Pilih kolom untuk difilter", list(data_gabungan.columns))

    filtered = data_gabungan.copy()
    active_filters = {}

    for col in filter_cols:
        options = filtered[col].dropna().unique().tolist()
        selected = st.multiselect(
            f"Pilih nilai untuk `{col}`", 
            options, 
            key=f"val_{col}"
        )
        if selected:
            active_filters[col] = selected

    # Terapkan semua filter secara cumulative
    if active_filters:
        for col, selected in active_filters.items():
            filtered = filtered[filtered[col].astype(str).isin([str(s) for s in selected])]

    # Tampilkan hanya kolom yang difilter
    if filter_cols:
        st.dataframe(filtered[filter_cols].reset_index(drop=True), use_container_width=True)

    # Tombol unduh hasil penyaringan
    csv_buffer = io.StringIO()
    filtered.to_csv(csv_buffer, index=False)
    st.download_button(
        label="⬇️ Unduh Hasil Penyaringan (CSV)",
        data=csv_buffer.getvalue(),
        file_name="hasil_penyaringan.csv",
        mime="text/csv"
    )

    # =========================
    # 4. Visualisasi Data
    # =========================
    st.subheader("4️⃣ Visualisasi Data")

    x_axis = st.selectbox("Pilih kolom sumbu X", data_gabungan.columns)
    y_axis = st.selectbox("Pilih kolom sumbu Y", data_gabungan.columns)
    chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

    chart = None
    if chart_type == "Diagram Batang":
        chart = alt.Chart(filtered).mark_bar().encode(x=x_axis, y=y_axis, tooltip=list(filtered.columns))
    elif chart_type == "Diagram Garis":
        chart = alt.Chart(filtered).mark_line().encode(x=x_axis, y=y_axis, tooltip=list(filtered.columns))
    elif chart_type == "Diagram Sebar":
        chart = alt.Chart(filtered).mark_circle(size=60).encode(x=x_axis, y=y_axis, tooltip=list(filtered.columns))

    if chart:
        st.altair_chart(chart, use_container_width=True)
