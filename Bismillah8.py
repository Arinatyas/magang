import streamlit as st
import pandas as pd
import os
import altair as alt

# ======================
# Konfigurasi Tampilan (Tema)
# ======================
st.set_page_config(
    page_title="📊 Gabung Data + Visualisasi",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS Kustom
st.markdown("""
    <style>
    body {
        background-color: #f5f9ff;
        color: #0d1b2a;
    }
    .main {
        background-color: #ffffff;
        border-radius: 15px;
        padding: 20px 25px;
        box-shadow: 0px 0px 10px rgba(0, 60, 120, 0.1);
    }
    h1, h2, h3, h4 {
        color: #0d47a1;
        font-weight: 700;
    }
    div.stButton > button:first-child {
        background-color: #1976d2;
        color: white;
        border-radius: 10px;
        border: none;
        padding: 0.5rem 1.5rem;
        transition: 0.3s;
    }
    div.stButton > button:first-child:hover {
        background-color: #1565c0;
        transform: scale(1.03);
    }
    .stDataFrame {
        border-radius: 12px !important;
        border: 1px solid #d0e3ff !important;
    }
    </style>
""", unsafe_allow_html=True)

# ======================
# Judul
# ======================
st.title("📊 Gabung Data Excel/ODS + Visualisasi")
st.caption("Versi biru–putih elegan dengan fitur group otomatis & transpose")

# ======================
# Upload atau Folder
# ======================
mode = st.radio("Pilih sumber data:", ["Upload File", "Pilih Folder"])
data_frames = []

if mode == "Upload File":
    uploaded_files = st.file_uploader("Upload file Excel/ODS", type=["xlsx", "xls", "ods"], accept_multiple_files=True)
    if uploaded_files:
        for uploaded_file in uploaded_files:
            try:
                sheets = pd.read_excel(uploaded_file, sheet_name=None, header=0, engine="openpyxl")
            except:
                sheets = pd.read_excel(uploaded_file, sheet_name=None, header=0, engine="odf")
            for name, df in sheets.items():
                df["__SHEET__"] = name
                df["__FILE__"] = uploaded_file.name
                data_frames.append(df)

elif mode == "Pilih Folder":
    folder = st.text_input("Masukkan path folder (isi file .xlsx/.ods)")
    if folder and os.path.isdir(folder):
        for fname in os.listdir(folder):
            if fname.endswith((".xlsx", ".xls", ".ods")):
                fpath = os.path.join(folder, fname)
                try:
                    sheets = pd.read_excel(fpath, sheet_name=None, header=0, engine="openpyxl")
                except:
                    sheets = pd.read_excel(fpath, sheet_name=None, header=0, engine="odf")
                for name, df in sheets.items():
                    df["__SHEET__"] = name
                    df["__FILE__"] = fname
                    data_frames.append(df)

# ======================
# Gabung Data
# ======================
if data_frames:
    data_gabungan = pd.concat(data_frames, ignore_index=True)
    st.subheader("📄 Data Gabungan")
    st.dataframe(data_gabungan)

    # ======================
    # Penyaringan Data
    # ======================
    st.subheader("🔍 Penyaringan Data")
    filter_columns = st.multiselect("Pilih kolom untuk filter", data_gabungan.columns)

    filtered_df = data_gabungan.copy()
    for kol in filter_columns:
        unique_vals = filtered_df[kol].dropna().unique().tolist()
        pilihan = st.multiselect(f"Pilih nilai untuk {kol}", unique_vals)
        if pilihan:
            filtered_df = filtered_df[filtered_df[kol].isin(pilihan)]

    # ======================
    # Grouping Otomatis (nama kapal / pekerjaan dll)
    # ======================
    # Deteksi kolom numerik & non-numerik
    num_cols = filtered_df.select_dtypes(include=["number"]).columns.tolist()
    non_num_cols = [c for c in filtered_df.columns if c not in num_cols and c not in ["__SHEET__", "__FILE__"]]

    if non_num_cols and num_cols:
        st.info("🔢 Data berisi kolom teks & angka — otomatis dikelompokkan dan ditotal berdasarkan kolom teks.")
        grouped_df = filtered_df.groupby(non_num_cols, dropna=False)[num_cols].sum(numeric_only=True).reset_index()
    else:
        grouped_df = filtered_df.copy()

    st.write("### 🧾 Data Setelah Penyaringan & Grouping Otomatis")
    st.dataframe(grouped_df, use_container_width=True)

    # ======================
    # Tombol Transpose (pindah ke sini)
    # ======================
    st.write("### 🔄 Transformasi Data")
    transpose_option = st.toggle("Transpose Data (baris jadi kolom, kolom jadi baris)")

    if transpose_option:
        grouped_df = grouped_df.transpose().reset_index()
        grouped_df.columns = [str(c) for c in grouped_df.columns]

    # ======================
    # Unduh Data
    # ======================
    st.subheader("💾 Unduh Data")
    out_excel = "data_hasil.xlsx"
    out_ods = "data_hasil.ods"
    grouped_df.to_excel(out_excel, index=False, engine="openpyxl")
    grouped_df.to_excel(out_ods, index=False, engine="odf")

    with open(out_excel, "rb") as f:
        st.download_button("📥 Unduh Excel (.xlsx)", f, file_name=out_excel)
    with open(out_ods, "rb") as f:
        st.download_button("📥 Unduh ODS (.ods)", f, file_name=out_ods)

    # ======================
    # Visualisasi Data
    # ======================
    if not grouped_df.empty:
        st.subheader("📈 Visualisasi Data")

        grouped_df.columns = [str(c) for c in grouped_df.columns]
        all_cols = grouped_df.columns.tolist()
        if len(all_cols) < 2:
            st.warning("⚠️ Data terlalu sedikit untuk divisualisasikan.")
        else:
            x_col = st.selectbox("Pilih kolom X", all_cols, key="x_col")
            y_col = st.selectbox("Pilih kolom Y", [c for c in all_cols if c != x_col], key="y_col")
            chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])
            posisi = st.radio("Posisi tampilan:", ["📄 Tabel kiri – Grafik kanan", "📊 Grafik kiri – Tabel kanan"], horizontal=True)

            df_filtered = grouped_df.copy()

            def detect_type(series):
                try:
                    pd.to_numeric(series)
                    return "quantitative"
                except:
                    return "nominal"

            x_type = detect_type(df_filtered[x_col])
            y_type = detect_type(df_filtered[y_col])

            # Pastikan angka dibaca numerik
            df_filtered[y_col] = pd.to_numeric(df_filtered[y_col], errors="coerce")

            if chart_type == "Diagram Batang":
                chart = alt.Chart(df_filtered).mark_bar(color="#1976d2").encode(
                    x=alt.X(x_col, type=x_type),
                    y=alt.Y(y_col, type="quantitative"),
                    tooltip=list(df_filtered.columns)
                )
            elif chart_type == "Diagram Garis":
                chart = alt.Chart(df_filtered).mark_line(color="#0d47a1", point=True).encode(
                    x=alt.X(x_col, type=x_type),
                    y=alt.Y(y_col, type="quantitative"),
                    tooltip=list(df_filtered.columns)
                )
            else:
                chart = alt.Chart(df_filtered).mark_circle(size=70, color="#42a5f5").encode(
                    x=alt.X(x_col, type=x_type),
                    y=alt.Y(y_col, type="quantitative"),
                    tooltip=list(df_filtered.columns)
                )

            col1, col2 = st.columns(2)
            if posisi == "📄 Tabel kiri – Grafik kanan":
                with col1:
                    st.dataframe(df_filtered, use_container_width=True)
                with col2:
                    st.altair_chart(chart, use_container_width=True)
            else:
                with col1:
                    st.altair_chart(chart, use_container_width=True)
                with col2:
                    st.dataframe(df_filtered, use_container_width=True)
        
