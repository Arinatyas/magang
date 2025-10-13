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

# CSS Kustom untuk tema biru putih elegan
st.markdown("""
    <style>
    body { background-color: #f5f9ff; color: #0d1b2a; }
    .main { background-color: #ffffff; border-radius: 15px; padding: 20px 25px; box-shadow: 0px 0px 10px rgba(0, 60, 120, 0.1); }
    h1, h2, h3, h4 { color: #0d47a1; font-weight: 700; }
    div.stButton > button:first-child {
        background-color: #1976d2; color: white; border-radius: 10px; border: none; padding: 0.5rem 1.5rem; transition: 0.3s;
    }
    div.stButton > button:first-child:hover { background-color: #1565c0; transform: scale(1.03); }
    .stRadio > label, .stSelectbox > label { font-weight: 600; color: #0d47a1; }
    .stDataFrame { border-radius: 12px !important; border: 1px solid #d0e3ff !important; }
    div[data-testid="stDownloadButton"] button {
        background-color: #0d47a1; color: white; border-radius: 8px; transition: 0.3s;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #1565c0; transform: scale(1.02);
    }
    section[data-testid="stSidebar"] { background-color: #e3f2fd; }
    </style>
""", unsafe_allow_html=True)

# ======================
# Judul
# ======================
st.title("📊 Gabung Data Excel/ODS + Visualisasi")
st.caption("Versi dengan deteksi header otomatis + opsi manual")

# ======================
# Fungsi Deteksi Header Otomatis
# ======================
def detect_header_row(df, max_rows=10):
    """
    Mendeteksi baris header paling mungkin berdasarkan skor kombinasi:
    - Banyak sel tidak kosong
    - Banyak nilai unik
    - Banyak teks pendek
    """
    best_score = -1
    best_row = 0

    for i in range(min(max_rows, len(df))):
        row = df.iloc[i]
        non_empty = row.notna().sum()
        unique_vals = row.nunique()
        short_texts = sum(isinstance(v, str) and len(v) <= 20 for v in row)

        score = (non_empty * 0.4) + (unique_vals * 0.4) + (short_texts * 0.2)
        if score > best_score:
            best_score = score
            best_row = i

    return best_row

# ======================
# Pilihan sumber data
# ======================
mode = st.radio("Pilih sumber data:", ["Upload File", "Pilih Folder"])
header_mode = st.radio("Metode pembacaan header:", ["Otomatis", "Manual"])

data_frames = []

# ======================
# Fungsi baca file dengan deteksi header
# ======================
def read_file_with_header(file_path_or_obj, file_name, header_mode):
    try:
        # baca tanpa header dulu
        df_raw = pd.read_excel(file_path_or_obj, header=None, engine="openpyxl")
    except:
        df_raw = pd.read_excel(file_path_or_obj, header=None, engine="odf")

    st.markdown(f"### 📄 {file_name}")
    st.write("Pratinjau data awal:")
    st.dataframe(df_raw.head(10))

    if header_mode == "Otomatis":
        best_row = detect_header_row(df_raw)
        st.info(f"🧠 Deteksi otomatis: baris ke-{best_row + 1} kemungkinan besar adalah header.")
        df = pd.read_excel(file_path_or_obj, header=best_row)
        st.success(f"✅ Menggunakan header dari baris ke-{best_row + 1}")
    else:
        # Mode manual
        st.write("📋 Pilih baris yang akan dijadikan header:")
        header_row = st.number_input(f"Pilih baris header untuk {file_name}", min_value=0, max_value=10, value=0, step=1)
        use_header = st.button(f"Gunakan baris ke-{header_row + 1} sebagai header untuk {file_name}")

        if use_header:
            df = pd.read_excel(file_path_or_obj, header=header_row)
            st.success(f"✅ Menggunakan baris ke-{header_row + 1} sebagai header.")
        else:
            df = pd.read_excel(file_path_or_obj, header=None)
            st.warning("⚠️ Belum memilih header — menggunakan tanpa header sementara.")

    df["__FILE__"] = file_name
    return df

# ======================
# Upload File
# ======================
if mode == "Upload File":
    uploaded_files = st.file_uploader(
        "Upload file Excel/ODS (bisa banyak)",
        type=["xlsx", "xls", "ods"],
        accept_multiple_files=True
    )

    if uploaded_files:
        for uploaded_file in uploaded_files:
            df = read_file_with_header(uploaded_file, uploaded_file.name, header_mode)
            data_frames.append(df)

# ======================
# Pilih Folder
# ======================
elif mode == "Pilih Folder":
    folder = st.text_input("Masukkan path folder (isi file .xlsx/.ods)")
    if folder and os.path.isdir(folder):
        for fname in os.listdir(folder):
            if fname.endswith((".xlsx", ".xls", ".ods")):
                fpath = os.path.join(folder, fname)
                df = read_file_with_header(fpath, fname, header_mode)
                data_frames.append(df)

# ======================
# Gabungkan Data
# ======================
if data_frames:
    data_gabungan = pd.concat(data_frames, ignore_index=True)

    st.subheader("📄 Data Gabungan")
    st.dataframe(data_gabungan)


# ======================
# Mode Folder
# ======================
elif mode == "Pilih Folder":
    folder = st.text_input("Masukkan path folder (isi file .xlsx/.ods)")
    if folder and os.path.isdir(folder):
        for fname in os.listdir(folder):
            if fname.endswith((".xlsx", ".xls", ".ods")):
                fpath = os.path.join(folder, fname)
                st.markdown(f"### 📄 {fname}")

                df = None
                header_row = None

                if header_mode == "Otomatis":
                    try:
                        df, header_row = detect_best_header_row(fpath, engine="openpyxl")
                    except:
                        df, header_row = detect_best_header_row(fpath, engine="odf")

                    if df is not None:
                        st.success(f"✅ Header otomatis terdeteksi di baris ke-{header_row+1}")
                    else:
                        st.warning(f"⚠️ Gagal mendeteksi header untuk {fname}")
                        continue

                else:  # Manual
                    preview = pd.read_excel(fpath, header=None, nrows=5)
                    st.write("Pratinjau 5 baris pertama:")
                    st.dataframe(preview)

                    header_row = st.selectbox(
                        f"Pilih baris header untuk {fname} (0 = tanpa header)",
                        list(range(0, 6))
                    )

                    if header_row == 0:
                        df = pd.read_excel(fpath, header=None)
                    else:
                        df = pd.read_excel(fpath, header=header_row - 1)

                if df is not None:
                    df["__FILE__"] = fname
                    data_frames.append(df)

# ======================
# Gabungkan Data
# ======================
if data_frames:
    data_gabungan = pd.concat(data_frames, ignore_index=True)

    st.subheader("📄 Data Gabungan")
    st.dataframe(data_gabungan)

    # ======================
    # Filter Data
    # ======================
    st.subheader("🔍 Penyaringan Data")
    filter_columns = st.multiselect("Pilih kolom untuk filter", data_gabungan.columns)

    filtered_df = data_gabungan.copy()
    tampilkan_kolom = []

    for kol in filter_columns:
        unique_vals = filtered_df[kol].dropna().unique().tolist()
        pilihan = st.multiselect(f"Pilih nilai untuk {kol}", unique_vals)
        tampilkan_kolom.append(kol)
        if pilihan:
            filtered_df = filtered_df[filtered_df[kol].isin(pilihan)]

    if tampilkan_kolom:
        filtered_df = filtered_df[tampilkan_kolom]

    st.write("### Data Setelah Penyaringan")
    st.dataframe(filtered_df)

    # ======================
    # Unduh Data
    # ======================
    st.subheader("💾 Unduh Data")
    out_excel = "data_gabungan.xlsx"
    out_ods = "data_gabungan.ods"
    filtered_df.to_excel(out_excel, index=False, engine="openpyxl")
    filtered_df.to_excel(out_ods, index=False, engine="odf")

    with open(out_excel, "rb") as f:
        st.download_button("📥 Unduh Excel (.xlsx)", f, file_name=out_excel)

    with open(out_ods, "rb") as f:
        st.download_button("📥 Unduh ODS (.ods)", f, file_name=out_ods)

    # ======================
    # Visualisasi
    # ======================
    if not filtered_df.empty and len(filtered_df.columns) > 1:
        st.subheader("📈 Visualisasi Data")

        filtered_df.columns = [str(c) for c in filtered_df.columns]
        all_cols = filtered_df.columns.tolist()

        x_col = st.selectbox("Pilih kolom sumbu X", all_cols)
        y_col = st.selectbox("Pilih kolom sumbu Y", [c for c in all_cols if c != x_col])
        chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

        df_filtered = filtered_df.dropna(subset=[x_col, y_col], how="any")

        try:
            tooltip_cols = [alt.Tooltip(str(c), type="nominal") for c in df_filtered.columns]
            if chart_type == "Diagram Batang":
                chart = alt.Chart(df_filtered).mark_bar(color="#1976d2").encode(
                    x=alt.X(x_col, type="nominal"),
                    y=alt.Y(y_col, type="quantitative"),
                    tooltip=tooltip_cols
                )
            elif chart_type == "Diagram Garis":
                chart = alt.Chart(df_filtered).mark_line(color="#0d47a1", point=True).encode(
                    x=alt.X(x_col, type="nominal"),
                    y=alt.Y(y_col, type="quantitative"),
                    tooltip=tooltip_cols
                )
            else:
                chart = alt.Chart(df_filtered).mark_circle(size=70, color="#42a5f5").encode(
                    x=alt.X(x_col, type="quantitative"),
                    y=alt.Y(y_col, type="quantitative"),
                    tooltip=tooltip_cols
                )
            st.altair_chart(chart, use_container_width=True)
        except Exception as e:
            st.warning(f"⚠️ Terjadi error saat membuat grafik: {e}")
        
