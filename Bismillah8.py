import streamlit as st
import pandas as pd
import os
import altair as alt

# ======================
# Konfigurasi Tampilan
# ======================
st.set_page_config(
    page_title="📊 Gabung Data + Visualisasi",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS Tema Biru Putih Elegan
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
    .stRadio > label, .stSelectbox > label {
        font-weight: 600;
        color: #0d47a1;
    }
    .stDataFrame {
        border-radius: 12px !important;
        border: 1px solid #d0e3ff !important;
    }
    div[data-testid="stDownloadButton"] button {
        background-color: #0d47a1;
        color: white;
        border-radius: 8px;
        transition: 0.3s;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #1565c0;
        transform: scale(1.02);
    }
    section[data-testid="stSidebar"] {
        background-color: #e3f2fd;
    }
    </style>
""", unsafe_allow_html=True)

# ======================
# Judul
# ======================
st.title("📊 Gabung Data Excel/ODS + Visualisasi")
st.caption("Versi dengan deteksi header otomatis + opsi manual")

# ======================
# Fungsi Baca Data Otomatis
# ======================
def read_excel_auto_header(file, engine=None):
    """
    Membaca file Excel/ODS dengan deteksi header otomatis.
    Jika gagal, mengembalikan None agar bisa ganti manual.
    """
    try:
        df_test = pd.read_excel(file, engine=engine)
        if df_test.columns.isnull().any() or df_test.columns.duplicated().any():
            return None
        else:
            return df_test
    except Exception:
        return None

# ======================
# Upload atau Folder
# ======================
mode = st.radio("Pilih sumber data:", ["Upload File", "Pilih Folder"])
data_frames = []

if mode == "Upload File":
    uploaded_files = st.file_uploader(
        "Upload file Excel/ODS (bisa banyak)",
        type=["xlsx", "xls", "ods"],
        accept_multiple_files=True
    )

    if uploaded_files:
        for uploaded_file in uploaded_files:
            # Coba baca otomatis
            df_auto = None
            try:
                df_auto = read_excel_auto_header(uploaded_file, engine="openpyxl")
            except:
                df_auto = read_excel_auto_header(uploaded_file, engine="odf")

            if df_auto is not None:
                df = df_auto
            else:
                st.warning(f"⚠️ Tidak dapat mendeteksi header otomatis untuk: {uploaded_file.name}")
                # Tampilkan beberapa baris awal biar user bisa pilih header
                try:
                    preview = pd.read_excel(uploaded_file, header=None).head(5)
                    st.write(f"Pratinjau awal file **{uploaded_file.name}**:")
                    st.dataframe(preview)

                    header_option = st.selectbox(
                        f"Pilih baris header untuk {uploaded_file.name} (0 = tanpa header)",
                        [0, 1, 2, 3, 4],
                        help="Pilih baris yang terlihat berisi nama kolom"
                    )

                    if header_option == 0:
                        df = pd.read_excel(uploaded_file, header=None)
                    else:
                        df = pd.read_excel(uploaded_file, header=header_option - 1)
                except Exception as e:
                    st.error(f"Gagal membaca file: {uploaded_file.name}, error: {e}")
                    continue

            df["__FILE__"] = uploaded_file.name
            data_frames.append(df)

elif mode == "Pilih Folder":
    folder = st.text_input("Masukkan path folder (isi file .xlsx/.ods)")

    if folder and os.path.isdir(folder):
        for fname in os.listdir(folder):
            if fname.endswith((".xlsx", ".xls", ".ods")):
                fpath = os.path.join(folder, fname)
                df_auto = None
                try:
                    df_auto = read_excel_auto_header(fpath, engine="openpyxl")
                except:
                    df_auto = read_excel_auto_header(fpath, engine="odf")

                if df_auto is not None:
                    df = df_auto
                else:
                    st.warning(f"⚠️ Header otomatis gagal untuk file: {fname}")
                    preview = pd.read_excel(fpath, header=None).head(5)
                    st.write(f"Pratinjau awal file **{fname}**:")
                    st.dataframe(preview)

                    header_option = st.selectbox(
                        f"Pilih baris header untuk {fname} (0 = tanpa header)",
                        [0, 1, 2, 3, 4]
                    )

                    if header_option == 0:
                        df = pd.read_excel(fpath, header=None)
                    else:
                        df = pd.read_excel(fpath, header=header_option - 1)

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
                        
