import streamlit as st
import pandas as pd
import os
import altair as alt

# ======================
# Konfigurasi Tampilan
# ======================
st.set_page_config(
    page_title="📊 Gabung Data Excel/ODS + Visualisasi",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS Tema Biru Putih Elegan
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
st.caption("Header otomatis/manual + nama sheet + filter + visualisasi")

# ======================
# Fungsi bantu deteksi header otomatis
# ======================
def detect_best_header_row(file, engine=None, max_check=5):
    """Deteksi baris header terbaik berdasarkan nilai unik & non-NaN terbanyak."""
    best_row = 0
    best_score = -1
    try:
        for i in range(max_check):
            df_temp = pd.read_excel(file, header=i, engine=engine)
            unique_score = len(set([c for c in df_temp.columns if pd.notna(c)]))
            nan_penalty = df_temp.columns.isnull().sum()
            score = unique_score - nan_penalty
            if score > best_score:
                best_score = score
                best_row = i
        df_final = pd.read_excel(file, header=best_row, engine=engine)
        return df_final, best_row
    except Exception:
        return None, None

# ======================
# Pilihan mode unggah
# ======================
mode = st.radio("Pilih sumber data:", ["Upload File", "Pilih Folder"])
header_mode = st.radio("Bagaimana membaca header?", ["Otomatis", "Manual"])

data_frames = []

# ======================
# Mode Upload File
# ======================
if mode == "Upload File":
    uploaded_files = st.file_uploader(
        "Upload file Excel/ODS (bisa banyak)",
        type=["xlsx", "xls", "ods"],
        accept_multiple_files=True
    )

    if uploaded_files:
        for uploaded_file in uploaded_files:
            st.markdown(f"### 📄 {uploaded_file.name}")
            try:
                # Baca semua sheet
                try:
                    sheets = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")
                except:
                    sheets = pd.read_excel(uploaded_file, sheet_name=None, engine="odf")
                
                for sheet_name, _ in sheets.items():
                    df = None
                    if header_mode == "Otomatis":
                        df, header_row = detect_best_header_row(uploaded_file, engine="openpyxl")
                        if df is not None:
                            st.success(f"✅ Header otomatis terdeteksi di baris ke-{header_row+1} sheet: {sheet_name}")
                        else:
                            st.warning("⚠️ Gagal mendeteksi header, pilih manual")
                            continue
                    else:
                        preview = pd.read_excel(uploaded_file, header=None, nrows=5)
                        st.write("Pratinjau 5 baris pertama:")
                        st.dataframe(preview)
                        header_row = st.selectbox(
                            f"Pilih baris header untuk {uploaded_file.name} sheet {sheet_name} (0 = tanpa header)",
                            list(range(0, 6))
                        )
                        if header_row == 0:
                            df = pd.read_excel(uploaded_file, header=None)
                        else:
                            df = pd.read_excel(uploaded_file, header=header_row - 1)

                    if df is None:
                        df = pd.read_excel(uploaded_file, header=0)

                    # Tambah info file & sheet
                    df["__FILE__"] = uploaded_file.name
                    df["__SHEET__"] = sheet_name
                    data_frames.append(df)

            except Exception as e:
                st.error(f"Gagal membaca file {uploaded_file.name}: {e}")

# ======================
# Mode Pilih Folder
# ======================
elif mode == "Pilih Folder":
    folder = st.text_input("Masukkan path folder (isi file .xlsx/.ods)")
    if folder and os.path.isdir(folder):
        for fname in os.listdir(folder):
            if fname.endswith((".xlsx", ".xls", ".ods")):
                fpath = os.path.join(folder, fname)
                st.markdown(f"### 📄 {fname}")
                try:
                    try:
                        sheets = pd.read_excel(fpath, sheet_name=None, engine="openpyxl")
                    except:
                        sheets = pd.read_excel(fpath, sheet_name=None, engine="odf")

                    for sheet_name, _ in sheets.items():
                        df = None
                        if header_mode == "Otomatis":
                            df, header_row = detect_best_header_row(fpath, engine="openpyxl")
                            if df is not None:
                                st.success(f"✅ Header otomatis terdeteksi di baris ke-{header_row+1} sheet: {sheet_name}")
                            else:
                                st.warning("⚠️ Gagal mendeteksi header, pilih manual")
                                continue
                        else:
                            preview = pd.read_excel(fpath, header=None, nrows=5)
                            st.write("Pratinjau 5 baris pertama:")
                            st.dataframe(preview)
                            header_row = st.selectbox(
                                f"Pilih baris header untuk {fname} sheet {sheet_name} (0 = tanpa header)",
                                list(range(0, 6))
                            )
                            if header_row == 0:
                                df = pd.read_excel(fpath, header=None)
                            else:
                                df = pd.read_excel(fpath, header=header_row - 1)

                        if df is None:
                            df = pd.read_excel(fpath, header=0)

                        df["__FILE__"] = fname
                        df["__SHEET__"] = sheet_name
                        data_frames.append(df)
                except Exception as e:
                    st.error(f"Gagal membaca file {fname}: {e}")

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

    # Bersihkan nama kolom agar aman
    filtered_df.columns = [str(c).strip().replace(":", "_").replace(" ", "_") for c in filtered_df.columns]
    all_cols = filtered_df.columns.tolist()

    x_col = st.selectbox("Pilih kolom kategori (sumbu X)", all_cols)
    y_col = st.selectbox("Pilih kolom numerik (sumbu Y)", [c for c in all_cols if c != x_col])
    chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang (Total)", "Diagram Garis (Total)", "Diagram Sebar"])

    df_filtered = filtered_df.dropna(subset=[x_col, y_col], how="any")

    try:
        # Tentukan tipe data
        x_type = "quantitative" if pd.api.types.is_numeric_dtype(df_filtered[x_col]) else "nominal"
        y_type = "quantitative" if pd.api.types.is_numeric_dtype(df_filtered[y_col]) else "quantitative"

        # Konversi nilai numerik
        df_filtered[y_col] = pd.to_numeric(df_filtered[y_col], errors="coerce")

        # Jika x bukan numerik → agregasi total per kategori
        if x_type == "nominal":
            df_vis = df_filtered.groupby(x_col, as_index=False)[y_col].sum()
        else:
            df_vis = df_filtered.copy()

        tooltip_cols = [alt.Tooltip(str(c), type="nominal") for c in df_vis.columns]

        # =====================
        # Jenis Grafik
        # =====================
        if chart_type == "Diagram Batang (Total)":
            chart = alt.Chart(df_vis).mark_bar(color="#1976d2").encode(
                x=alt.X(x_col, type=x_type, title=x_col),
                y=alt.Y(y_col, type=y_type, title=f"Total {y_col}"),
                tooltip=tooltip_cols
            )

        elif chart_type == "Diagram Garis (Total)":
            chart = alt.Chart(df_vis).mark_line(color="#0d47a1", point=True).encode(
                x=alt.X(x_col, type=x_type, title=x_col),
                y=alt.Y(y_col, type=y_type, title=f"Total {y_col}"),
                tooltip=tooltip_cols
            )

        else:  # Diagram Sebar
            chart = alt.Chart(df_vis).mark_circle(size=70, color="#42a5f5").encode(
                x=alt.X(x_col, type=x_type),
                y=alt.Y(y_col, type=y_type),
                tooltip=tooltip_cols
            )

        st.altair_chart(chart, use_container_width=True)
        st.caption("🔢 Nilai numerik ditampilkan sebagai total per kategori (agregasi sum).")

    except Exception as e:
        st.warning(f"⚠️ Terjadi error saat membuat grafik: {e}")
