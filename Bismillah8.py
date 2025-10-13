import streamlit as st
import pandas as pd
import os
import altair as alt
from io import BytesIO

# ======================
# Konfigurasi Tampilan (Tema)
# ======================
st.set_page_config(
    page_title="🌐 Web untuk Menggabungkan, Memfilter, dan Memvisualisasikan Data",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS Kustom
st.markdown("""
    <style>
    body {background-color: #f5f9ff; color: #0d1b2a;}
    .main {
        background-color: #ffffff;
        border-radius: 15px;
        padding: 20px 25px;
        box-shadow: 0px 0px 10px rgba(0, 60, 120, 0.1);
    }
    h1, h2, h3, h4 {color: #0d47a1; font-weight: 700;}
    div.stButton > button:first-child {
        background-color: #1976d2; color: white; border-radius: 10px; border: none;
        padding: 0.5rem 1.5rem; transition: 0.3s;
    }
    div.stButton > button:first-child:hover {
        background-color: #1565c0; transform: scale(1.03);
    }
    div[data-testid="stDownloadButton"] button {
        background-color: #0d47a1; color: white; border-radius: 8px; transition: 0.3s;
    }
    div[data-testid="stDownloadButton"] button:hover {
        background-color: #1565c0; transform: scale(1.02);
    }
    section[data-testid="stSidebar"] {background-color: #e3f2fd;}
    </style>
""", unsafe_allow_html=True)

# ======================
# Judul
# ======================
st.title("🌐 Web untuk Menggabungkan, Memfilter, dan Memvisualisasikan Data")
st.caption("Versi fleksibel dengan deteksi otomatis header, merge cell, dan visualisasi interaktif")

# ======================
# Upload atau Folder
# ======================
mode = st.radio("Pilih sumber data:", ["Upload File", "Pilih Folder"])
data_frames = []

# ======================
# Fungsi bantu
# ======================
def baca_sheet_dengan_deteksi(file, sheet_name, ext):
    try:
        if ext == "ods":
            df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None, engine="odf")
        else:
            df_raw = pd.read_excel(file, sheet_name=sheet_name, header=None, engine="openpyxl")

        header_row = df_raw.apply(lambda x: x.notna().sum(), axis=1).idxmax()
        if ext == "ods":
            df = pd.read_excel(file, sheet_name=sheet_name, header=header_row, engine="odf")
        else:
            df = pd.read_excel(file, sheet_name=sheet_name, header=header_row, engine="openpyxl")

        df.columns = (
            df.columns.astype(str)
            .str.strip()
            .str.lower()
            .str.replace(" ", "_")
            .str.replace(r"[^a-zA-Z0-9_]", "", regex=True)
        )
        df = df.ffill(axis=0)
        return df
    except Exception as e:
        st.warning(f"❗ Gagal membaca sheet {sheet_name}: {e}")
        return None

# ======================
# Mode Upload
# ======================
if mode == "Upload File":
    uploaded_files = st.file_uploader(
        "Upload file Excel/ODS (bisa banyak)",
        type=["xlsx", "xls", "ods"],
        accept_multiple_files=True
    )
    if uploaded_files:
        for uploaded_file in uploaded_files:
            ext = uploaded_file.name.split(".")[-1].lower()
            try:
                if ext == "ods":
                    sheets = pd.read_excel(uploaded_file, sheet_name=None, engine="odf")
                else:
                    sheets = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")

                for name in sheets.keys():
                    df = baca_sheet_dengan_deteksi(uploaded_file, name, ext)
                    if df is not None:
                        df["__SHEET__"] = name
                        df["__FILE__"] = uploaded_file.name
                        data_frames.append(df)
            except Exception as e:
                st.warning(f"Gagal membaca file {uploaded_file.name}: {e}")

# ======================
# Mode Folder
# ======================
elif mode == "Pilih Folder":
    folder = st.text_input("Masukkan path folder (isi file .xlsx/.ods)")
    if folder and os.path.isdir(folder):
        for fname in os.listdir(folder):
            if fname.endswith((".xlsx", ".xls", ".ods")):
                fpath = os.path.join(folder, fname)
                ext = fname.split(".")[-1].lower()
                try:
                    if ext == "ods":
                        sheets = pd.read_excel(fpath, sheet_name=None, engine="odf")
                    else:
                        sheets = pd.read_excel(fpath, sheet_name=None, engine="openpyxl")
                    for name in sheets.keys():
                        df = baca_sheet_dengan_deteksi(fpath, name, ext)
                        if df is not None:
                            df["__SHEET__"] = name
                            df["__FILE__"] = fname
                            data_frames.append(df)
                except Exception as e:
                    st.warning(f"Gagal membaca {fname}: {e}")

# ======================
# Gabungkan Data
# ======================
if data_frames:
    data_gabungan = pd.concat(data_frames, ignore_index=True).fillna("")
    st.subheader("📄 Data Gabungan")
    st.dataframe(data_gabungan)

    # ======================
    # Tombol Unduh
    # ======================
    st.subheader("💾 Unduh Data Gabungan")
    csv_data = data_gabungan.to_csv(index=False).encode("utf-8")
    st.download_button("📥 Unduh CSV", csv_data, file_name="data_gabungan.csv", mime="text/csv")

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        data_gabungan.to_excel(writer, index=False, sheet_name="Gabungan")
    st.download_button(
        label="📥 Unduh Excel (.xlsx)",
        data=buffer.getvalue(),
        file_name="data_gabungan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ======================
    # Filter Data
    # ======================
    st.subheader("🔍 Penyaringan Data")
    filter_columns = st.multiselect("Pilih kolom untuk filter", data_gabungan.columns)
    filtered_df = data_gabungan.copy()

    for kol in filter_columns:
        unique_vals = filtered_df[kol].dropna().unique().tolist()
        pilihan = st.multiselect(f"Filter nilai untuk {kol}", unique_vals, default=unique_vals)
        filtered_df = filtered_df[filtered_df[kol].isin(pilihan)]

    st.write("### Data Setelah Penyaringan")
    st.dataframe(filtered_df)

    # ======================
    # Transpose Otomatis (opsional)
    # ======================
    transpose_opt = st.checkbox("🔄 Tampilkan versi transpose sebelum visualisasi")
    if transpose_opt:
        filtered_df = filtered_df.T.reset_index()
        filtered_df.columns = ["Kolom"] + [f"Baris_{i}" for i in range(1, len(filtered_df.columns))]
        st.dataframe(filtered_df)

    # ======================
    # Visualisasi
    # ======================
    if not filtered_df.empty:
        st.subheader("📈 Visualisasi Data")
        all_cols = filtered_df.columns.tolist()
        if len(all_cols) >= 2:
            x_col = st.selectbox("Pilih kolom sumbu X", all_cols, key="x_col")
            y_col = st.selectbox("Pilih kolom sumbu Y", [c for c in all_cols if c != x_col], key="y_col")
            chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

            df_vis = filtered_df.dropna(subset=[x_col, y_col])

            try:
                tooltip_cols = [alt.Tooltip(str(c), type="nominal") for c in df_vis.columns]
                if chart_type == "Diagram Batang":
                    chart = alt.Chart(df_vis).mark_bar(color="#1976d2").encode(
                        x=alt.X(x_col, type="nominal"),
                        y=alt.Y(y_col, type="quantitative"),
                        tooltip=tooltip_cols
                    )
                elif chart_type == "Diagram Garis":
                    chart = alt.Chart(df_vis).mark_line(color="#0d47a1", point=True).encode(
                        x=alt.X(x_col, type="nominal"),
                        y=alt.Y(y_col, type="quantitative"),
                        tooltip=tooltip_cols
                    )
                else:
                    chart = alt.Chart(df_vis).mark_circle(size=70, color="#42a5f5").encode(
                        x=alt.X(x_col, type="quantitative"),
                        y=alt.Y(y_col, type="quantitative"),
                        tooltip=tooltip_cols
                    )
                st.altair_chart(chart, use_container_width=True)
            except Exception as e:
                st.warning(f"⚠️ Terjadi error saat membuat grafik: {e}")
else:
    st.info("💡 Silakan upload file Excel/ODS atau pilih folder untuk mulai.")
    
