import streamlit as st
import pandas as pd
import os
import altair as alt

st.set_page_config(page_title="📊 Web Gabung, Filter & Visualisasi Data", layout="wide")

st.title("📊 Web Gabung, Filter & Visualisasi Data")
st.caption("Mendukung Excel & ODS, fleksibel untuk struktur berbeda")

# Pilihan upload
mode = st.radio("Pilih sumber data:", ["Upload File", "Pilih Folder"])

data_frames = []

def read_file(filepath):
    ext = os.path.splitext(filepath)[-1].lower()
    engine = "openpyxl" if ext in [".xlsx", ".xls"] else "odf"
    try:
        sheets = pd.read_excel(filepath, sheet_name=None, header=None, engine=engine)
    except Exception as e:
        st.warning(f"Gagal membaca {filepath}: {e}")
        return []
    
    dfs = []
    for sheet_name, df in sheets.items():
        # Hilangkan kolom/baris kosong penuh
        df = df.dropna(how='all').dropna(axis=1, how='all')

        # Coba deteksi header otomatis
        try:
            header_row = df.notna().sum(axis=1).idxmax()
            df.columns = df.iloc[header_row].astype(str)
            df = df.iloc[header_row+1:].reset_index(drop=True)
        except Exception:
            # Jika gagal, pakai default index saja
            df.columns = [f"Kolom_{i}" for i in range(df.shape[1])]
            df = df.reset_index(drop=True)

        # Ganti kolom kosong dengan nama generik
        df.columns = [str(col).strip() if str(col).strip() != '' else f"Kolom_{i}" for i, col in enumerate(df.columns)]
        df.columns = [c if c.lower() != 'nan' else f"Kolom_{i}" for i, c in enumerate(df.columns)]
        
        # Tambahkan metadata
        df["__SHEET__"] = sheet_name
        df["__FILE__"] = os.path.basename(filepath)
        dfs.append(df)
    return dfs

# ===============================
# Mode Upload
# ===============================
if mode == "Upload File":
    uploaded_files = st.file_uploader("Upload Excel/ODS", type=["xlsx", "xls", "ods"], accept_multiple_files=True)
    if uploaded_files:
        for f in uploaded_files:
            data_frames.extend(read_file(f))

# ===============================
# Mode Folder
# ===============================
else:
    folder = st.text_input("Masukkan path folder")
    if folder and os.path.isdir(folder):
        for fname in os.listdir(folder):
            if fname.endswith((".xlsx", ".xls", ".ods")):
                fpath = os.path.join(folder, fname)
                data_frames.extend(read_file(fpath))

# ===============================
# Gabung Data
# ===============================
if data_frames:
    for i, df in enumerate(data_frames):
        df.columns = df.columns.astype(str).str.strip()
        df = df.loc[:, ~df.columns.duplicated()]
        data_frames[i] = df

    try:
        data_gabungan = pd.concat(data_frames, ignore_index=True)
    except Exception as e:
        st.error(f"Gagal menggabungkan data: {e}")
        st.stop()

    st.subheader("📄 Data Gabungan")
    st.dataframe(data_gabungan)

    # ===============================
    # Transpose Opsional
    # ===============================
    transpose_opt = st.checkbox("🔄 Transpose Data (baris <-> kolom)")
    if transpose_opt:
        data_gabungan = data_gabungan.transpose()
        st.dataframe(data_gabungan)

    # ===============================
    # Filter
    # ===============================
    st.subheader("🔍 Penyaringan Data")
    filter_columns = st.multiselect("Pilih kolom untuk filter", data_gabungan.columns)
    df_filtered = data_gabungan.copy()
    for kol in filter_columns:
        unique_vals = df_filtered[kol].dropna().unique().tolist()
        pilihan = st.multiselect(f"Pilih nilai {kol}", unique_vals)
        if pilihan:
            df_filtered = df_filtered[df_filtered[kol].isin(pilihan)]

    st.write("### Data Setelah Penyaringan")
    st.dataframe(df_filtered)

    # ===============================
    # Unduh hasil
    # ===============================
    out_excel = "hasil_gabungan.xlsx"
    out_ods = "hasil_gabungan.ods"

    try:
        df_filtered.to_excel(out_excel, index=False)
        df_filtered.to_excel(out_ods, index=False, engine="odf")
    except Exception as e:
        st.warning(f"Gagal menyimpan ke file: {e}")

    with open(out_excel, "rb") as f:
        st.download_button("📥 Unduh Excel", f, file_name=out_excel)
    with open(out_ods, "rb") as f:
        st.download_button("📥 Unduh ODS", f, file_name=out_ods)

    # ===============================
    # Visualisasi
    # ===============================
    st.subheader("📊 Visualisasi Data")
    if len(df_filtered.columns) >= 2:
        x_col = st.selectbox("Kolom X", df_filtered.columns)
        y_col = st.selectbox("Kolom Y", [c for c in df_filtered.columns if c != x_col])
        chart_type = st.radio("Jenis Grafik", ["Bar", "Line", "Scatter"])

        def convert(series):
            try:
                return pd.to_numeric(series, errors="coerce")
            except:
                return series.astype(str)
        
        df_filtered[y_col] = convert(df_filtered[y_col])
        df_filtered[x_col] = df_filtered[x_col].astype(str)

        if chart_type == "Bar":
            chart = alt.Chart(df_filtered).mark_bar(color="#1976d2").encode(x=x_col, y=y_col, tooltip=list(df_filtered.columns))
        elif chart_type == "Line":
            chart = alt.Chart(df_filtered).mark_line(point=True, color="#0d47a1").encode(x=x_col, y=y_col, tooltip=list(df_filtered.columns))
        else:
            chart = alt.Chart(df_filtered).mark_circle(size=80, color="#42a5f5").encode(x=x_col, y=y_col, tooltip=list(df_filtered.columns))

        st.altair_chart(chart, use_container_width=True)
    
