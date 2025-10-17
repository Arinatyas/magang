import streamlit as st
import pandas as pd
import os
import altair as alt
from io import BytesIO
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P

# ======================
# Konfigurasi Tampilan
# ======================
st.set_page_config(
    page_title="📊 Gabung Data Excel/ODS + Visualisasi",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ======================
# CSS Tema Biru Putih
# ======================
st.markdown("""
    <style>
    body { background-color: #f5f9ff; color: #0d1b2a; }
    .main { background-color: #ffffff; border-radius: 15px; padding: 20px 25px; box-shadow: 0px 0px 10px rgba(0, 60, 120, 0.1); }
    h1, h2, h3, h4 { color: #0d47a1; font-weight: 700; }
    div.stButton > button:first-child {
        background-color: #1976d2; color: white; border-radius: 10px; border: none; padding: 0.5rem 1.5rem; transition: 0.3s;
    }
    div.stButton > button:first-child:hover { background-color: #1565c0; transform: scale(1.03); }
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
# Fungsi membaca sheet
# ======================
def read_sheet_with_header_option(file_path, sheet_name=None, header_mode="Otomatis", max_preview=9):
    df, header_row = None, None
    preview_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=max_preview)

    if header_mode == "Otomatis":
        best_score = -1
        for i in range(max_preview):
            row_values = preview_df.iloc[i].astype(str)
            unique_count = len(set(row_values)) - row_values.isin(['nan', 'None']).sum()
            if unique_count > best_score:
                best_score = unique_count
                header_row = i
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
            st.success(f"✅ Header otomatis terdeteksi di baris ke-{header_row+1} sheet {sheet_name or ''}")
        except Exception as e:
            st.warning(f"⚠️ Gagal membaca sheet {sheet_name or ''}: {e}")
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    else:
        st.write("Pratinjau 9 baris pertama:")
        st.dataframe(preview_df)
        header_row = st.selectbox(
            f"Pilih baris header untuk sheet {sheet_name or ''} (0 = tanpa header)",
            list(range(0, max_preview))
        )
        try:
            if header_row == 0:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row - 1)
        except Exception as e:
            st.warning(f"⚠️ Gagal membaca sheet {sheet_name or ''}: {e}")
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    return df, header_row


# ======================
# Fungsi baca file otomatis (Excel, ODS, CSV)
# ======================
def load_sheets_any_format(uploaded_file):
    file_name = uploaded_file.name.lower()
    try:
        if file_name.endswith((".xlsx", ".xls")):
            try:
                return pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")
            except:
                uploaded_file.seek(0)
                return pd.read_excel(uploaded_file, sheet_name=None, engine="odf")
        elif file_name.endswith(".ods"):
            try:
                return pd.read_excel(uploaded_file, sheet_name=None, engine="odf")
            except:
                uploaded_file.seek(0)
                return pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")
        elif file_name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
            return {"Sheet1": df}
        else:
            uploaded_file.seek(0)
            return pd.read_excel(uploaded_file, sheet_name=None)
    except Exception as e:
        st.error(f"❌ Gagal membaca {uploaded_file.name}: {e}")
        return None


# ======================
# Upload File
# ======================
data_frames = []
header_mode = st.radio("Bagaimana membaca header?", ["Otomatis", "Manual"])
uploaded_files = st.file_uploader(
    "Upload file Excel/ODS (bisa banyak)",
    type=["xlsx", "xls", "ods", "csv"],
    accept_multiple_files=True
)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.markdown(f"### 📄 {uploaded_file.name}")
        sheets = load_sheets_any_format(uploaded_file)
        if sheets:
            for sheet_name, _ in sheets.items():
                df, header_row = read_sheet_with_header_option(uploaded_file, sheet_name, header_mode)
                if df is not None:
                    df["__FILE__"] = uploaded_file.name
                    df["__SHEET__"] = sheet_name
                    data_frames.append(df)

# ======================
# Gabung Data + Filter
# ======================
if data_frames:
    data_gabungan = pd.concat(data_frames, ignore_index=True)
    st.subheader("📄 Data Gabungan")
    st.dataframe(data_gabungan)

    st.subheader("🔍 Penyaringan Data")
    filter_columns = st.multiselect("Pilih kolom untuk filter", data_gabungan.columns)
    filtered_df = data_gabungan.copy()

    if filter_columns:
        for kol in filter_columns:
            unique_vals = filtered_df[kol].dropna().unique().tolist()
            pilihan = st.multiselect(f"Pilih nilai untuk {kol}", unique_vals)
            if pilihan:
                filtered_df = filtered_df[filtered_df[kol].isin(pilihan)]

    st.write("### Data Setelah Penyaringan")
    st.dataframe(filtered_df)

# ======================
# Unduh Data Gabungan
# ======================
if "filtered_df" in locals() and not filtered_df.empty:
    st.subheader("💾 Unduh Data")
    try:
        out_excel = "data_gabungan.xlsx"
        out_ods = "data_gabungan.ods"
        filtered_df.to_excel(out_excel, index=False, engine="openpyxl")
        try:
            filtered_df.to_excel(out_ods, index=False, engine="odf")
        except:
            out_ods = None

        with open(out_excel, "rb") as f:
            st.download_button("📥 Unduh Excel (.xlsx)", f, file_name=out_excel)
        if out_ods:
            with open(out_ods, "rb") as f:
                st.download_button("📥 Unduh ODS (.ods)", f, file_name=out_ods)
    except Exception as e:
        st.error(f"❌ Terjadi error saat proses unduh: {e}")

# ======================
# Visualisasi Data + Unduh Agregasi
# ======================
if "filtered_df" in locals() and not filtered_df.empty:
    st.subheader("📈 Visualisasi Data")

    filtered_df.columns = [str(c).strip().replace(":", "_").replace(" ", "_") for c in filtered_df.columns]
    all_cols = filtered_df.columns.tolist()

    x_col = st.selectbox("Pilih kolom X (kategori atau numerik)", all_cols)
    y_col = st.selectbox("Pilih kolom Y (kategori atau numerik)", [c for c in all_cols if c != x_col])
    chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

    df_vis = filtered_df.dropna(subset=[x_col, y_col], how="any").copy()
    df_vis[x_col] = df_vis[x_col].replace("", "Kosong")
    df_vis[y_col] = df_vis[y_col].replace("", "Kosong")

    try:
        x_is_num = pd.api.types.is_numeric_dtype(df_vis[x_col])
        y_is_num = pd.api.types.is_numeric_dtype(df_vis[y_col])

        if not x_is_num and not y_is_num:
            df_agg = df_vis.groupby([x_col, y_col], dropna=False).size().reset_index(name="Jumlah")
            y_field = "Jumlah"
            st.write("### 🔢 Jumlah Kombinasi per Kategori")
        elif not x_is_num and y_is_num:
            df_agg = df_vis.groupby(x_col, dropna=False)[y_col].sum().reset_index()
            y_field = y_col
            st.write("### 🔢 Total per Kategori (Sum)")
        elif x_is_num and not y_is_num:
            df_agg = df_vis.groupby([x_col, y_col], dropna=False).size().reset_index(name="Jumlah")
            y_field = "Jumlah"
            st.write("### 🔢 Jumlah per Nilai Numerik")
        else:
            df_agg = df_vis[[x_col, y_col]].copy()
            y_field = y_col
            st.write("### 🔢 Data Numerik (tanpa agregasi)")

        st.dataframe(df_agg)

        # Unduh hasil agregasi
        csv_agg = df_agg.to_csv(index=False).encode('utf-8')
        st.download_button("💾 Unduh CSV (.csv)", csv_agg, file_name="hasil_agregasi.csv")

        excel_buf = BytesIO()
        with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
            df_agg.to_excel(writer, index=False, sheet_name="Hasil_Agregasi")
        st.download_button("📘 Unduh Excel (.xlsx)", excel_buf.getvalue(), file_name="hasil_agregasi.xlsx")

        ods_doc = OpenDocumentSpreadsheet()
        table = Table(name="Hasil_Agregasi")
        header_row = TableRow()
        for col in df_agg.columns:
            cell = TableCell(); cell.addElement(P(text=str(col))); header_row.addElement(cell)
        table.addElement(header_row)
        for _, row in df_agg.iterrows():
            tr = TableRow()
            for val in row:
                cell = TableCell(); cell.addElement(P(text=str(val))); tr.addElement(cell)
            table.addElement(tr)
        ods_doc.spreadsheet.addElement(table)
        ods_buf = BytesIO(); ods_doc.save(ods_buf)
        st.download_button("📗 Unduh ODS (.ods)", ods_buf.getvalue(), file_name="hasil_agregasi.ods")

        # Visualisasi
        tooltip_cols = [alt.Tooltip(str(c), type="nominal") for c in df_agg.columns]
        if chart_type == "Diagram Batang":
            chart = alt.Chart(df_agg).mark_bar(color="#1976d2").encode(
                x=alt.X(x_col, title=x_col),
                y=alt.Y(y_field, title=y_field),
                tooltip=tooltip_cols)
        elif chart_type == "Diagram Garis":
            chart = alt.Chart(df_agg).mark_line(point=True, color="#0d47a1").encode(
                x=alt.X(x_col, title=x_col),
                y=alt.Y(y_field, title=y_field),
                tooltip=tooltip_cols)
        else:
            chart = alt.Chart(df_agg).mark_circle(size=70, color="#42a5f5").encode(
                x=alt.X(x_col, title=x_col),
                y=alt.Y(y_field, title=y_field),
                tooltip=tooltip_cols)

        st.altair_chart(chart, use_container_width=True)
        st.caption("📊 Visualisasi otomatis menyesuaikan jenis data kategori atau numerik.")
    except Exception as e:
        st.warning(f"⚠️ Terjadi error saat membuat grafik: {e}")
                                    
