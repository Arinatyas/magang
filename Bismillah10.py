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
# Fungsi baca sheet dengan opsi header otomatis/manual
# ======================
def read_sheet_with_header_option(file_path, sheet_name=None, header_mode="Otomatis", max_preview=9):
    """
    Membaca sheet Excel/ODS dengan opsi header otomatis atau manual.
    """
    df = None
    header_row = None

    # Preview untuk manual
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
            st.success(f"✅ Header otomatis terdeteksi di baris ke-{header_row+1} sheet {sheet_name if sheet_name else ''}")
        except:
            st.warning(f"⚠️ Gagal membaca sheet {sheet_name if sheet_name else ''} dengan header otomatis")
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    else:
        st.write("Pratinjau 9 baris pertama:")
        st.dataframe(preview_df)
        header_row = st.selectbox(
            f"Pilih baris header untuk sheet {sheet_name if sheet_name else ''} (0 = tanpa header)",
            list(range(0, max_preview))
        )
        try:
            if header_row == 0:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row-1)
        except:
            st.warning(f"⚠️ Gagal membaca sheet {sheet_name if sheet_name else ''}")
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    return df, header_row

# ======================
# Pilihan mode unggah
# ======================
mode = st.radio("Pilih sumber data:", ["Upload File"])
header_mode = st.radio("Bagaimana membaca header?", ["Otomatis", "Manual"])
data_frames = []



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

    # Jika tidak ada filter, tetap buat filtered_df agar tidak error
    filtered_df = data_gabungan.copy()
    tampilkan_kolom = []

    if filter_columns:
        for kol in filter_columns:
            unique_vals = filtered_df[kol].dropna().unique().tolist()
            pilihan = st.multiselect(f"Pilih nilai untuk {kol}", unique_vals)
            tampilkan_kolom.append(kol)
            if pilihan:
                filtered_df = filtered_df[filtered_df[kol].isin(pilihan)]

        # Jika user hanya ingin melihat kolom yang difilter
        if tampilkan_kolom:
            filtered_df = filtered_df[tampilkan_kolom]

    st.write("### Data Setelah Penyaringan")
    st.dataframe(filtered_df)



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
# Mode Upload File
# ======================
def load_sheets_any_format(uploaded_file):
    """
    Membaca file Excel/ODS/CSV secara otomatis dan memperbaiki kesalahan format.
    """
    file_name = uploaded_file.name.lower()
    try:
        if file_name.endswith((".xlsx", ".xls")):
            try:
                return pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")
            except Exception as e:
                st.warning(f"⚠️ File {uploaded_file.name} bukan Excel ZIP valid ({e}), mencoba format ODS...")
                uploaded_file.seek(0)
                return pd.read_excel(uploaded_file, sheet_name=None, engine="odf")

        elif file_name.endswith(".ods"):
            try:
                return pd.read_excel(uploaded_file, sheet_name=None, engine="odf")
            except Exception as e:
                st.warning(f"⚠️ File {uploaded_file.name} gagal dibaca sebagai ODS ({e}), mencoba Excel...")
                uploaded_file.seek(0)
                return pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")

        elif file_name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
            return {"Sheet1": df}

        else:
            st.warning(f"⚠️ Ekstensi file {file_name} tidak dikenal. Mencoba semua cara...")
            uploaded_file.seek(0)
            try:
                return pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")
            except:
                uploaded_file.seek(0)
                return pd.read_excel(uploaded_file, sheet_name=None, engine="odf")

    except Exception as e:
        st.error(f"❌ Gagal membaca {uploaded_file.name}: {e}")
        return None


# ======================
# Proses Upload File
# ======================
if mode == "Upload File":
    uploaded_files = st.file_uploader(
        "Upload file Excel/ODS (bisa banyak)",
        type=["xlsx", "xls", "ods", "csv"],
        accept_multiple_files=True
    )

    if uploaded_files:
        for uploaded_file in uploaded_files:
            st.markdown(f"### 📄 {uploaded_file.name}")

            # Gunakan fungsi auto-format
            sheets = load_sheets_any_format(uploaded_file)
            if sheets is None:
                continue

            for sheet_name, _ in sheets.items():
                df, header_row = read_sheet_with_header_option(uploaded_file, sheet_name, header_mode)
                if df is not None:
                    df["__FILE__"] = uploaded_file.name
                    df["__SHEET__"] = sheet_name
                    data_frames.append(df)

# ======================
# Unduh Data
# ======================
st.subheader("💾 Unduh Data")

# Pastikan filtered_df sudah ada
if 'filtered_df' not in locals() or filtered_df is None:
    if 'data_gabungan' in locals():
        filtered_df = data_gabungan.copy()
    else:
        st.warning("⚠️ Tidak ada data untuk diunduh.")
        st.stop()

# Simpan Excel dan ODS
out_excel = "data_gabungan.xlsx"
out_ods = "data_gabungan.ods"

try:
    filtered_df.to_excel(out_excel, index=False, engine="openpyxl")
    filtered_df.to_excel(out_ods, index=False, engine="odf")

    with open(out_excel, "rb") as f:
        st.download_button("📥 Unduh Excel (.xlsx)", f, file_name=out_excel)
    with open(out_ods, "rb") as f:
        st.download_button("📥 Unduh ODS (.ods)", f, file_name=out_ods)
except Exception as e:
    st.error(f"❌ Gagal membuat file unduhan: {e}")

# ======================
# ======================
# VISUALISASI DATA (FIXED + DOWNLOAD EXCEL/ODS)
# ======================
from io import BytesIO
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P

if data_frames:
    if 'filtered_df' not in locals():
        filtered_df = data_gabungan.copy()
    # Pastikan kolom identifikasi file dan sheet ada
    if "__FILE__" not in data_gabungan.columns:
        data_gabungan["__FILE__"] = "Tidak diketahui"
    if "__SHEET__" not in data_gabungan.columns:
        data_gabungan["__SHEET__"] = "Tidak diketahui"
    if not filtered_df.empty and len(filtered_df.columns) > 1:
        st.subheader("📈 Visualisasi Data")

        # Bersihkan nama kolom
        filtered_df.columns = [
            str(c).strip().replace(":", "_").replace(" ", "_") for c in filtered_df.columns
        ]
        all_cols = filtered_df.columns.tolist()

        # Pilihan kolom X dan Y
        x_col = st.selectbox("Pilih kolom X (kategori atau numerik)", all_cols)
        y_col = st.selectbox("Pilih kolom Y (kategori atau numerik)", [c for c in all_cols if c != x_col])
        chart_type = st.radio(
            "Pilih jenis grafik",
            ["Diagram Batang", "Diagram Garis", "Diagram Sebar"]
        )

        # Pastikan kolom tidak kosong
        df_vis = filtered_df.dropna(subset=[x_col, y_col], how="any").copy()

        # Ganti string kosong jadi "Kosong"
        df_vis[x_col] = df_vis[x_col].replace("", "Kosong")
        df_vis[y_col] = df_vis[y_col].replace("", "Kosong")

        try:
            # Deteksi tipe data
            x_is_num = pd.api.types.is_numeric_dtype(df_vis[x_col])
            y_is_num = pd.api.types.is_numeric_dtype(df_vis[y_col])

            # Agregasi fleksibel
            if not x_is_num and not y_is_num:
                # Kedua kolom kategori → hitung jumlah kombinasi
                df_agg = df_vis.groupby([x_col, y_col], dropna=False).size().reset_index(name="Jumlah")
                y_field = "Jumlah"
                st.write("### 🔢 Jumlah Kombinasi per Kategori")
            elif not x_is_num and y_is_num:
                # X kategori, Y numerik → jumlahkan
                df_agg = df_vis.groupby(x_col, dropna=False)[y_col].sum().reset_index()
                y_field = y_col
                st.write("### 🔢 Total per Kategori (Sum)")
            elif x_is_num and not y_is_num:
                # X numerik, Y kategori → hitung jumlah per kategori
                df_agg = df_vis.groupby([x_col, y_col], dropna=False).size().reset_index(name="Jumlah")
                y_field = "Jumlah"
                st.write("### 🔢 Jumlah per Nilai Numerik")
            else:
                # Dua-duanya numerik → tampilkan tanpa agregasi
                df_agg = df_vis[[x_col, y_col]].copy()
                y_field = y_col
                st.write("### 🔢 Data Numerik (tanpa agregasi)")

            # Tampilkan hasil agregasi
            st.dataframe(df_agg)

            # ========== DOWNLOAD HASIL AGREGASI ==========
            # CSV
            csv_agg = df_agg.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="💾 Unduh CSV (.csv)",
                data=csv_agg,
                file_name="hasil_agregasi.csv",
                mime="text/csv"
            )

            # Excel
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_agg.to_excel(writer, index=False, sheet_name="Hasil_Agregasi")
            st.download_button(
                label="📘 Unduh Excel (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name="hasil_agregasi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ODS
            ods_doc = OpenDocumentSpreadsheet()
            table = Table(name="Hasil_Agregasi")
            header_row = TableRow()
            for col in df_agg.columns:
                cell = TableCell()
                cell.addElement(P(text=str(col)))
                header_row.addElement(cell)
            table.addElement(header_row)
            for _, row in df_agg.iterrows():
                tr = TableRow()
                for val in row:
                    cell = TableCell()
                    cell.addElement(P(text=str(val)))
                    tr.addElement(cell)
                table.addElement(tr)
            ods_doc.spreadsheet.addElement(table)
            ods_buf = BytesIO()
            ods_doc.save(ods_buf)
            st.download_button(
                label="📗 Unduh ODS (.ods)",
                data=ods_buf.getvalue(),
                file_name="hasil_agregasi.ods",
                mime="application/vnd.oasis.opendocument.spreadsheet"
            )

            # ========== VISUALISASI ==========
            tooltip_cols = [alt.Tooltip(str(c), type="nominal") for c in df_agg.columns]
            if chart_type == "Diagram Batang":
                chart = alt.Chart(df_agg).mark_bar(color="#1976d2").encode(
                    x=alt.X(x_col, title=x_col),
                    y=alt.Y(y_field, title=y_field),
                    tooltip=tooltip_cols
                )
            elif chart_type == "Diagram Garis":
                chart = alt.Chart(df_agg).mark_line(point=True, color="#0d47a1").encode(
                    x=alt.X(x_col, title=x_col),
                    y=alt.Y(y_field, title=y_field),
                    tooltip=tooltip_cols
                )
            else:
                chart = alt.Chart(df_agg).mark_circle(size=70, color="#42a5f5").encode(
                    x=alt.X(x_col, title=x_col),
                    y=alt.Y(y_field, title=y_field),
                    tooltip=tooltip_cols
                )

            st.altair_chart(chart, use_container_width=True)
            st.caption("📊 Visualisasi otomatis menyesuaikan jenis data kategori atau numerik.")

        except Exception as e:
            st.warning(f"⚠️ Terjadi error saat membuat grafik: {e}")
