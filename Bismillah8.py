import streamlit as st
import pandas as pd
import altair as alt
import numpy as np
import os

# ======================
# Konfigurasi Tampilan
# ======================
st.set_page_config(page_title="Web untuk Menggabungkan, Memfilter, dan Memvisualisasikan Data", layout="wide")
st.title("🌐 Web untuk Menggabungkan, Memfilter, dan Memvisualisasikan Data")

# ======================
# Fungsi bantu
# ======================
def deteksi_header_otomatis(df, max_row=10):
    skor_terbaik = -1
    baris_terbaik = 0
    for i in range(min(max_row, len(df))):
        baris = df.iloc[i]
        skor = 0
        for val in baris:
            if pd.notna(val):
                skor += 1
                if isinstance(val, str) and len(val) < 25:
                    skor += 1
        if skor > skor_terbaik:
            skor_terbaik = skor
            baris_terbaik = i
    return baris_terbaik

# ======================
# Upload file
# ======================
uploaded_files = st.file_uploader("📂 Upload file Excel/ODS (bisa banyak)", type=["xlsx", "xls", "ods"], accept_multiple_files=True)
data_frames = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        ext = os.path.splitext(uploaded_file.name)[1].lower()

        try:
            if ext in [".xlsx", ".xls"]:
                sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine="openpyxl")
            elif ext == ".ods":
                sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine="odf")
            else:
                st.warning(f"❗ Format file tidak dikenal: {uploaded_file.name}")
                continue
        except Exception as e:
            st.error(f"❌ Gagal membaca {uploaded_file.name}: {e}")
            continue

        for sheet_name, df_raw in sheets.items():
            if df_raw.empty:
                continue

            # Deteksi header otomatis
            header_row = deteksi_header_otomatis(df_raw)

            try:
                if ext in [".xlsx", ".xls"]:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row, engine="openpyxl")
                else:
                    df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row, engine="odf")
            except Exception as e:
                st.error(f"Gagal membaca sheet {sheet_name} dari {uploaded_file.name}: {e}")
                continue

            df["__SHEET__"] = sheet_name  # Nama kapal
            df["__FILE__"] = uploaded_file.name
            data_frames.append(df)

    if not data_frames:
        st.warning("⚠️ Tidak ada data valid yang bisa digabung.")
    else:
        data_gabungan = pd.concat(data_frames, ignore_index=True)
        st.success(f"✅ Berhasil menggabungkan {len(uploaded_files)} file dengan total {len(data_gabungan)} baris.")
        
        st.markdown("### 📄 Data Gabungan (Sebelum Filter)")
        st.dataframe(data_gabungan.head(50), use_container_width=True)

        # ======================
        # Pilih kolom manual
        # ======================
        st.markdown("### ⚙️ Pilih Kolom untuk Analisis")
        all_columns = data_gabungan.columns.tolist()
        kolom_kategori = st.multiselect("Pilih kolom kategori (contoh: Nama Pekerjaan, Jenis, dsb)", all_columns)
        kolom_numerik = st.multiselect("Pilih kolom numerik (akan dijumlahkan)", all_columns)
        
        if kolom_kategori and kolom_numerik:
            st.markdown("### 🔄 Proses Gabungan dan Agregasi Otomatis")

            grouped = data_gabungan.groupby(kolom_kategori, dropna=False)[kolom_numerik].sum().reset_index()

            if "__SHEET__" in data_gabungan.columns:
                grouped["Nama Kapal"] = data_gabungan["__SHEET__"]

            # Transpose sebelum visualisasi
            df_transposed = grouped.set_index(kolom_kategori).T.reset_index()
            df_transposed.rename(columns={"index": "Kategori"}, inplace=True)

            st.markdown("## 📋 Hasil Gabungan (Sudah Ditotal & Transpose)")
            st.dataframe(df_transposed, use_container_width=True)

            # ======================
            # Filter fleksibel
            # ======================
            st.markdown("### 🔍 Filter Data")
            filter_cols = st.multiselect("Pilih kolom untuk difilter", df_transposed.columns)
            filtered_df = df_transposed.copy()

            for col in filter_cols:
                unique_vals = filtered_df[col].dropna().unique().tolist()
                selected_vals = st.multiselect(f"Filter {col}:", unique_vals)
                if selected_vals:
                    filtered_df = filtered_df[filtered_df[col].isin(selected_vals)]

            st.markdown("### 📊 Data Setelah Filter")
            st.dataframe(filtered_df, use_container_width=True)

            # ======================
            # Visualisasi
            # ======================
            st.markdown("## 📈 Visualisasi Data")
            all_cols = filtered_df.columns.tolist()
            if len(all_cols) >= 2:
                x_col = st.selectbox("Sumbu X", all_cols)
                y_col = st.selectbox("Sumbu Y", [c for c in all_cols if c != x_col])
                chart_type = st.radio("Jenis Grafik", ["Bar", "Line", "Scatter"], horizontal=True)

                def detect_type(series):
                    try:
                        series.astype(float)
                        return "quantitative"
                    except:
                        return "nominal"

                x_type = detect_type(filtered_df[x_col])
                y_type = detect_type(filtered_df[y_col])

                tooltip_cols = [alt.Tooltip(str(c), type="nominal") for c in filtered_df.columns]

                if chart_type == "Bar":
                    chart = alt.Chart(filtered_df).mark_bar().encode(
                        x=alt.X(x_col, type=x_type),
                        y=alt.Y(y_col, type=y_type),
                        tooltip=tooltip_cols
                    )
                elif chart_type == "Line":
                    chart = alt.Chart(filtered_df).mark_line(point=True).encode(
                        x=alt.X(x_col, type=x_type),
                        y=alt.Y(y_col, type=y_type),
                        tooltip=tooltip_cols
                    )
                else:
                    chart = alt.Chart(filtered_df).mark_circle(size=70).encode(
                        x=alt.X(x_col, type=x_type),
                        y=alt.Y(y_col, type=y_type),
                        tooltip=tooltip_cols
                    )

                st.altair_chart(chart, use_container_width=True)
            else:
                st.warning("⚠️ Kolom terlalu sedikit untuk membuat grafik.")

            # ======================
            # Unduh hasil akhir (Excel + ODS)
            # ======================
            st.markdown("## 💾 Unduh Hasil Gabungan")

            out_excel = "hasil_gabungan.xlsx"
            out_ods = "hasil_gabungan.ods"

            try:
                filtered_df.to_excel(out_excel, index=False, engine="openpyxl")
                filtered_df.to_excel(out_ods, index=False, engine="odf")

                with open(out_excel, "rb") as f:
                    st.download_button("📥 Unduh Excel (.xlsx)", f, file_name=out_excel)

                with open(out_ods, "rb") as f:
                    st.download_button("📥 Unduh ODS (.ods)", f, file_name=out_ods)
            except Exception as e:
                st.error(f"Gagal menyimpan hasil gabungan: {e}")

        else:
            st.info("Pilih setidaknya satu kolom kategori dan satu kolom numerik untuk melanjutkan.")
                                       
