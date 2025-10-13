import streamlit as st
import pandas as pd
import altair as alt
import os

# ======================
# Konfigurasi Tampilan
# ======================
st.set_page_config(
    page_title="📑 Baca Semua Sheet (Header Otomatis/Manual)",
    page_icon="📘",
    layout="wide",
)

st.title("📑 Baca File Excel/ODS per Sheet dengan Opsi Header Otomatis atau Manual")

# ======================
# Upload File
# ======================
uploaded_file = st.file_uploader("📂 Upload file Excel (.xlsx) atau ODS (.ods)", type=["xlsx", "xls", "ods"])
data_frames = []  # Inisialisasi list kosong untuk menampung data tiap sheet

if uploaded_file is not None:
    file_name = uploaded_file.name
    st.info(f"📄 File terdeteksi: **{file_name}**")

    # Deteksi ekstensi file
    ext = os.path.splitext(file_name)[-1].lower()

    try:
        # ======================
        # Ambil semua sheet
        # ======================
        if ext == ".ods":
            sheets = pd.read_excel(uploaded_file, sheet_name=None, engine="odf")
        else:
            sheets = pd.read_excel(uploaded_file, sheet_name=None)

        st.success(f"✅ Berhasil membaca {len(sheets)} sheet dari file ini!")

        for sheet_name, sheet_data in sheets.items():
            st.subheader(f"📄 Sheet: {sheet_name}")

            # ================
            # Pilihan Header
            # ================
            pilihan_header = st.radio(
                f"Pilih metode header untuk **{sheet_name}**:",
                ["Otomatis", "Manual"],
                key=f"pilihan_{sheet_name}"
            )

            if pilihan_header == "Otomatis":
                # Coba deteksi baris header otomatis
                header_row = None
                for i in range(min(10, len(sheet_data))):  # cek 10 baris pertama
                    non_null_ratio = sheet_data.iloc[i].notna().mean()
                    if non_null_ratio > 0.7:
                        header_row = i
                        break
                if header_row is None:
                    header_row = 0

                df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row, engine="odf" if ext == ".ods" else None)
                st.write(f"🧭 Header otomatis diambil dari baris ke-{header_row + 1}")

            else:
                # Manual: user pilih baris ke berapa untuk header
                max_row = len(sheet_data)
                header_row_manual = st.number_input(
                    f"Masukkan baris header untuk {sheet_name} (mulai dari 1):",
                    min_value=1,
                    max_value=max_row,
                    value=1,
                    key=f"header_manual_{sheet_name}"
                )
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row_manual - 1, engine="odf" if ext == ".ods" else None)
                st.write(f"🧭 Header manual diambil dari baris ke-{header_row_manual}")

            # Simpan dataframe ke daftar
            data_frames.append(df)

            # =================
            # Tampilkan data
            # =================
            st.dataframe(df.head())

    except Exception as e:
        st.error(f"❌ Gagal membaca file: {e}")

else:
    st.warning("⚠️ Silakan upload file Excel atau ODS terlebih dahulu.")

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

    # Pastikan semua kolom bertipe string dan aman untuk Altair
    filtered_df.columns = [str(c).strip().replace(":", "_").replace(" ", "_") for c in filtered_df.columns]

    all_cols = filtered_df.columns.tolist()

    x_col = st.selectbox("Pilih kolom sumbu X", all_cols)
    y_col = st.selectbox("Pilih kolom sumbu Y", [c for c in all_cols if c != x_col])
    chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

    df_filtered = filtered_df.dropna(subset=[x_col, y_col], how="any")

    try:
        tooltip_cols = [alt.Tooltip(str(c), type="nominal") for c in df_filtered.columns]

        # Tentukan tipe data secara dinamis
        x_type = "quantitative" if pd.api.types.is_numeric_dtype(df_filtered[x_col]) else "nominal"
        y_type = "quantitative" if pd.api.types.is_numeric_dtype(df_filtered[y_col]) else "nominal"

        if chart_type == "Diagram Batang":
            chart = alt.Chart(df_filtered).mark_bar(color="#1976d2").encode(
                x=alt.X(x_col, type=x_type),
                y=alt.Y(y_col, type=y_type),
                tooltip=tooltip_cols
            )
        elif chart_type == "Diagram Garis":
            chart = alt.Chart(df_filtered).mark_line(color="#0d47a1", point=True).encode(
                x=alt.X(x_col, type=x_type),
                y=alt.Y(y_col, type=y_type),
                tooltip=tooltip_cols
            )
        else:
            chart = alt.Chart(df_filtered).mark_circle(size=70, color="#42a5f5").encode(
                x=alt.X(x_col, type=x_type),
                y=alt.Y(y_col, type=y_type),
                tooltip=tooltip_cols
            )

        st.altair_chart(chart, use_container_width=True)

    except Exception as e:
        st.warning(f"⚠️ Terjadi error saat membuat grafik: {e}")

