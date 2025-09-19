import streamlit as st
import pandas as pd
import altair as alt
import matplotlib.pyplot as plt
from io import BytesIO

st.title("📊 Aplikasi Penggabungan & Visualisasi Data Excel")

# --- Upload File ---
uploaded_files = st.file_uploader("Unggah File Excel", type=["xlsx", "xls"], accept_multiple_files=True)

data_gabungan = pd.DataFrame()

if uploaded_files:
    dfs = []
    for file in uploaded_files:
        try:
            df = pd.read_excel(file, header=[0, 1])  # antisipasi header 2 baris
            # Jika ada merge header → gabungkan nama kolom
            df.columns = [' '.join([str(lv) for lv in col if str(lv) != 'nan']).strip() for col in df.columns.values]
        except Exception:
            df = pd.read_excel(file)  # fallback kalau tidak multi header
        dfs.append(df)

    # Gabungkan semua file
    data_gabungan = pd.concat(dfs, ignore_index=True)

    st.success(f"✅ {len(uploaded_files)} file berhasil digabung. Total baris: {len(data_gabungan)}")

    # --- Tampilkan Data ---
    st.subheader("📄 Data Gabungan")
    st.dataframe(data_gabungan)

    # --- Filter Opsional ---
    st.subheader("🔎 Filter Data (Opsional)")
    kolom_filter = st.selectbox("Pilih kolom untuk filter", ["(Tidak Ada)"] + list(data_gabungan.columns))
    if kolom_filter != "(Tidak Ada)":
        nilai_filter = st.multiselect("Pilih nilai filter", data_gabungan[kolom_filter].dropna().unique())
        if nilai_filter:
            data_penyaringan = data_gabungan[data_gabungan[kolom_filter].isin(nilai_filter)]
        else:
            data_penyaringan = data_gabungan
    else:
        data_penyaringan = data_gabungan

    st.write("📊 Data Hasil Penyaringan")
    st.dataframe(data_penyaringan)

    # --- Unduh Data ---
    st.subheader("⬇️ Unduh Data")
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        data_penyaringan.to_excel(writer, index=False, sheet_name="Data")
    st.download_button("📥 Unduh sebagai Excel (.xlsx)", data=excel_buffer.getvalue(),
                       file_name="data_gabungan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    ods_buffer = BytesIO()
    try:
        from odf.opendocument import OpenDocumentSpreadsheet
        from odf.table import Table, TableRow, TableCell
        from odf.text import P

        ods = OpenDocumentSpreadsheet()
        table = Table(name="Data")
        for row in data_penyaringan.itertuples(index=False, name=None):
            tr = TableRow()
            for val in row:
                tc = TableCell()
                tc.addElement(P(text=str(val)))
                tr.addElement(tc)
            table.addElement(tr)
        ods.spreadsheet.addElement(table)
        ods.save(ods_buffer)
        st.download_button("📥 Unduh sebagai ODS", data=ods_buffer.getvalue(),
                           file_name="data_gabungan.ods", mime="application/vnd.oasis.opendocument.spreadsheet")
    except Exception:
        st.warning("⚠️ Modul `odfpy` belum terinstal, unduhan ODS tidak tersedia.")

    # --- Visualisasi ---
    st.subheader("📊 Visualisasi Data")

    kolom_x = st.selectbox("Pilih Kolom X", data_penyaringan.columns)
    kolom_y = st.selectbox("Pilih Kolom Y", data_penyaringan.columns)

    tipe_x = st.radio("Tipe X", ["Kategori", "Numerik"],
                      index=0 if not pd.api.types.is_numeric_dtype(data_penyaringan[kolom_x]) else 1, horizontal=True)
    tipe_y = st.radio("Tipe Y", ["Kategori", "Numerik"],
                      index=0 if not pd.api.types.is_numeric_dtype(data_penyaringan[kolom_y]) else 1, horizontal=True)

    # Isi NaN
    df_vis = data_penyaringan.copy()
    for col in [kolom_x, kolom_y]:
        if pd.api.types.is_numeric_dtype(df_vis[col]):
            df_vis[col] = df_vis[col].fillna(0)
        else:
            df_vis[col] = df_vis[col].fillna("Tidak Ada")

    # --- Tentukan Grafik Altair ---
    if tipe_x == "Kategori" and tipe_y == "Kategori":
        chart = alt.Chart(df_vis).mark_bar().encode(
            x=kolom_x,
            y="count()",
            color=kolom_y
        )
    elif tipe_x == "Kategori" and tipe_y == "Numerik":
        chart = alt.Chart(df_vis).mark_bar().encode(
            x=kolom_x,
            y=f"mean({kolom_y})",
            color=kolom_x
        )
    elif tipe_x == "Numerik" and tipe_y == "Numerik":
        chart = alt.Chart(df_vis).mark_circle(size=60).encode(
            x=kolom_x,
            y=kolom_y,
            tooltip=[kolom_x, kolom_y]
        ).interactive()
    else:
        chart = alt.Chart(df_vis).mark_bar().encode(
            x=kolom_y,
            y=f"mean({kolom_x})",
            color=kolom_y
        )

    st.altair_chart(chart, use_container_width=True)
                
