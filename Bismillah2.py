import streamlit as st
import pandas as pd
import altair as alt
import io
from io import BytesIO

st.title("📊 Aplikasi Penggabungan & Penyaringan Data")

# Upload file
uploaded_files = st.file_uploader(
    "Upload file (bisa lebih dari satu, format: xlsx/csv/ods)",
    type=["xlsx", "csv", "ods"],
    accept_multiple_files=True
)

# Pilih header
header_input = st.number_input("Pilih baris header (mulai dari 0)", min_value=0, value=0, step=1)

data_gabungan = pd.DataFrame()

if uploaded_files:
    all_dfs = []

    for file in uploaded_files:
        filename = file.name.lower()

        try:
            if filename.endswith(".csv"):
                df = pd.read_csv(file, header=header_input)
                all_dfs.append(df)

            elif filename.endswith(".xlsx"):
                xls = pd.ExcelFile(file)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(file, sheet_name=sheet, header=header_input)
                    df["__sheetname__"] = sheet
                    df["__filename__"] = file.name
                    all_dfs.append(df)

            elif filename.endswith(".ods"):
                xls = pd.ExcelFile(file, engine="odf")
                for sheet in xls.sheet_names:
                    df = pd.read_excel(file, sheet_name=sheet, header=header_input, engine="odf")
                    df["__sheetname__"] = sheet
                    df["__filename__"] = file.name
                    all_dfs.append(df)

        except Exception as e:
            st.warning(f"Gagal membaca {file.name}: {e}")

    if all_dfs:
        data_gabungan = pd.concat(all_dfs, ignore_index=True)

        # Bersihkan nama kolom & atasi duplikat
        data_gabungan = data_gabungan.rename(columns=lambda x: str(x).strip())
        if data_gabungan.columns.duplicated().any():
            cols = pd.Series(data_gabungan.columns)
            for dup in cols[cols.duplicated()].unique():
                mask = cols == dup
                cols[mask] = [f"{dup}.{i}" if i > 0 else dup for i in range(mask.sum())]
            data_gabungan.columns = cols

        st.subheader("📂 Data Gabungan")
        st.dataframe(data_gabungan)

        # Tombol unduh CSV
        csv_gabungan = data_gabungan.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Unduh Data Gabungan (CSV)", csv_gabungan, "data_gabungan.csv", "text/csv")

        # Tombol unduh Excel
        excel_buffer = BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            data_gabungan.to_excel(writer, index=False, sheet_name="Gabungan")
        st.download_button("⬇️ Unduh Data Gabungan (Excel)", excel_buffer.getvalue(),
                           "data_gabungan.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Tombol unduh ODS
        ods_buffer = BytesIO()
        data_gabungan.to_excel(ods_buffer, engine="odf", index=False, sheet_name="Gabungan")
        st.download_button("⬇️ Unduh Data Gabungan (ODS)", ods_buffer.getvalue(),
                           "data_gabungan.ods", "application/vnd.oasis.opendocument.spreadsheet")

        # Penyaringan
        kolom_filter = st.multiselect("Pilih kolom untuk filter", data_gabungan.columns)

        if kolom_filter:
            data_penyaringan = data_gabungan[kolom_filter].copy()

            # Ganti nilai kosong
            for col in data_penyaringan.columns:
                if pd.api.types.is_numeric_dtype(data_penyaringan[col]):
                    data_penyaringan[col] = data_penyaringan[col].fillna(0)
                else:
                    data_penyaringan[col] = data_penyaringan[col].fillna("Tidak Ada")

            st.subheader("🔎 Data Hasil Penyaringan")
            st.dataframe(data_penyaringan)

            # Unduh CSV
            csv_penyaringan = data_penyaringan.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Unduh Data Penyaringan (CSV)", csv_penyaringan,
                               "data_penyaringan.csv", "text/csv")

            # Unduh Excel
            excel_buffer2 = BytesIO()
            with pd.ExcelWriter(excel_buffer2, engine="openpyxl") as writer:
                data_penyaringan.to_excel(writer, index=False, sheet_name="Penyaringan")
            st.download_button("⬇️ Unduh Data Penyaringan (Excel)", excel_buffer2.getvalue(),
                               "data_penyaringan.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # Unduh ODS
            ods_buffer2 = BytesIO()
            data_penyaringan.to_excel(ods_buffer2, engine="odf", index=False, sheet_name="Penyaringan")
            st.download_button("⬇️ Unduh Data Penyaringan (ODS)", ods_buffer2.getvalue(),
                               "data_penyaringan.ods", "application/vnd.oasis.opendocument.spreadsheet")

            # Visualisasi
            st.subheader("📊 Visualisasi Data")
            kolom_visual = st.selectbox("Pilih kolom untuk visualisasi", data_penyaringan.columns)

            if kolom_visual:
                df_vis = data_penyaringan.copy()

                # Pastikan tidak ada nilai kosong
                if pd.api.types.is_numeric_dtype(df_vis[kolom_visual]):
                    df_vis[kolom_visual] = df_vis[kolom_visual].fillna(0)
                    chart = alt.Chart(df_vis).mark_bar().encode(
                        x=alt.X(kolom_visual, bin=alt.Bin(maxbins=30)),
                        y='count()'
                    )
                else:
                    df_vis[kolom_visual] = df_vis[kolom_visual].fillna("Tidak Ada")
                    chart = alt.Chart(df_vis).mark_bar().encode(
                        x=kolom_visual,
                        y='count()'
                    )

                st.altair_chart(chart, use_container_width=True)

else:
    st.info("Silakan upload file terlebih dahulu.")
                                      
