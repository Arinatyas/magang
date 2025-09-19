import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Aplikasi Gabung, Filter, dan Visualisasi Data Multi-Sheet")

# Fungsi baca file dengan pembersihan
def baca_file(file):
    ext = file.name.split(".")[-1].lower()
    all_dfs = []

    if ext in ["xlsx", "ods"]:
        xls = pd.ExcelFile(file, engine="openpyxl" if ext=="xlsx" else "odf")
        for sheet in xls.sheet_names:
            try:
                # cek header baris mana yang paling penuh
                df_preview = pd.read_excel(
                    file, sheet_name=sheet, nrows=3, header=None,
                    engine="openpyxl" if ext=="xlsx" else "odf"
                )
                filled_counts = df_preview.notna().sum(axis=1)
                max_filled_idx = filled_counts.idxmax()

                # baca full sheet pakai header yang benar
                df_sheet = pd.read_excel(
                    file, sheet_name=sheet, header=max_filled_idx,
                    engine="openpyxl" if ext=="xlsx" else "odf"
                )

                # --- Hapus kolom Unnamed ---
                df_sheet = df_sheet.loc[:, ~df_sheet.columns.astype(str).str.contains("^Unnamed")]

                # --- Bersihkan baris kosong / tidak lengkap ---
                df_sheet = df_sheet.dropna(how="all")       # hapus baris kosong total
                df_sheet = df_sheet.dropna(thresh=2)        # hapus baris yg isi < 2 kolom

                # --- Hapus baris yang duplikat header ---
                df_sheet = df_sheet[~df_sheet.apply(
                    lambda row: row.astype(str).tolist() == df_sheet.columns.astype(str).tolist(),
                    axis=1
                )]

                # --- Isi nilai kosong ---
                df_sheet = df_sheet.replace(r'^\s*$', pd.NA, regex=True)
                for col in df_sheet.columns:
                    if pd.api.types.is_numeric_dtype(df_sheet[col]):
                        df_sheet[col] = pd.to_numeric(df_sheet[col], errors="coerce").fillna(0)
                    else:
                        df_sheet[col] = (
                            df_sheet[col].astype(str)
                            .replace("nan", "data kosong")
                            .replace("<NA>", "data kosong")
                            .fillna("data kosong")
                        )

                # Tambahkan sumber file & sheet
                df_sheet["Sumber_File"] = file.name
                df_sheet["Sumber_Sheet"] = sheet
                all_dfs.append(df_sheet)

            except Exception as e:
                st.warning(f"Gagal membaca sheet {sheet} dari {file.name}: {e}")

    elif ext == "csv":
        df_csv = pd.read_csv(file)

        # --- Hapus kolom Unnamed ---
        df_csv = df_csv.loc[:, ~df_csv.columns.astype(str).str.contains("^Unnamed")]

        # isi kosong
        df_csv = df_csv.replace(r'^\s*$', pd.NA, regex=True)
        for col in df_csv.columns:
            if pd.api.types.is_numeric_dtype(df_csv[col]):
                df_csv[col] = pd.to_numeric(df_csv[col], errors="coerce").fillna(0)
            else:
                df_csv[col] = (
                    df_csv[col].astype(str)
                    .replace("nan", "data kosong")
                    .replace("<NA>", "data kosong")
                    .fillna("data kosong")
                )

        df_csv["Sumber_File"] = file.name
        df_csv["Sumber_Sheet"] = "CSV"
        all_dfs.append(df_csv)
    else:
        st.warning(f"Format {file.name} tidak didukung")

    return all_dfs


# Upload beberapa file
uploaded_files = st.file_uploader(
    "Unggah beberapa file (Excel/CSV/ODS) sekaligus",
    type=["xlsx", "csv", "ods"],
    accept_multiple_files=True
)

if uploaded_files:
    semua_data = []
    for file in uploaded_files:
        semua_data.extend(baca_file(file))

    if semua_data:
        all_data = pd.concat(semua_data, ignore_index=True)
        st.subheader("🔎 Data Gabungan")
        st.dataframe(all_data.head())

        # Filter sheet
        sheet_options = all_data["Sumber_Sheet"].unique().tolist()
        pilih_sheet = st.multiselect("Filter berdasarkan Sheet", sheet_options)
        filtered_df = all_data.copy()
        if pilih_sheet:
            filtered_df = filtered_df[filtered_df["Sumber_Sheet"].isin(pilih_sheet)]

        # Filter kolom dinamis
        filter_columns = st.multiselect("Pilih kolom yang ingin difilter & ditampilkan", filtered_df.columns.tolist())
        if filter_columns:
            for col in filter_columns:
                unique_vals = filtered_df[col].dropna().unique().tolist()
                selected_vals = st.multiselect(f"Pilih nilai untuk kolom {col}", unique_vals)
                if selected_vals:
                    filtered_df = filtered_df[filtered_df[col].isin(selected_vals)]
            filtered_df = filtered_df[filter_columns]

        st.subheader("📌 Hasil Penyaringan")
        st.dataframe(filtered_df)

        # =====================
        # Visualisasi
        # =====================
        if filter_columns and not filtered_df.empty:
            st.subheader("📈 Visualisasi Data")
            all_cols = filtered_df.columns.tolist()

            x_axis = st.selectbox("Pilih kolom sumbu X", all_cols)
            y_axis = st.selectbox("Pilih kolom sumbu Y", all_cols)
            chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

            def tipe_kolom(col):
                if pd.api.types.is_numeric_dtype(filtered_df[col]):
                    return "quantitative"
                elif pd.api.types.is_datetime64_any_dtype(filtered_df[col]):
                    return "temporal"
                else:
                    return "nominal"

            x_type = tipe_kolom(x_axis)
            y_type = tipe_kolom(y_axis)

            if chart_type == "Diagram Batang":
                chart = alt.Chart(filtered_df).mark_bar().encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=all_cols
                )
            elif chart_type == "Diagram Garis":
                chart = alt.Chart(filtered_df).mark_line(point=True).encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=all_cols
                )
            else:  # Diagram Sebar
                chart = alt.Chart(filtered_df).mark_circle(size=60).encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=all_cols
                )

            st.altair_chart(chart, use_container_width=True)

            # Unduh chart
            try:
                png = chart.properties(width=600, height=400).to_image(format="png")
                st.download_button("📥 Unduh Grafik PNG", data=png, file_name="visualisasi.png", mime="image/png")

                jpg = chart.properties(width=600, height=400).to_image(format="jpg")
                st.download_button("📥 Unduh Grafik JPG", data=jpg, file_name="visualisasi.jpg", mime="image/jpg")
            except Exception as e:
                st.warning("⚠️ Export PNG/JPG membutuhkan 'altair_saver' dan backend seperti selenium/cairosvg.")

        # =====================
        # Unduh hasil data
        # =====================
        st.subheader("💾 Unduh Hasil")
        if not filtered_df.empty:
            csv = filtered_df.to_csv(index=False).encode("utf-8")
            st.download_button("Unduh sebagai CSV", csv, "hasil_filter.csv", "text/csv")

            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
            st.download_button(
                "Unduh sebagai Excel (.xlsx)",
                excel_buffer.getvalue(),
                "hasil_filter.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            ods_buffer = BytesIO()
            with pd.ExcelWriter(ods_buffer, engine="odf") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
            st.download_button(
                "Unduh sebagai ODS (.ods)",
                ods_buffer.getvalue(),
                "hasil_filter.ods",
                "application/vnd.oasis.opendocument.spreadsheet"
            )
            
