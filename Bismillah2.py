import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Aplikasi Gabung, Filter, dan Visualisasi Data Multi-Sheet")

# Fungsi: bersihkan nama kolom supaya tidak ada "Unnamed"
def bersihkan_header(columns):
    new_cols = []
    for col in columns:
        if isinstance(col, tuple):  # multiindex header
            col = [c for c in col if str(c) != "nan"]
            name = " - ".join(map(str, col)) if col else "data kosong"
        else:
            if str(col).startswith("Unnamed"):
                name = "data kosong"
            else:
                name = str(col)
        new_cols.append(name.strip())
    return new_cols

# Fungsi baca file
def baca_file(file):
    ext = file.name.split(".")[-1].lower()
    all_dfs = []

    if ext in ["xlsx", "ods"]:
        xls = pd.ExcelFile(file, engine="openpyxl" if ext=="xlsx" else "odf")
        for sheet in xls.sheet_names:
            try:
                # baca semua sheet tanpa header dulu
                df_preview = pd.read_excel(file, sheet_name=sheet, header=None,
                                           engine="openpyxl" if ext=="xlsx" else "odf")
                # cari baris dengan isi terbanyak → jadi kandidat header
                filled_counts = df_preview.notna().sum(axis=1)
                header_row = filled_counts.idxmax()

                # baca ulang dengan header di baris terbaik
                df_sheet = pd.read_excel(
                    file, sheet_name=sheet, header=header_row,
                    engine="openpyxl" if ext=="xlsx" else "odf"
                )

                # bersihkan nama kolom
                df_sheet.columns = bersihkan_header(df_sheet.columns)

                # drop baris kosong
                df_sheet = df_sheet.dropna(how="all")

                # isi NaN
                for col in df_sheet.columns:
                    if pd.api.types.is_numeric_dtype(df_sheet[col]):
                        df_sheet[col] = df_sheet[col].fillna(0)
                    else:
                        df_sheet[col] = df_sheet[col].fillna("data kosong")

                # simpan sumber
                df_sheet["Sumber_File"] = file.name
                df_sheet["Sumber_Sheet"] = sheet

                all_dfs.append(df_sheet)
            except Exception as e:
                st.warning(f"Gagal membaca sheet {sheet} dari {file.name}: {e}")
    elif ext == "csv":
        df_csv = pd.read_csv(file)
        df_csv.columns = bersihkan_header(df_csv.columns)
        df_csv["Sumber_File"] = file.name
        df_csv["Sumber_Sheet"] = "CSV"
        all_dfs.append(df_csv)
    else:
        st.warning(f"Format {file.name} tidak didukung")

    return all_dfs


# Upload file
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

        # Filter
        filter_columns = st.multiselect("Pilih kolom yang ingin difilter & ditampilkan", all_data.columns.tolist())
        filtered_df = all_data.copy()

        if filter_columns:
            for col in filter_columns:
                unique_vals = filtered_df[col].dropna().unique().tolist()
                selected_vals = st.multiselect(f"Pilih nilai untuk kolom {col}", unique_vals)
                if selected_vals:
                    filtered_df = filtered_df[filtered_df[col].isin(selected_vals)]

            filtered_df = filtered_df[filter_columns]

        st.subheader("📌 Hasil Penyaringan")
        st.dataframe(filtered_df)

        # Visualisasi
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
            else:
                chart = alt.Chart(filtered_df).mark_circle(size=60).encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=all_cols
                )

            st.altair_chart(chart, use_container_width=True)

        # Unduh hasil
        st.subheader("💾 Unduh Hasil")

        if not filtered_df.empty:
            csv = filtered_df.to_csv(index=False).encode("utf-8")
            st.download_button("Unduh sebagai CSV", csv, "hasil_filter.csv", "text/csv")

            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
            st.download_button("Unduh sebagai Excel (.xlsx)", excel_buffer.getvalue(),
                               "hasil_filter.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            
