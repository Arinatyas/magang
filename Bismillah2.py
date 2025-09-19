import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO

# ========================
# Fungsi bantu
# ========================

# Buat nama kolom unik (jika duplikat)
def make_unique(columns):
    seen = {}
    new_cols = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            new_cols.append(col)
    return new_cols

# Baca file dengan header yang dipilih
def baca_file(file, header_row):
    ext = file.name.split(".")[-1].lower()
    df = None

    try:
        if ext in ["xls", "xlsx"]:
            df = pd.read_excel(file, header=header_row, engine="openpyxl")
        elif ext == "ods":
            df = pd.read_excel(file, header=header_row, engine="odf")
        elif ext == "csv":
            df = pd.read_csv(file, header=header_row)
        else:
            st.warning(f"Format file {ext} tidak didukung")
            return []
    except Exception as e:
        st.warning(f"Gagal membaca {file.name}: {e}")
        return []

    # Hapus baris kosong total
    df = df.dropna(how="all")

    # Isi missing values
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].fillna(0)
        else:
            df[col] = df[col].fillna("data kosong")

    # Tambah metadata
    df["Sumber_File"] = file.name

    return [df]

# Tentukan tipe kolom (untuk Altair)
def tipe_kolom(col, df):
    if pd.api.types.is_numeric_dtype(df[col]):
        return "quantitative"
    elif pd.api.types.is_datetime64_any_dtype(df[col]):
        return "temporal"
    else:
        return "nominal"


# ========================
# Streamlit App
# ========================

st.set_page_config(layout="wide")
st.title("📊 Aplikasi Gabung, Filter, dan Visualisasi Data Multi-Sheet")

# Upload file
uploaded_files = st.file_uploader(
    "Unggah beberapa file (Excel/CSV/ODS) sekaligus",
    type=["xlsx", "csv", "ods"],
    accept_multiple_files=True
)

# Pilih header row
header_row = st.number_input(
    "Pilih baris header (mulai dari 0, misalnya baris ke-6 berarti 5)",
    min_value=0, step=1, value=0
)

if uploaded_files:
    semua_data = []
    for file in uploaded_files:
        semua_data.extend(baca_file(file, header_row))

    if semua_data:
        all_data = pd.concat(semua_data, ignore_index=True)

        # Pastikan nama kolom unik
        all_data.columns = make_unique(all_data.columns)

        st.subheader("🔎 Data Gabungan")
        st.dataframe(all_data.head())

        # ========================
        # Filter
        # ========================
        filter_columns = st.multiselect(
            "Pilih kolom yang ingin difilter & ditampilkan", all_data.columns.tolist()
        )
        filtered_df = all_data.copy()

        if filter_columns:
            for col in filter_columns:
                unique_vals = filtered_df[col].dropna().unique().tolist()
                selected_vals = st.multiselect(f"Pilih nilai untuk kolom {col}", unique_vals)
                if selected_vals:
                    filtered_df = filtered_df[filtered_df[col].isin(selected_vals)]

            # tampilkan hanya kolom yang difilter
            filtered_df = filtered_df[filter_columns]

        st.subheader("📌 Hasil Penyaringan")
        st.dataframe(filtered_df)

        # ========================
        # Visualisasi
        # ========================
        if filter_columns and not filtered_df.empty:
            st.subheader("📈 Visualisasi Data")
            all_cols = filtered_df.columns.tolist()

            # Pilih X dan Y
            x_axis = st.selectbox("Pilih kolom sumbu X", all_cols)
            y_axis = st.selectbox("Pilih kolom sumbu Y", all_cols)

            chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

            x_type = tipe_kolom(x_axis, filtered_df)
            y_type = tipe_kolom(y_axis, filtered_df)

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

            # ========================
            # Unduh visualisasi (SVG)
            # ========================
            st.subheader("💾 Unduh Visualisasi")
            try:
                svg = chart.to_svg()
                st.download_button(
                    "Unduh Grafik sebagai SVG",
                    svg,
                    "visualisasi.svg",
                    "image/svg+xml"
                )
            except Exception as e:
                st.warning(f"Gagal mengekspor visualisasi: {e}")

        # ========================
        # Unduh hasil
        # ========================
        st.subheader("💾 Unduh Data")

        if not filtered_df.empty:
            # CSV
            csv = filtered_df.to_csv(index=False).encode("utf-8")
            st.download_button("Unduh sebagai CSV", csv, "hasil_filter.csv", "text/csv")

            # Excel
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
            st.download_button(
                "Unduh sebagai Excel (.xlsx)", excel_buffer.getvalue(),
                "hasil_filter.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ODS
            ods_buffer = BytesIO()
            with pd.ExcelWriter(ods_buffer, engine="odf") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
            st.download_button(
                "Unduh sebagai ODS (.ods)", ods_buffer.getvalue(),
                "hasil_filter.ods", "application/vnd.oasis.opendocument.spreadsheet"
                )
                    
