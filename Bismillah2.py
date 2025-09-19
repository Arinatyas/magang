import streamlit as st
import pandas as pd
import io
import altair as alt

st.set_page_config(page_title="📊 Gabung & Visualisasi Data", layout="wide")

st.title("📊 Aplikasi Gabung, Filter, dan Visualisasi Data Multi-Sheet")

# === UPLOAD FILE ===
uploaded_files = st.file_uploader(
    "Unggah beberapa file (Excel/CSV/ODS) sekaligus",
    type=["xlsx", "csv", "ods"],
    accept_multiple_files=True
)

header_row = st.number_input(
    "Pilih baris header (mulai dari 0, misalnya baris ke-6 berarti 5)",
    min_value=0, value=0
)

all_data = []

def clean_headers(df):
    """Isi header kosong dengan nilai dari sebelumnya"""
    new_cols = []
    prev = "Kolom"
    for i, col in enumerate(df.columns):
        if "Unnamed" in str(col) or str(col).strip() == "":
            new_cols.append(f"{prev}_{i}")
        else:
            new_cols.append(str(col))
            prev = str(col)
    df.columns = new_cols
    return df

def clean_rows(df):
    """Isi baris kosong dengan 0/kosong sesuai tipe"""
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].fillna(0)
        else:
            df[col] = df[col].fillna("kosong")
    return df

if uploaded_files:
    for file in uploaded_files:
        if file.name.endswith(".csv"):
            df = pd.read_csv(file, header=header_row)
            df = clean_headers(df)
            df = clean_rows(df)
            all_data.append(df)
        else:  # Excel atau ODS
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet, header=header_row)
                df = clean_headers(df)
                df = clean_rows(df)
                df["Sumber_Sheet"] = f"{file.name} - {sheet}"
                all_data.append(df)

    data_gabungan = pd.concat(all_data, ignore_index=True)
else:
    data_gabungan = pd.DataFrame()

# === DATA GABUNGAN ===
if not data_gabungan.empty:
    st.subheader("🔎 Data Gabungan (setelah diproses)")
    st.dataframe(data_gabungan.head())

    # === DATA PENYARINGAN ===
    st.subheader("📌 Hasil Penyaringan")

    kolom_pilihan = st.multiselect(
        "Pilih kolom yang ingin ditampilkan & difilter:",
        options=data_gabungan.columns.tolist(),
        default=data_gabungan.columns[:2].tolist()
    )

    if kolom_pilihan:
        data_penyaringan = data_gabungan[kolom_pilihan].copy()

        for kol in kolom_pilihan:
            nilai_unik = data_gabungan[kol].dropna().unique().tolist()
            nilai_pilihan = st.multiselect(f"Pilih nilai untuk kolom {kol}:", nilai_unik)
            if nilai_pilihan:
                data_penyaringan = data_penyaringan[data_penyaringan[kol].isin(nilai_pilihan)]

        st.dataframe(data_penyaringan)

        # === DOWNLOAD FILTERED DATA ===
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine="xlsxwriter") as writer:
            data_penyaringan.to_excel(writer, index=False, sheet_name="FilteredData")
        st.download_button(
            "⬇️ Unduh Hasil Penyaringan (Excel)",
            data=buffer_excel,
            file_name="hasil_penyaringan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        buffer_ods = io.BytesIO()
        with pd.ExcelWriter(buffer_ods, engine="odf") as writer:
            data_penyaringan.to_excel(writer, index=False, sheet_name="FilteredData")
        st.download_button(
            "⬇️ Unduh Hasil Penyaringan (ODS)",
            data=buffer_ods,
            file_name="hasil_penyaringan.ods",
            mime="application/vnd.oasis.opendocument.spreadsheet"
        )

        # === VISUALISASI ===
        st.subheader("📈 Visualisasi Data")
        if len(kolom_pilihan) >= 2:
            col1, col2 = kolom_pilihan[:2]

            if pd.api.types.is_numeric_dtype(data_penyaringan[col1]) and pd.api.types.is_numeric_dtype(data_penyaringan[col2]):
                chart = alt.Chart(data_penyaringan).mark_circle(size=60).encode(
                    x=alt.X(col1, type="quantitative"),
                    y=alt.Y(col2, type="quantitative"),
                    tooltip=kolom_pilihan
                ).interactive()
                st.altair_chart(chart, use_container_width=True)

            elif pd.api.types.is_numeric_dtype(data_penyaringan[col1]) and not pd.api.types.is_numeric_dtype(data_penyaringan[col2]):
                chart = alt.Chart(data_penyaringan).mark_bar().encode(
                    x=alt.X(col2, type="nominal"),
                    y=alt.Y(col1, type="quantitative"),
                    tooltip=kolom_pilihan
                )
                st.altair_chart(chart, use_container_width=True)

            elif not pd.api.types.is_numeric_dtype(data_penyaringan[col1]) and pd.api.types.is_numeric_dtype(data_penyaringan[col2]):
                chart = alt.Chart(data_penyaringan).mark_bar().encode(
                    x=alt.X(col1, type="nominal"),
                    y=alt.Y(col2, type="quantitative"),
                    tooltip=kolom_pilihan
                )
                st.altair_chart(chart, use_container_width=True)

            else:
                chart = alt.Chart(data_penyaringan).mark_bar().encode(
                    x=alt.X(col1, type="nominal"),
                    y=alt.Y(col2, type="nominal"),
                    tooltip=kolom_pilihan
                )
                st.altair_chart(chart, use_container_width=True)

else:
    st.info("📂 Silakan unggah file terlebih dahulu.")
            
