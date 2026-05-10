import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(layout="wide")
st.title("Gabung & Format Anomali")

st.markdown("""
Web ini digunakan untuk menggabungkan hasil anomali dari beberapa file Excel,
kemudian melakukan formatting otomatis anomali podes terhadap label A1-A154.
            
Dibawah ini merupakan link sql lab untuk rekap 1-4:
""")

st.link_button(
    "Rekap Anomali 1-52",
    "https://fasih-dashboard.bps.go.id/superset/sqllab?savedQueryId=26553"
)
st.link_button(
    "Rekap Anomali 53-90",
    "https://fasih-dashboard.bps.go.id/superset/sqllab?savedQueryId=26555"
)
st.link_button(
    "Rekap Anomali 91-128",
    "https://fasih-dashboard.bps.go.id/superset/sqllab?savedQueryId=26554"
)
st.link_button(
    "Rekap Anomali 130-154",
    "https://fasih-dashboard.bps.go.id/superset/sqllab?savedQueryId=26552"
)

st.markdown("""
## Petunjuk Penggunaan

1. Upload file Excel hasil rekap.
2. Sistem akan menggabungkan seluruh data otomatis.
3. Label anomali akan diformat otomatis.
4. Download hasil akhir pada tombol yang tersedia.
""")

# =========================
# 1. UPLOAD FILE
# =========================
uploaded_files = st.file_uploader(
    "Upload file Excel (bisa banyak)",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    st.success(f"{len(uploaded_files)} file berhasil diupload")

    with st.spinner("Memproses data..."):

        # =========================
        # 2. LOAD & CONCAT
        # =========================
        needed_cols = [
            "NO.",
            "KODE",
            "KODE_KEC",
            "KEC",
            "KODE_DESA",
            "DESA",
            "ANOMALI",
            "LINK",
            "CATATAN"
        ]

        all_df = []

        for f in uploaded_files:

            temp = pd.read_excel(
                    f,
                    dtype={
                        "KODE": str,
                        "KODE_KEC": str,
                        "KODE_DESA": str
                    }
                )

            # validasi kolom wajib
            required_cols = ["KODE", "ANOMALI"]

            if not all(col in temp.columns for col in required_cols):
                st.error(
                    f"File {f.name} tidak memiliki kolom wajib: KODE dan ANOMALI"
                )
                st.write("Kolom ditemukan:", temp.columns.tolist())
                st.stop()

            # ambil hanya kolom yang tersedia
            available_cols = [
                col for col in needed_cols
                if col in temp.columns
            ]

            temp = temp[available_cols]

            all_df.append(temp)

        df = pd.concat(
            all_df,
            ignore_index=True
        )

        # =========================
        # 4. CLEANING
        # =========================
        df["KODE"] = df["KODE"].astype(str).str.strip()

        df["ANOMALI"] = (
            df["ANOMALI"]
            .fillna("")
            .astype(str)
            .str.replace(";", ",", regex=False)
            .str.replace(r"\s+", "", regex=True)
            .str.upper()
        )

        # =========================
        # 5. BASE DATA
        # =========================
        base = df.copy()

        # =========================
        # 6. SPLIT + EXPLODE
        # =========================
        exploded = (
            df.assign(
                ANOMALI=df["ANOMALI"].str.split(",")
            )
            .explode("ANOMALI")
        )

        exploded["ANOMALI"] = (
            exploded["ANOMALI"]
            .fillna("")
            .str.strip()
        )

        # =========================
        # 7. FILTER VALID
        # =========================
        exploded = exploded[
            exploded["ANOMALI"].str.match(r"^A\d+$", na=False)
        ]

        st.write(
            "Jumlah baris setelah filter:",
            len(exploded)
        )

        # =========================
        # 8. HANDLE EMPTY
        # =========================
        if exploded.empty:

            st.warning("Tidak ditemukan anomali valid")

            anomali_result = pd.DataFrame({
                "KODE": [],
                "ANOMALI": []
            })

        else:

            # =========================
            # 9. AMBIL ANGKA
            # =========================
            exploded["num"] = (
                exploded["ANOMALI"]
                .str.extract(r"(\d+)")
                .astype(int)
            )

            # =========================
            # 10. DROP DUPLIKAT
            # =========================
            exploded = exploded.drop_duplicates(
                subset=["KODE", "ANOMALI"]
            )

            # =========================
            # 11. FORMAT #Axx
            # =========================
            exploded["ANOMALI"] = (
                "#" + exploded["ANOMALI"]
            )

            # =========================
            # 12. SORT
            # =========================
            exploded = exploded.sort_values(
                ["KODE", "num"]
            )

            # =========================
            # 13. GROUP
            # =========================
            anomali_result = (
                exploded.groupby("KODE")["ANOMALI"]
                .agg(", ".join)
                .reset_index()
            )

        # =========================
        # 14. BASE UNIQUE
        # =========================
        base_unique = (
            base
            .drop(columns=["ANOMALI"], errors="ignore")
            .drop_duplicates(subset=["KODE"])
        )

        # =========================
        # 15. MERGE
        # =========================
        result = base_unique.merge(
            anomali_result,
            on="KODE",
            how="left"
        )

        # =========================
        # 16. FILLNA
        # =========================
        result["ANOMALI"] = (
            result["ANOMALI"]
            .fillna("")
        )

        # =========================
        # 17. JUMLAH ANOMALI
        # =========================
        result["TOTAL ANOMALI"] = (
            result["ANOMALI"]
            .str.count("#")
        )

    # =========================
    # 18. FILTER UI
    # =========================
    st.subheader("Filter")

    only_anomaly = st.checkbox(
        "Hanya tampilkan yang ada anomali"
    )

    if only_anomaly:
        result_show = result[
            result["ANOMALI"] != ""
        ]
    else:
        result_show = result

    # =========================
    # 19. TAMPILKAN
    # =========================
    st.subheader("Hasil Akhir")

    def highlight_row(row):
        color = (
            "background-color: #ffe6e6"
            if row["ANOMALI"]
            else ""
        )
        return [color] * len(row)

    st.dataframe(
        result_show.style.apply(
            highlight_row,
            axis=1
        ),
        use_container_width=True
    )

    # =========================
    # 20. DOWNLOAD
    # =========================
    def to_excel(df_export):

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            df_export.to_excel(
                writer,
                index=False,
                sheet_name="hasil"
            )

            ws = writer.sheets["hasil"]

            # Auto width
            for column_cells in ws.columns:

                length = max(
                    len(str(cell.value))
                    if cell.value is not None else 0
                    for cell in column_cells
                )

                ws.column_dimensions[
                    column_cells[0].column_letter
                ].width = min(length + 5, 50)

        return output.getvalue()

    st.download_button(
        label="Download Excel",
        data=to_excel(result_show),
        file_name="hasil_anomali.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
