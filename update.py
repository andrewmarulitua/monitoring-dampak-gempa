import os
import pandas as pd

# 1. Definisikan Path dan Nama File
BASE_PATH = "data"
target_file_path = os.path.join(
    BASE_PATH, "SCRAPING_REKAP_DAMPAK_BENCANA_ENDE_LATEST.xlsx"
)
master_file_path = os.path.join(BASE_PATH, "master_data.xlsx")

# 2. Buka ExcelFile untuk memeriksa sheet yang tersedia
print("Membaca file Excel...")
excel_target = pd.ExcelFile(target_file_path)
excel_master = pd.ExcelFile(master_file_path)

# Daftar sheet yang akan diproses beserta kolom spesifik yang ingin di-update/dilengkapi
sheets_config = {
    "Keluarga": [
        "nmkab",
        "nmkec",
        "nmdesa",
        "nmsls",
        "nmsubsls",
        "pengawas",
        "pencacah",
        "nama_pcl",
        "nama_pml",
    ],
    "Infrastruktur": [
        "nmkab",
        "nmkec",
        "nmdesa",
        "pengawas",
        "pencacah",
        "nama_pcl",
        "nama_pml",
    ],
}

# Kolom angka yang perlu dikonversi ke tipe data numeric
numeric_cols = [
    "total_data",
    "APPROVED BY Pengawas",
    "SUBMITTED BY Pencacah",
    "OPEN",
    "REJECTED BY Pengawas",
    "DRAFT",
    "EDITED BY Pengawas",
    "REVOKED BY Pengawas",
    "SUBMITTED RESPONDENT",
]

# Dictionary untuk menampung DataFrame hasil pembaruan
updated_dfs = {}

# Dictionary untuk menyimpan sheet lainnya jika ada di target agar tidak hilang saat overwriting
for sheet in excel_target.sheet_names:
    if sheet not in sheets_config:
        updated_dfs[sheet] = pd.read_excel(target_file_path, sheet_name=sheet)

# 3. Iterasi dan Proses Masing-Masing Sheet (Keluarga & Infrastruktur)
for sheet_name, columns_to_update in sheets_config.items():
    if (
        sheet_name not in excel_target.sheet_names
        or sheet_name not in excel_master.sheet_names
    ):
        print(
            f"⚠️ Sheet '{sheet_name}' tidak ditemukan di salah satu file. Melewati..."
        )
        continue

    print(f"\n--- Memproses Sheet: {sheet_name} ---")

    # Baca sheet masing-masing
    df_target = pd.read_excel(target_file_path, sheet_name=sheet_name)
    df_master = pd.read_excel(master_file_path, sheet_name=sheet_name)

    # Antisipasi inkonsistensi tipe data pada regionCode
    df_target["regionCode"] = df_target["regionCode"].astype(str).str.strip()
    df_master["regionCode"] = df_master["regionCode"].astype(str).str.strip()

    # Filter kolom master data agar hanya mengambil regionCode dan kolom yang dibutuhkan
    available_master_cols = [
        col for col in columns_to_update if col in df_master.columns
    ]
    df_master_subset = df_master[["regionCode"] + available_master_cols]

    # Gabungkan data berdasarkan regionCode (Left Join)
    df_merged = pd.merge(
        df_target,
        df_master_subset,
        on="regionCode",
        how="left",
        suffixes=("", "_master"),
    )

    # Pembaruan dan Penyelarasan 'pengawas' & 'nama_pml' dari Master Data
    strict_overwrite_cols = ["pengawas", "nama_pml"]

    for col in available_master_cols:
        master_col_name = f"{col}_master"

        if col not in df_merged.columns:
            # Jika kolom belum ada di target, langsung pakai dari master
            df_merged[col] = df_merged[master_col_name]
        else:
            if col in strict_overwrite_cols:
                # Gunakan nilai master_data sebagai acuan utama jika ada
                df_merged[col] = df_merged[master_col_name].fillna(
                    df_merged[col]
                )
            else:
                # Untuk kolom wilayah/lainnya, lengkapi jika kosong (fillna)
                df_merged[col] = df_merged[col].fillna(
                    df_merged[master_col_name]
                )

        # Hapus kolom bantuan '_master'
        if master_col_name in df_merged.columns:
            df_merged.drop(columns=[master_col_name], inplace=True)

    # Penyelarasan Username, Pencacah, dan Nama PCL
    if "username" in df_merged.columns and "pencacah" in df_merged.columns:
        mask_mismatch = (
            df_merged["username"].astype(str).str.strip()
            != df_merged["pencacah"].astype(str).str.strip()
        )
        df_merged.loc[mask_mismatch, "pencacah"] = df_merged.loc[
            mask_mismatch, "username"
        ]

    if "pencacah" in df_master.columns and "nama_pcl" in df_master.columns:
        df_master_clean = df_master.dropna(
            subset=["pencacah", "nama_pcl"]
        ).drop_duplicates(subset=["pencacah"])
        pcl_map = dict(
            zip(
                df_master_clean["pencacah"].astype(str).str.strip(),
                df_master_clean["nama_pcl"],
            )
        )

        df_merged["nama_pcl"] = (
            df_merged["pencacah"]
            .astype(str)
            .str.strip()
            .map(pcl_map)
            .fillna(df_merged["nama_pcl"])
        )

    # Konversi Kolom Numerik
    for col in numeric_cols:
        if col in df_merged.columns:
            df_merged[col] = pd.to_numeric(df_merged[col], errors="coerce")

    # Simpan ke dictionary penampung
    updated_dfs[sheet_name] = df_merged

# 4. Simpan kembali seluruh sheet ke file Excel target
print("\nMenyimpan kembali semua sheet ke file Excel target...")
with pd.ExcelWriter(target_file_path, engine="openpyxl") as writer:
    for sheet_name, df in updated_dfs.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("✅ Selesai! Sheet 'Keluarga' dan 'Infrastruktur' berhasil diperbarui.")