import os
import pandas as pd

# 1. Definisikan Path dan Nama File
BASE_PATH = "data"
target_file_path = os.path.join(
    BASE_PATH, "SCRAPING_REKAP_DAMPAK_BENCANA_ENDE_LATEST.xlsx"
)
master_file_path = os.path.join(BASE_PATH, "master_data.xlsx")

# 2. Baca kedua file Excel
print("Membaca file data...")
df_target = pd.read_excel(target_file_path)
df_master = pd.read_excel(master_file_path)

# Kolom-kolom yang ingin disalin/dilengkapi dari master data
columns_to_update = [
    "nmkab",
    "nmkec",
    "nmdesa",
    "nmsls",
    "nmsubsls",
    "pengawas",
    "pencacah",
    "nama_pcl",
    "nama_pml",
]

# 3. Antisipasi inkonsistensi tipe data pada regionCode
df_target["regionCode"] = df_target["regionCode"].astype(str).str.strip()
df_master["regionCode"] = df_master["regionCode"].astype(str).str.strip()

# 4. Filter kolom master data agar hanya mengambil regionCode dan kolom yang dibutuhkan
available_master_cols = [
    col for col in columns_to_update if col in df_master.columns
]
df_master_subset = df_master[["regionCode"] + available_master_cols]

# 5. Gabungkan data berdasarkan regionCode (Left Join)
df_merged = pd.merge(
    df_target,
    df_master_subset,
    on="regionCode",
    how="left",
    suffixes=("", "_master"),
)

# -------------------------------------------------------------
# 6. Pembaruan dan Penyelarasan 'pengawas' & 'nama_pml' dari Master Data
# -------------------------------------------------------------
print(
    "Memeriksa kesesuaian dan memperbarui kolom 'pengawas' serta 'nama_pml'..."
)

# Daftar kolom yang HARUS ditimpa/diperbarui jika berbeda dengan master data
strict_overwrite_cols = ["pengawas", "nama_pml"]

for col in available_master_cols:
  master_col_name = f"{col}_master"

  if col not in df_merged.columns:
    # Jika kolom belum ada di target, langsung pakai dari master
    df_merged[col] = df_merged[master_col_name]
  else:
    if col in strict_overwrite_cols:
      # PERBAIKAN: Gunakan nilai master_data sebagai acuan utama.
      # Jika di master_data ada nilainya, perbarui/timpa isi kolom target.
      df_merged[col] = df_merged[master_col_name].fillna(df_merged[col])
    else:
      # Untuk kolom wilayah/lainnya, cukup lengkapi jika ada yang kosong (fillna)
      df_merged[col] = df_merged[col].fillna(df_merged[master_col_name])

  # Hapus kolom bantuan '_master' setelah selesai diproses
  if master_col_name in df_merged.columns:
    df_merged.drop(columns=[master_col_name], inplace=True)

# -------------------------------------------------------------
# 7. Penyelarasan Kolom Username, Pencacah, dan Nama PCL
# -------------------------------------------------------------
print("Menyelaraskan kolom username, pencacah, dan nama_pcl...")

if "username" in df_merged.columns and "pencacah" in df_merged.columns:
  # Deteksi perbedaan antara username dan pencacah
  mask_mismatch = (
      df_merged["username"].astype(str).str.strip()
      != df_merged["pencacah"].astype(str).str.strip()
  )
  df_merged.loc[mask_mismatch, "pencacah"] = df_merged.loc[
      mask_mismatch, "username"
  ]

if "pencacah" in df_master.columns and "nama_pcl" in df_master.columns:
  # Peta hubungan pencacah -> nama_pcl dari master data
  df_master_clean = df_master.dropna(
      subset=["pencacah", "nama_pcl"]
  ).drop_duplicates(subset=["pencacah"])
  pcl_map = dict(
      zip(
          df_master_clean["pencacah"].astype(str).str.strip(),
          df_master_clean["nama_pcl"],
      )
  )

  # Sesuaikan nama_pcl berdasarkan kode pencacah
  df_merged["nama_pcl"] = (
      df_merged["pencacah"]
      .astype(str)
      .str.strip()
      .map(pcl_map)
      .fillna(df_merged["nama_pcl"])
  )

# -------------------------------------------------------------
# 8. Konversi Kolom Hasil Rekapitulasi Menjadi Format Number
# -------------------------------------------------------------
print("Mengonversi kolom-kolom numerik ke format Number...")

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

for col in numeric_cols:
  if col in df_merged.columns:
    df_merged[col] = pd.to_numeric(df_merged[col], errors="coerce")

# -------------------------------------------------------------
# 9. Simpan kembali hasil update ke file target asli
# -------------------------------------------------------------
print("Menyimpan data kembali ke file target...")
df_merged.to_excel(target_file_path, index=False)

print(
    "✅ Selesai! Kolom 'pengawas' dan 'nama_pml' pada file scraping berhasil"
    " disesuaikan dengan master_data."
)