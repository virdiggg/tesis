import pandas as pd
import os

# 1. Definisi Path & Data
path_asli = os.path.join('target', 'tangsel', 'pernyataan_tangsel.csv')
path_edit = os.path.join('result', 'tangsel', 'to_smartpls.csv')
output_excel = os.path.join('result', 'tangsel', 'perbandingan_full_41.xlsx')

df_asli = pd.read_csv(path_asli)
df_edit = pd.read_csv(path_edit)

indikator_asli = [
    'BK1','BK2','BK3','BK4','BK5','BK6','WB1','WB2','WB3','WB4','WB5','WB6',
    'P1','P2','P3','P4','P5','P6','P7','P8','P9','P10','D1','D2','D3','D4',
    'D5','D6','D7','D8','D9','D10','D11','PK1','PK2','PK3','PK4','PK5','PK6','PK7','PK8'
]
skip_indicators = ['P1', 'WB2', 'WB4', 'D2', 'D4', 'D6', 'D10']

# 2. Identifikasi urutan kolom berdasarkan Data Edit
# Kita mapping dulu kolom df_edit ke nama indikator aslinya
indikator_setelah_filter = [i for i in indikator_asli if i not in skip_indicators]
mapping_dict = dict(zip(df_edit.columns, indikator_setelah_filter))

# Urutan 1: Kolom yang ada di Data Edit (sudah di-rename ke nama asli)
urutan_utama = list(mapping_dict.values())

# Urutan 2: Kolom yang dihapus (skip_indicators) diletakkan di akhir
# Kita pastikan urutannya tetap konsisten sesuai daftar asli
urutan_tambahan = [i for i in indikator_asli if i in skip_indicators]

# Gabungkan keduanya: Total harus 41 kolom
urutan_final_41 = urutan_utama + urutan_tambahan

# 3. Susun Ulang df_asli berdasarkan urutan_final_41
df_asli_reconstructed = df_asli[urutan_final_41]

# 4. Simpan ke Excel dan CSV
os.makedirs(os.path.dirname(output_excel), exist_ok=True)

with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
    # Sheet ini berisi 41 kolom dengan urutan: [Kolom Aktif (Urutan Edit)] + [Kolom Skip]
    df_asli_reconstructed.to_excel(writer, sheet_name='Full_41_Indicators', index=False)
    
    # Simpan juga mapping-nya di sheet terpisah untuk dokumentasi
    df_mapping = pd.DataFrame(list(mapping_dict.items()), columns=['Nama_di_File_Edit', 'Nama_Asli_Sesuai_Urutan'])
    df_mapping.to_excel(writer, sheet_name='Mapping_Log', index=False)

df_asli_reconstructed.to_csv(os.path.join('result', 'tangsel', 'full_41_reordered.csv'), index=False)

print(f"Selesai! File dengan 41 kolom (urutan mengikuti file edit) telah disimpan.")
print(f"Total kolom: {len(df_asli_reconstructed.columns)}")