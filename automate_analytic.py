import pandas as pd
import matplotlib.pyplot as plt
from util import formatting_excel
import os, re

target_dir = 'result'
os.makedirs('target', exist_ok=True)
os.makedirs(target_dir, exist_ok=True)

# Bagian ini untuk konfigurasi, disesuaikan dengan kebutuhan
# =============================================================================
# File input
input_file = os.path.join('target', 'Hasil_Profil_Responden.xlsx')

# Mapping profil responden
profile_cols = ['usia', 'jenis kelamin', 'pendidikan terakhir', 'pengalaman kerja']

# Mapping variabel
full_mapping = {
    "P (X1)": "Pelatihan",
    "WB (X2)": "Work-Life Balance",
    "BK (X3)": "Beban Kerja",
    "D (Z)": "Digitalisasi",
    "PK (Y)": "Produktivitas Karyawan"
}

# Mapping variabel
short_mapping = {
    "P (X1)": "P",
    "WB (X2)": "WB",
    "BK (X3)": "BK",
    "D (Z)": "D",
    "PK (Y)": "PK"
}
# =============================================================================

# =============================================================================
# Mulai dari sini untuk proses, jangan diubah
def process_mapping_auto(full_mapping, short_mapping, df):
    """
    Membuat variabel_config otomatis berdasarkan kolom DataFrame
    """

    variabel_config = {}

    columns = list(df.columns)

    for key, full_name in full_mapping.items():
        short = short_mapping[key]

        code = key[key.find("(")+1:key.find(")")]

        var_name = full_name.replace(" ", "_")

        pattern = re.compile(rf"^{short}\d+$")
        indikator_cols = sorted(
            [col for col in columns if pattern.match(col)],
            key=lambda x: int(re.findall(r"\d+", x)[0])
        )

        if indikator_cols:
            variabel_config[var_name] = {
                "cols": indikator_cols,
                "code": code
            }

    return variabel_config

if not os.path.exists(input_file):
    print(f"Error: File {input_file} tidak ditemukan!")
else:
    df_final = pd.read_excel(input_file)

    variabel_config = process_mapping_auto(
        full_mapping,
        short_mapping,
        df_final
    )

    pernyataan_cols = []
    for v in variabel_config.values():
        pernyataan_cols.extend(v['cols'])

    df_pernyataan_only = df_final[pernyataan_cols]
    df_pernyataan_only.to_csv(os.path.join(target_dir, 'to_smartpls.csv'), index=False)

    for col in profile_cols:
        plt.figure(figsize=(8, 6))
        data_counts = df_final[col].value_counts()

        plt.pie(data_counts, labels=data_counts.index, autopct='%1.1f%%',
                startangle=140, colors=plt.cm.Paired.colors)
        plt.title(f'Distribusi Responden Berdasarkan {col.title()}')
        plt.axis('equal')

        output_file = os.path.join(target_dir, f'pie_chart_{col.replace(" ", "_")}.png')
        plt.savefig(output_file)
        plt.close()
        print("File disimpan ke:", output_file)

    def get_kategori(persentase):
        if 20.00 <= persentase <= 36.00: return 'Sangat Tidak Setuju'
        elif 36.01 <= persentase <= 52.00: return 'Tidak Setuju'
        elif 52.01 <= persentase <= 68.00: return 'Netral'
        elif 68.01 <= persentase <= 84.00: return 'Setuju'
        elif 84.01 <= persentase <= 100.00: return 'Sangat Setuju'
        return '-'

    total_responden = len(df_final)
    skor_ideal = total_responden * 5

    for var_name, config in variabel_config.items():
        summary_data = []
        columns = config['cols']
        var_code = config['code']

        for idx, col in enumerate(columns, 1):
            counts = df_final[col].value_counts().reindex([1, 2, 3, 4, 5], fill_value=0)

            skor_aktual = sum(counts[i] * i for i in range(1, 6))
            persentase = (skor_aktual / skor_ideal) * 100

            row = {
                'No': idx,
                'Indikator': col,
                'Pertanyaan': f"{var_code}.{idx}",
                'Skor 1': counts[1],
                'Skor 2': counts[2],
                'Skor 3': counts[3],
                'Skor 4': counts[4],
                'Skor 5': counts[5],
                'Skor Aktual': skor_aktual,
                'Skor Ideal': skor_ideal,
                'Persentase (%)': round(persentase, 2),
                'Kategori': get_kategori(persentase)
            }
            summary_data.append(row)

        df_summary = pd.DataFrame(summary_data)
        output_path = os.path.join(target_dir, f'Analisis_{var_name}.xlsx')
        df_summary.to_excel(output_path, index=False)

        formatting_excel(output_path)