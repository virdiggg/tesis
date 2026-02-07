import pandas as pd
import matplotlib.pyplot as plt
from util import formatting_excel
import os, re, uuid, datetime

target_dir = 'result'
os.makedirs('target', exist_ok=True)
os.makedirs(target_dir, exist_ok=True)
timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
random_string = str(uuid.uuid4()).replace('-', '')

# Bagian ini untuk konfigurasi, disesuaikan dengan kebutuhan
# =============================================================================
# Mapping profil responden
profile_cols = ['usia', 'jenis kelamin', 'pendidikan terakhir', 'pengalaman kerja', 'nama_perusahaan']

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

# Prefix file output/input
_postfix = '_Tangsel'

# File input
input_file = os.path.join('target', 'tangsel', f'Hasil_Profil_Responden{_postfix}.xlsx')

# Kalo ada indikator yang perlu di-skip (Biar hasil bootstrapping bisa ijo)
skip_indicators = ['P1', 'WB2', 'WB4', 'D2', 'D4', 'D6', 'D10']
# =============================================================================

# =============================================================================
# Mulai dari sini untuk proses, jangan diubah
def process_mapping_and_rename(full_mapping, short_mapping, df, skip_list):
    variabel_config = {}
    rename_map = {}

    cols_to_drop = [c for c in skip_list if c in df.columns]
    df = df.drop(columns=cols_to_drop)

    current_columns = list(df.columns)

    for key, full_name in full_mapping.items():
        short = short_mapping[key]
        code = key[key.find("(")+1:key.find(")")]
        var_name = full_name.replace(" ", "_")

        pattern = re.compile(rf"^{short}\d+$")

        original_cols = sorted(
            [col for col in current_columns if pattern.match(col)],
            key=lambda x: int(re.findall(r"\d+", x)[0])
        )

        new_cols = []
        for i, old_col in enumerate(original_cols, 1):
            new_col_name = f"{short}{i}"
            rename_map[old_col] = new_col_name
            new_cols.append(new_col_name)

        if new_cols:
            variabel_config[var_name] = {
                "cols": new_cols,
                "code": code
            }

    df.rename(columns=rename_map, inplace=True)

    return variabel_config, df

def get_kategori(persentase):
    if 20.00 <= persentase <= 36.00: return 'Sangat Tidak Setuju'
    elif 36.01 <= persentase <= 52.00: return 'Tidak Setuju'
    elif 52.01 <= persentase <= 68.00: return 'Netral'
    elif 68.01 <= persentase <= 84.00: return 'Setuju'
    elif 84.01 <= persentase <= 100.00: return 'Sangat Setuju'
    return '-'

if not os.path.exists(input_file):
    print(f"Error: File {input_file} tidak ditemukan!")
else:
    df_final = pd.read_excel(input_file)

    variabel_config, df_final = process_mapping_and_rename(
        full_mapping,
        short_mapping,
        df_final,
        skip_indicators
    )

    pernyataan_cols = []
    for v in variabel_config.values():
        pernyataan_cols.extend(v['cols'])

    smartpls = os.path.join(target_dir, f'to_smartpls_{timestamp}{_postfix}_{random_string}.csv')
    df_pernyataan_only = df_final[pernyataan_cols]
    df_pernyataan_only.to_csv(smartpls, index=False)
    print("File disimpan ke:", smartpls)

    for col in profile_cols:
        if col == 'nama_perusahaan':
            continue

        plt.figure(figsize=(8, 6))
        data_counts = df_final[col].value_counts()

        plt.pie(data_counts, labels=data_counts.index, autopct='%1.1f%%',
                startangle=140, colors=plt.cm.Paired.colors)
        plt.title(f'Distribusi Responden Berdasarkan {col.title()}')
        plt.axis('equal')

        output_file = os.path.join(target_dir, f'pie_chart_{col.replace(" ", "_")}{_postfix}.png')
        plt.savefig(output_file)
        plt.close()
        print("File disimpan ke:", output_file)

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
        df_summary['Urutan'] = df_summary['Persentase (%)'].rank(method='min', ascending=False).astype(int)

        output_path = os.path.join(target_dir, f'Analisis_{var_name}{_postfix}.xlsx')
        df_summary.to_excel(output_path, index=False)

        formatting_excel(output_path)