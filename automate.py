import pandas as pd
import matplotlib.pyplot as plt
import os

os.makedirs('target', exist_ok=True)
os.makedirs('result', exist_ok=True)
input_file = os.path.join('target', 'Hasil_Profil_Responden.xlsx')
target_dir = 'result'

if not os.path.exists(input_file):
    print(f"Error: File {input_file} tidak ditemukan!")
else:
    df_final = pd.read_excel(input_file)

    profile_cols = ['usia', 'jenis kelamin', 'pendidikan terakhir', 'pengalaman kerja']

    variabel_config = {
        'Pelatihan': {'cols': [f'P{i}' for i in range(1, 11)], 'code': 'X1'},
        'Work_Life_Balance': {'cols': [f'WB{i}' for i in range(1, 7)], 'code': 'X2'},
        'Beban_Kerja': {'cols': [f'BK{i}' for i in range(1, 7)], 'code': 'X3'},
        'Digitalisasi': {'cols': [f'D{i}' for i in range(1, 12)], 'code': 'Z'},
        'Produktivitas': {'cols': [f'PK{i}' for i in range(1, 9)], 'code': 'Y'}
    }

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

        plt.savefig(os.path.join(target_dir, f'pie_chart_{col.replace(" ", "_")}.png'))
        plt.close()

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