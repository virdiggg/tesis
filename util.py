from openpyxl import load_workbook
from openpyxl.styles import Font
import os

def formatting_excel(file_path, font_size=12):
    """
    Mengubah seluruh font di file Excel menjadi Times New Roman.
    """
    if not os.path.exists(file_path):
        print(f"Error: File {file_path} tidak ditemukan.")
        return

    try:
        wb = load_workbook(file_path)
        ws = wb.active

        default_font = Font(name="Times New Roman", size=font_size)

        for row in ws.iter_rows():
            for cell in row:
                cell.font = default_font

        wb.save(file_path)
        print(f"File disimpan di: {file_path}")

    except Exception as e:
        print(f"Terjadi kesalahan saat memproses Excel: {e}")