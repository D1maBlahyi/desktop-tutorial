from pathlib import Path  # Standard Python Module
import time  # Standard Python Module
import xlwings as xw  # pip install xlwings

# Adjust Paths
BASE_DIR = Path(__file__).parent
SOURCE_DIR = BASE_DIR / 'files'
OUTPUT_DIR = BASE_DIR / 'out'

# Create output directory
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

excel_files = Path(SOURCE_DIR).glob('*.xlsx')

# Create timestamp
t = time.localtime()
timestamp = time.strftime('%Y-%m-%d_%H%M', t)

with xw.App(visible=False) as app:
    combined_wb = app.books.add()
    for excel_file in excel_files:
        wb = app.books.open(excel_file)
        for sheet in wb.sheets:
            sheet.copy(after=combined_wb.sheets[0])
        wb.close()
    combined_wb.sheets[0].delete()
    combined_wb.save(OUTPUT_DIR / 'Бізнес_2023.xlsx')
    combined_wb.close()