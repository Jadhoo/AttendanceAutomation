from PIL import Image
from pytesseract import pytesseract
from datetime import datetime
import openpyxl
from pathlib import Path
from glob import glob
from os.path import abspath
from os import remove
from re import compile, sub

path_to_tesseract = r'C:\Users\user\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

participants_list = []

for screenshot in glob("screenshots/*.*"):
    image_path = abspath(screenshot)
    img = Image.open(image_path)
    pytesseract.tesseract_cmd = path_to_tesseract
    participant_names = pytesseract.image_to_string(img)
    participant_names = list(participant_names[:-1].split("\n"))
    participant_names = list(filter(None, participant_names))
    participants_list += participant_names
    print('removing screenshot')
    remove(image_path)

if participants_list != []:
    folder_name = f"{datetime.now().strftime('%d-%m-%Y')}"
    Path(folder_name).mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    sheet = wb.active
    cleaned_name = compile('\\.{2,}')
    for i in range(len(participants_list)):
        sheet[f'A{i + 1}'] = cleaned_name.sub("", participants_list[i])
    sheet.column_dimensions['A'].width = 25
    title = "Attendance"
    sheet.title = title
    wb_name = f"Attendance {len(glob(folder_name + '/*.xlsx'))}.xlsx"
    wb.save(f'{folder_name}/{wb_name}')
