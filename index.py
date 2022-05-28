from datetime import datetime
import math
import os
import re
from tempfile import NamedTemporaryFile
from zipfile import ZipFile
import pdfplumber
from tqdm import tqdm

pdf_name = 'Anexo_I_Rol_2021RN_465.2021_RN473_RN478_RN480_RN513_RN536'
pdf_path = f'{pdf_name}.pdf'

print('Scaneando PDF...')
pdf = pdfplumber.open(pdf_path)

def formatText(text):
    if type(text) == str:
        value = text.replace('\n', '')
        return f'"{value}"' if ";" in value else value
    elif not math.isnan(text):
        return f'"{text}"'
    else:
        return ''

def writeTempFile(bytes):
    file = NamedTemporaryFile(delete=False)
    file.write(bytes)
    file.close()
    return file.name

print('Carregando Legenda...')
first_page = pdf.pages[3]
text = first_page.extract_text()

legenda = {}
pares = re.findall(r'([A-Z]+):\s+(.+?)\s+(?:(?=[A-Z]+:)|\d+)', text)
for par in pares:
    legenda[par[0]] = par[1]

date = datetime.now().strftime('%d_%m_%Y_%H_%M_%S')

print('Iniciando a criação do arquivo zip e csv...')

zip_name = f'Anexo_I_Rol_Procedimentos_Eventos_Saúde_{date}.zip'
zip_path = f'{os.getcwd()}\{zip_name}'
csv_name = f'{pdf_name}.csv'

new_list = []

columns = []
for col in first_page.extract_table()[0]:
    columns.append(formatText(col))
new_list.append(';'.join(columns))

replace_list = ['OD', 'AMB']

with ZipFile(zip_path, 'w') as zip:
    for page in tqdm(pdf.pages[3:], ascii=True):
        table = page.extract_table()
        for row in table[1:]:
            new_row = []
            for i in range(len(row)):
                text = row[i]
                if text in replace_list:
                    text = legenda[text]
                new_row.append(formatText(text))
            new_list.append(';'.join(new_row))
    csv_text = '\n'.join(new_list)
    temp_name = writeTempFile(str.encode(csv_text, encoding='utf-8'))
    zip.write(temp_name, csv_name)
    os.unlink(temp_name)
print(f'Arquivo salvo em: "{zip_path}"')
