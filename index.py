from datetime import datetime
import math
import os
from tempfile import NamedTemporaryFile
from zipfile import ZipFile
from PyPDF2 import PdfReader
import tabula

pdf_name = 'Anexo_I_Rol_2021RN_465.2021_RN473_RN478_RN480_RN513_RN536'
pdf_path = f'{pdf_name}.pdf'

print('Scaneando PDF...')
pdf_count = PdfReader(pdf_path).numPages
tables = tabula.read_pdf(pdf_path, pages=f'3-{pdf_count}', encoding='cp1252', lattice=True)

def formatText(text):
    if type(text) == str:
        value = text.replace('\r', ' ')
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

date = datetime.now().strftime('%d_%m_%Y_%H_%M_%S')

print('Iniciando a criação do arquivo zip e csv...')
zip_name = f'Anexo_I_Rol_Procedimentos_Eventos_Saúde_{date}.zip'
zip_path = f'{os.getcwd()}\{zip_name}'

csv_name = f'{pdf_name}.csv'

new_list = []
with ZipFile(zip_path, 'w') as zip:
    columns = []
    for col in tables[0].columns:
        columns.append(formatText(col))
    new_list.append(';'.join(columns))
    for table in tables:
        for row in table.values:
            new_row = []
            for col in row:
                new_row.append(formatText(col))
            new_list.append(';'.join(new_row))
    csv_text = '\n'.join(new_list)
    temp_name = writeTempFile(str.encode(csv_text, encoding='utf-8'))
    zip.write(temp_name, csv_name)
    os.unlink(temp_name)

print(f'Arquivo salvo em: "{zip_path}"')
