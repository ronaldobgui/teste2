from datetime import datetime
import os
from pathlib import Path
import re
from zipfile import ZipFile
import pdfplumber
from tqdm import tqdm


def formatText(text):
    value = text.replace('\n', '')
    return f'"{value}"' if ";" in value else value


def convertPdfToCsv(pdf_file_path, replace_list):
    print('Scaneando PDF...')
    with pdfplumber.open(pdf_file_path) as pdf:
        print('Carregando legenda...')
        first_page = pdf.pages[3]
        first_page_text = first_page.extract_text()

        legenda = {}
        key_values_list = re.findall(
            r'([A-Z]+):\s+(.+?)\s+(?:(?=[A-Z]+:)|\d+)', first_page_text)
        for key_value in key_values_list:
            legenda[key_value[0]] = key_value[1]

        print('Gerando arquivo csv...')
        csv_rows = []
        csv_column_titles = []
        for cell in first_page.extract_table()[0]:
            formatted_text = formatText(cell)
            csv_column_titles.append(formatted_text)
        csv_rows.append(';'.join(csv_column_titles))

        def formatTextAndReplace(cell):
            if cell in replace_list:
                cell = legenda[cell]
            return formatText(cell)

        for page in tqdm(pdf.pages[3:], ascii=True):
            rows = page.extract_table()
            # ignora a que contém os títulos
            for row in rows[1:]:
                new_row = map(formatTextAndReplace, row)
                csv_rows.append(';'.join(new_row))

        return '\n'.join(csv_rows)


def saveZipFromTextData(file_name, text_data):
    date = datetime.now().strftime('%d_%m_%Y_%H_%M_%S')
    zip_name = f'Teste_2_Ronaldo_Borges_Guimarães_{date}.zip'
    zip_path = os.path.join(os.getcwd(), zip_name)

    try:
        with ZipFile(zip_path, 'w') as zip:
            zip.writestr(file_name, text_data)
    except:
        if(os.path.exists(zip_path)):
            os.unlink(zip_path)

        raise Exception('Erro ao salvar arquivo zip...')

    print(f'Arquivo salvo em: "{zip_path}"')


def generateCsvZipFromPdf(pdf_file_path, replace_list):
    try:
        if(os.path.exists(pdf_file_path) == False):
            raise Exception('Anexo não encontrado...')

        pdf_name = Path(pdf_file_path).stem
        csv_file_name = f'{pdf_name}.csv'
        csv_data_text = convertPdfToCsv(pdf_file_path, replace_list)
        saveZipFromTextData(csv_file_name, csv_data_text)

    except Exception as e:
        print(f'Erro: {e}')


generateCsvZipFromPdf(
    pdf_file_path='Anexo_I_Rol_2021RN_465.2021_RN473_RN478_RN480_RN513_RN536.pdf',
    replace_list=['OD', 'AMB']
)
