import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import numpy as np

import openpyxl
from openpyxl.reader.excel import load_workbook

scope = ['https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive"]

credentials = ServiceAccountCredentials.from_json_keyfile_name("stim-downtime-credentials.json", scope)
client = gspread.authorize(credentials)

spreadsheet = client.open_by_key("1FbJYDcC-Ab2lEtS2hljMvCIGXDSXyNYSa_9p7VbavHM")

file_name = 'Контракты 2024.xlsx'


def google_table_read():
    sheet = spreadsheet.worksheet('Контракты 2024')
    columns_to_keep = ['Контракты 2024\n (наименование  по Оду)', 'ТБЕ', 'Номер и дата договора']
    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    df_next = df[columns_to_keep]
    df_next.replace('', np.nan, inplace=True)
    # df['Контракты 2024\n (наименование  по Оду)'].fillna(method='ffill', inplace=True)
    df_cleaned = df_next.dropna(subset=['ТБЕ'])
    df_cleaned.columns = ['Контракт', 'Регион', '1С']
    df_cleaned = df_cleaned.sort_values(by='Регион', ascending=True)
    df_cleaned.to_excel(file_name, index=False, header=True)


def edit_table():
    workbook = load_workbook(file_name)
    worksheet = workbook.active
    worksheet.column_dimensions["A"].width = 30
    worksheet.column_dimensions["B"].width = 10
    worksheet.column_dimensions["C"].width = 30
    workbook.save(file_name)


google_table_read()
edit_table()