import re
from collections import Counter

import gspread
from gspread.cell import Cell
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import time

scope = ['https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive"]

credentials = ServiceAccountCredentials.from_json_keyfile_name("stim-downtime-credentials.json", scope)
client = gspread.authorize(credentials)

spreadsheet = client.open_by_key("16H-eMNEUPi6PIKgPlKUHUQltg9T_3CqZdb3IhZ77NAg")
sheet_id = spreadsheet.worksheet("Простои")._properties['sheetId']
sheet = spreadsheet.sheet1


df = pd.read_csv('prostoi.csv')

df.iloc[:, 1].fillna(method='ffill', inplace=True)
df.iloc[:, 9].fillna(method='ffill', inplace=True)
df = df.sort_values(["Бригада", "Активности/Начало"])
df = df[~df['Статус'].isin(['Открыта', 'Завершена'])]
df = df[~df['Активности/Начало'].isna()]

"""
light_green = "red": 74, "green": 41, "blue": 88
blue = "red": 147, "green": 98, "blue": 3
light_blue = "red": 55, "green": 38, "blue": 8
purple = "red": 43, "green": 90, "blue": 67
light_red = "red": 12, "green": 52, "blue": 52 
dark_red = "red": 32, "green": 154, "blue": 154
light_yellow = "red": 1, "green": 27, "blue": 103
brown = "red": 126, "green": 132, "blue": 132
orange = "red": 1, "green": 103, "blue": 256

"""

column_value = {
    1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J', 11: 'K', 12: 'L', 13: 'M', 14: 'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R', 19: 'S', 20: 'T', 21: 'U', 22: 'V', 23: 'W', 24: 'X', 25: 'Y', 26: 'Z', 27: 'AA', 28: 'AB', 29: 'AC', 30: 'AD', 31: 'AE', 32: 'AF', 33: 'AG', 34: 'AH', 35: 'AI'
}

# clear_format_request = [
#     {
#     'updateCells': {
#         'range': {
#                 'sheetId': sheet_id,
#                 'startRowIndex': 1,
#                 'endRowIndex': 25,
#                 'startColumnIndex': 1,
#                 'endColumnIndex': 32,
#             },
#         'fields': 'userEnteredFormat,textFormat,numberFormat,formulaValue,userEnteredValue'
#         }
#     }
# ]
range_to_clear = 'B2:AF22'

sheet.batch_clear([range_to_clear])
sheet.format(range_to_clear, {"backgroundColor": {"red": 1, "green": 1, "blue": 1}})
sheet.clear_notes([range_to_clear])
# requests = [
#     {
#         'deleteRange': {
#             'range': {
#                 'sheetId': sheet_id,
#                 'startRowIndex': 1,
#                 'endRowIndex': 25,
#                 'startColumnIndex': 1,
#                 'endColumnIndex': 32,
#             },
#             'shiftDimension': 'ROWS',
#         }
#     }
# ]
# spreadsheet.batch_update({'requests': clear_format_request})

cell_text = []
cell_format = []
cell_notes = dict()

values_col_list = sheet.col_values(1)
values_row_list = sheet.row_values(1)

# for i in df["Активности/Начало"].head(30):
#     print(re.split("-| ", i))
sheet_brigades = values_col_list[1:30]
sheet_days = values_row_list[1:32]
reason = values_row_list[32:]
sheet_info = [sheet_brigades, sheet_days]


def values_add(brigade_downtime_info):
    taskmaster_value = brigade_downtime_info["brigade"]
    if taskmaster_value in values_col_list:
        row = values_col_list.index(taskmaster_value) + 1
        for day_value in sheet_days:
            if day_value in brigade_downtime_info:
                col = values_row_list.index(day_value) + 1
                cell_address = f"{column_value[col]}{row}"
                if brigade_downtime_info.get(day_value) == "Погодные условия":
                    cell_text.append(Cell(row=row, col=col, value="Непогода"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                   "red": 147, "green": 98, "blue": 3
                                },
                            "horizontalAlignment":
                                "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Переезды":
                    cell_text.append(Cell(row=row, col=col, value="Переезд"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 55, "green": 38, "blue": 8
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Доезд до места работ":
                    cell_text.append(Cell(row=row, col=col, value="Доезд"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                   "red": 55, "green": 38, "blue": 8
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Поломка":
                    cell_text.append(Cell(row=row, col=col, value="Поломка"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 32, "green": 154, "blue": 154
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Настройка оборудования":
                    cell_text.append(Cell(row=row, col=col, value="Настройка"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                   "red": 12, "green": 52, "blue": 52
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Мелкосрочный ремонт":
                    cell_text.append(Cell(row=row, col=col, value="Мелк.ремонт"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 12, "green": 52, "blue": 52
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Отдых (указывать при нахождении дома)":
                    cell_text.append(Cell(row=row, col=col, value="Дома"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 126, "green": 132, "blue": 132
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Выходной":
                    cell_text.append(Cell(row=row, col=col, value="Выходной"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                   "red": 1, "green": 27, "blue": 103
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Отсутствие Работы":
                    cell_text.append(Cell(row=row, col=col, value="Нет работы"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 126, "green": 132, "blue": 132
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Разгрузка материала":
                    cell_text.append(Cell(row=row, col=col, value="Разгрузка"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                   "red": 43, "green": 90, "blue": 67
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Отсутствие материала":
                    cell_text.append(Cell(row=row, col=col, value="Нет материала"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                   "red": 43, "green": 90, "blue": 67
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Болезнь":
                    cell_text.append(Cell(row=row, col=col, value="Болезнь"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                   "red": 43, "green": 90, "blue": 67
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Остановка работ из-за ДТП":
                    cell_text.append(Cell(row=row, col=col, value="ДТП"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                   "red": 43, "green": 90, "blue": 67
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                elif brigade_downtime_info.get(day_value) == "Check":
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 1, "green": 103, "blue": 256
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
                    cell_notes[cell_address] = "Проверить"
                elif pd.isnull(brigade_downtime_info.get(day_value)):
                    cell_text.append(Cell(row=row, col=col, value="Работа"))
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                   "red": 74, "green": 41, "blue": 88
                                },
                                "horizontalAlignment":
                                    "CENTER",
                            },
                        })
            else:
                break


for brigade in sheet_brigades:
    brigade_info = dict()
    one_brigade_values = df[(df['Бригада'] == brigade)]
    one_brigade_values = one_brigade_values.copy()
    brigade_info["brigade"] = brigade
    # brigade_info["days"] = set([date.split(' ')[0].split('-')[2] for date in one_brigade_values['Активности/Начало'].values.tolist()])
    brigade_have_two_days = Counter(date.split(' ')[0] for date in one_brigade_values['Активности/Начало'].values.tolist())
    for sheet_day in brigade_have_two_days:
        if brigade_have_two_days[sheet_day] > 1:
            # more_than_two_days = one_brigade_values[one_brigade_values['Активности/Начало'].dt.floor("D") == sheet_day]
            brigade_info[sheet_day.split('-')[2]] = "Check"
        else:
            # one_brigade_values.loc[:, 'Активности/Начало'] = pd.to_datetime(
            #     one_brigade_values.loc[:, 'Активности/Начало'], errors='coerce')
            one_brigade_values['Активности/Начало'] = pd.to_datetime(one_brigade_values['Активности/Начало'],
                                                                     errors='coerce')
            day_info = one_brigade_values[one_brigade_values['Активности/Начало'].dt.floor("D") == sheet_day].head(1).values.tolist()[0]
            brigade_info[sheet_day.split('-')[2]] = day_info[7]
    values_add(brigade_info)


sheet.update_cells(cell_text)
sheet.batch_format(cell_format)
try:
    sheet.update_notes(cell_notes)
except gspread.exceptions.APIError:
    pass
