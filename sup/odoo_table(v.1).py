import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import math
import time

# Подсоединение к Google Таблицам
scope = ['https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive"]

credentials = ServiceAccountCredentials.from_json_keyfile_name("gs-credentials.json", scope)
client = gspread.authorize(credentials)

# sheet = client.open("StimDB").sheet1
spreadsheet = client.open_by_key("1nqwdAH7T1yKiP267bf0YdvV2pD8xmvcze4Wnj86w3Xk")
sheet_id = spreadsheet.worksheet("Октябрь")._properties['sheetId']
sheet = spreadsheet.sheet1

df = pd.read_csv('odoo.csv').values.tolist()
# print(df.values)


# cell = sheet.acell(cell_address)
# sheet.update_note(cell_address, comment_text)
# sheet.update_notes()
# sheet.batch_format(format)
"""
green = "red": 144, "green": 83, "blue": 185
yellow = "red": 1, "green": 1, "blue": 256
orange = "red": 1, "green": 103, "blue": 256
red = "red": 1, "green": 256, "blue": 256
dark_red = "red": 154, "green": 256, "blue": 256
light_yellow = "red": 1, "green": 49, "blue": 154


"""

column_value = {
    1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J', 11: 'K', 12: 'L', 13: 'M', 14: 'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R', 19: 'S', 20: 'T', 21: 'U', 22: 'V', 23: 'W', 24: 'X', 25: 'Y', 26: 'Z', 27: 'AA', 28: 'AB', 29: 'AC', 30: 'AD', 31: 'AE', 32: 'AF', 33: 'AG', 34: 'AH', 35: 'AI'
}
# column_mapping = {chr(65 + i // 26) + chr(65 + i % 26): i + 1 for i in range(35)}
requests = [
    {
        'deleteRange': {
            'range': {
                'sheetId': sheet_id,
                'startRowIndex': 1,
                'endRowIndex': 22,
                'startColumnIndex': 1,
                'endColumnIndex': 35,
            },
            'shiftDimension': 'ROWS',
        }
    }
]
spreadsheet.batch_update({'requests': requests})


def is_nan(x):
    return (x != x)


cell_format = []
cell_notes = dict()

values_col_list = sheet.col_values(1)
values_row_list = sheet.row_values(1)


def values_add(brigade_day_info, day_value):
    try:
        taskmaster_value = list(brigade_day_info.keys())[0]
        if taskmaster_value in values_col_list:
            row = values_col_list.index(taskmaster_value) + 1
            col = values_row_list.index(day_value) + 1
            cell_address = f"{column_value[col]}{row}"
            # cell_address = sheet.cell(row, col).address
            for key, value in brigade_day_info.items():
                # kontur_and_shmel_exists = any(
                #     map(lambda x: ("Контур" in x) or ("шмель" in x) or ("Шмель" in x) if not is_nan(x) else False,
                #         value[7:]))
                kontur_and_shmel_exists = any(
                    map(lambda x: ("контур" in x.lower()) or ("шмель" in x.lower()) if not is_nan(x) else False,
                        value[7:]))
                preliminary_marking = any(
                    map(lambda x: ("предварительная" in x.lower()) if not is_nan(x) else False,
                        value[7:]))
                if value[0] == 'Закрыта' and value[1] == 'Линейные' and value[2] > 0 and value[6] > 0 and (
                        value[4] == 0 and value[5] == 0)\
                        and kontur_and_shmel_exists:
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 1, "green": 49, "blue": 154
                                },
                            },
                        })
                    cell_notes[cell_address] = "Нет операционного времени"
                elif value[0] == "Завершена" or value[0] == "Открыта":
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 1, "green": 103, "blue": 256
                                },
                            },
                        })
                    cell_notes[cell_address] = "Смена не закрыта"
                elif value[0] == 'Закрыта' and value[2] == 0 and (value[1] != 'Линейные' and value[1] != 'Ручные'):
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 144, "green": 83, "blue": 185
                                },
                            },
                        })
                elif value[0] == 'Закрыта' and value[2] >= 0 and value[1] == 'Демаркировка':
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 144, "green": 83, "blue": 185
                                },
                            },
                        })
                elif value[0] == 'Закрыта' and value[1] == 'Ручные' and value[2] > 0 and value[6] > 0:
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 144, "green": 83, "blue": 185
                                },
                            },
                        })
                elif value[0] == 'Закрыта' and value[1] == 'Линейные' and value[2] > 0 and value[6] > 0 and (
                        value[4] > 0 or value[5] > 0) \
                        and kontur_and_shmel_exists:
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 144, "green": 83, "blue": 185
                                },
                            },
                        })
                elif value[0] == 'Закрыта' and (0 < value[2] < 200) and value[6] > 0 and preliminary_marking:
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 144, "green": 83, "blue": 185
                                },
                            },
                        })
                elif value[0] == 'Закрыта' and (0 < value[2] < 200) and preliminary_marking:
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 144, "green": 83, "blue": 185
                                },
                            },
                        })
                elif value[0] == 'Закрыта' and value[2] > 0 and (value[1] == 'Линейные' and value[1] == 'Ручные')\
                        and value[6] == 0:
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 154, "green": 256, "blue": 256
                                },
                            },
                        })
                    cell_notes[cell_address] = "Не указан расход материала"
                else:
                    cell_format.append(
                        {
                            "range": cell_address,
                            "format": {
                                "backgroundColor": {
                                    "red": 1, "green": 256, "blue": 256
                                },
                            },
                        })
                    cell_notes[cell_address] = "Проверить"
    except gspread.exceptions.APIError:
        print("Waiting a little")
        time.sleep(65)


for id, brigade in enumerate(df):
    nan = is_nan(brigade[0])
    if not nan:
        brigade_info = dict()
        i = 1
        date = brigade[5].split(' ')
        day = date[0].split('-')[2]
        taskmaster = brigade[1]
        brigade_activity = brigade[2:]
        try:
            while math.isnan(df[id + i][1]):
                brigade_activity.append(df[id + i][9])
                i += 1
        except TypeError:
            continue
        except IndexError:
            print('End')
        finally:
            brigade_info[taskmaster] = brigade_activity
            values_add(brigade_info, day)
    # brigade_info[brigade[1]]

sheet.batch_format(cell_format)
sheet.update_notes(cell_notes)
