import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import numpy as np
from pprint import pprint

scope = ['https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive"]

credentials = ServiceAccountCredentials.from_json_keyfile_name("stim-downtime-credentials.json", scope)
client = gspread.authorize(credentials)

spreadsheet = client.open_by_key("153loT8FiuqTjECvGdFFpgyjcu5GuctrCrgO0RTELi5U")


def google_table_read():
    sheet = spreadsheet.worksheet('РБ')
    columns_to_keep = ['Контракт', 'Вид работ', 'Норма расхода(контракт)', 'Норма расхода(план)']
    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])
    df = df[columns_to_keep]
    df.replace('', np.nan, inplace=True)
    df['Контракт'].fillna(method='ffill', inplace=True)
    for region in spreadsheet.worksheets()[1:]:
        sheet = spreadsheet.worksheet(region.title)
        data = sheet.get_all_values()
        df_next = pd.DataFrame(data[1:], columns=data[0])
        df_next = df_next[columns_to_keep]
        df_next.replace('', np.nan, inplace=True)
        df_next['Контракт'].fillna(method='ffill', inplace=True)
        df = pd.concat([df, df_next])

# google_table_read()


def csv_edit():
    file_name = "Экономисты/square&material.csv"
    data = pd.read_csv(file_name)
    data.drop(columns="External ID", inplace=True)
    data['Бригада/Отображаемое Имя'].fillna(method='ffill', inplace=True)
    data['Начало смены'] = pd.to_datetime(data['Начало смены'])
    data['Начало смены'].fillna(method='ffill', inplace=True)
    data['Начало смены'] = data['Начало смены'].dt.strftime('%e')
    data['Активности/Контракты/Субрегион'].fillna(method='ffill', inplace=True)
    data['Event Prod Item/Контракт/Отображаемое Имя'].fillna(method='ffill', inplace=True)
    brigades = data['Бригада/Отображаемое Имя'].drop_duplicates()
    dates = data['Начало смены'].drop_duplicates()

    brigades_materials_used_info = {}

    for brigade in brigades:
        brigades_materials_used_info[brigade] = {}
        for date in dates:
            brigade_day_info = data[data["Бригада/Отображаемое Имя"] == brigade][data["Начало смены"] == date]
            category_work_materials = brigade_day_info["Event Prod Item/Категория материала/Отображаемое Имя"].drop_duplicates()
            category_out_materials = brigade_day_info["Material stock picking out items/Категория материала/Отображаемое Имя"].drop_duplicates()
            type_of_works = brigade_day_info["Event Prod Item/Вид работ"].drop_duplicates().dropna()
            brigades_materials_used_info[brigade][date] = {}
            brigade_all_category_material_count = 0
            for work_material in category_work_materials:
                for out_material in category_out_materials:
                        try:
                            if work_material in out_material:
                                for work_type in type_of_works:
                                    brigade_one_category_material_count = brigade_day_info[data["Material stock picking out items/Категория материала/Отображаемое Имя"] == out_material]["Material stock picking out items/Количество в базовых"].sum()
                                    brigade_material_work_square = brigade_day_info[data["Event Prod Item/Категория материала/Отображаемое Имя"] == work_material][data["Event Prod Item/Вид работ"] == work_type]["Event Prod Item/Площадь, м²"].sum()
                                    try:
                                        if work_material in brigades_materials_used_info[brigade][date][work_type]:
                                            brigade_all_category_material_count += brigade_one_category_material_count
                                    except KeyError as e:
                                        brigades_materials_used_info[brigade][date][work_type] = {}
                                        brigade_all_category_material_count = brigade_one_category_material_count
                                    brigade_flow_rate_metr = brigade_all_category_material_count/brigade_material_work_square
                                    brigades_materials_used_info[brigade][date][work_type][work_material] = brigade_flow_rate_metr.round(2)

                            if out_material == "Все / Материалы / Добавка / СШ крупные" or out_material == "Все / Материалы / Добавка / СШ мелкие":
                                    brigade_one_category_material_count = brigade_day_info[data["Material stock picking out items/Категория материала/Отображаемое Имя"] == out_material]["Material stock picking out items/Количество в базовых"].sum()
                                    brigade_material_work_square = brigade_day_info[data["Event Prod Item/Категория материала/Отображаемое Имя"] == work_material]["Event Prod Item/Площадь, м²"].sum()
                                    brigade_flow_rate_metr = brigade_one_category_material_count / brigade_material_work_square
                                    brigades_materials_used_info[brigade][date][out_material] = brigade_flow_rate_metr.round(2)
                        except TypeError as e:
                            continue
    pprint(brigades_materials_used_info)











csv_edit()