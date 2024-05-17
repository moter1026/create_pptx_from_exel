import os

import pandas as pd

from typing import Any
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader
from PIL import Image
from io import BytesIO

from work_with_json import read_json_file


def get_data_from_sheet(file_name: str, name_sheet: str) -> Any:
    """
    Возвращает данные из exel файла file_name из таблицы name_sheet
    :param file_name: файл exel где будут искаться данные
    :param name_sheet: название таблицы, на которой необходимо найти данные
    :return:
    """
    xl_file = pd.ExcelFile(file_name)
    return xl_file.parse(name_sheet)


def find_and_save_img_from_exel(file_name: str, name_sheet: str) -> list:
    """
    Функция ищет и сохраняет в виде файла изображения, найденные в таблице.

    :param file_name: Имя exel файла
    :param name_sheet: Название нужной таблицы
    :return: Список с путями, куда были сохранены изображения
    """
    json_data = read_json_file("./files.json")

    xl_file = load_workbook(file_name)
    sheet = xl_file[name_sheet]

    # calling the image_loader
    image_loader = SheetImageLoader(sheet)

    result = []
    count_img = 0
    for coordinate in image_loader._images:
        if not os.path.exists(json_data["img_directory"]):
            os.makedirs(json_data["img_directory"])

        image = image_loader.get(coordinate)
        image_name = f"{json_data["img_directory"]}/{name_sheet}_img_№{count_img}.png"
        image.save(image_name)
        result.append(image_name)
        count_img += 1

    return result


def edit_data_from_sheet(data: pd.DataFrame()):
    """
    Разбивает данные со страницы на три таблицы
    """
    indexes = ['C_plus', 'E_plus', 'КУО']
    result = {}
    for i in range(0, len(indexes)):
        table = data[data.columns[i+1:i+2]]
        table.insert(0, 'ID', range(1, len(table) + 1))
        table = table.sort_values(by=data.columns[i+1], ascending=False)
        table.insert(0, 'п/п', range(1, len(table) + 1))
        result[indexes[i]] = table
    return result


