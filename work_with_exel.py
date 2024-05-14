import pandas as pd

from typing import Any
from openpyxl import load_workbook
from PIL import Image
from io import BytesIO


def get_data_from_sheet(file_name: str, name_sheet: str) -> Any:
    """

    :param file_name:
    :param name_sheet:
    :return:
    """
    xl_file = pd.ExcelFile(file_name)
    return xl_file.parse(name_sheet)


def find_and_save_img_from_exel(file_name: str, name_sheet: str, names_img_file: list) -> bool:
    """
    Функция ищет и сохраняет в виде файла изображения, найденные в таблице.

    :param file_name: Имя exel файла
    :param name_sheet: Название нужной таблицы
    :param names_img_file: Список с именами для изображений в таблице
    :return: True - если изображение найдено, False - если нет.
    """
    xl_file = load_workbook(file_name)
    sheet = xl_file[name_sheet]

    find_img = False
    count_img = 0
    for row in sheet.iter_rows():
        for cell in row:
            if cell._image is not None:  # Если ячейка содержит изображение
                img = cell._image.img
                img_data = img._data

                # Преобразование данных изображения в байты
                img_bytes = BytesIO(img_data)
                # Открываем изображение с помощью Pillow
                image = Image.open(img_bytes)

                # Можно сохранить изображение или выполнить другие операции с ним
                image.save(f'{names_img_file[count_img]}.png')  # Сохраняем изображение с именем ячейки
                count_img += 1
                find_img = True

    return find_img
