import os
import zipfile
import pandas as pd

from typing import List

from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader


def get_table_data(xlsx_file: str, name_sheet: str) -> pd.DataFrame:
    """
    Извлекает данные из указанной таблицы в файла Excel.

    Args:
    - xlsx_file (str): Путь к файлу Excel, из которого нужно извлечь данные.
    - name_sheet (str): Название листа в файле Excel, на котором находится 
    таблица.

    Returns:
    - pd.DataFrame: DataFrame с данными из указанной таблицы(библиотека pandas).

    Example:
    - extract_table_data("example.xlsx", "Sheet1")
    """

    return pd.ExcelFile(xlsx_file).parse(name_sheet)


def extract_all_images(xlsx_file: str, extract_dir: str) -> None:
    """
    Извлекает все изображения формата PNG и JPEG из указанного 
    XLSX-файла и сохраняет их в указанную директорию.
    
    Args:
    - xlsx_file (str): Путь к XLSX-файлу, из которого нужно извлечь изображения.
    - extract_dir (str): Путь к директории, в которую будут сохранены 
    извлеченные изображения.

    Returns:
    - Не возвращает никакого значения (None).

    Example:
    - extract_all_images('example.xlsx', 'extracted_images')
    """

    with zipfile.ZipFile(xlsx_file, 'r') as zip_ref:

        for file in zip_ref.namelist():
            if file.endswith((".png", ".jpeg")):
                zip_ref.extract(file, os.path.join(extract_dir))


def save_image_from_excel(xlsx_file: str, name_sheet: str, extract_dir: str) -> List[str]:
    """
    Ищет изображения в указанном файле Excel, на указанном листе и сохраняет их в указанную директорию.

    Args:
    - xlsx_file (str): - имя Excel файла, содержащего изображения
    - name_sheet (str): - название листа в Excel файле, на котором находятся изображения
    - extract_dir (str): - папка, куда будут сохранены изображения(если её нету, создаст)

    Returns:
    - list (str): список с путями к сохраненным изображениям

    Example:
    - save_image_from_excel("output_socio.xlsx", "Статистика1", "output")
    """

    xl_file = load_workbook(xlsx_file)
    sheet = xl_file[name_sheet]
    image_loader = SheetImageLoader(sheet)

    result = []
    count_img = 0

    for coordinate in image_loader._images:

        if not os.path.exists(extract_dir):
            os.makedirs(extract_dir)

        image = image_loader.get(coordinate)
        image_name = f'{extract_dir}/{name_sheet}_{count_img}.png'

        image.save(image_name)

        result.append(image_name)
        count_img += 1

    return result


def split_table_into_parts(data: pd.DataFrame,
                           headers: List[str] = ['C_plus', 'E_plus', 'КУО'],
                           group_table_indexes: List[str] = ['S_group', 'E_group', 'BB_group']
                           ) -> tuple[dict, dict]:
    # TODO: нужно ли кидать исключения, если headers != group_table_indexes? @nick-vivo

    """
    Принимает данные из DataFrame и разбивает их на три таблицы с указанными заголовками и индексами групп.

    Args:
    - data (pd.DataFrame): данные в формате DataFrame, которые требуется разбить на таблицы
    - headers (List[str]): список с заголовками таблиц (по умолчанию ['C_plus', 'E_plus', 'КУО'])
    - group_table_indexes (List[str]): список с индексами групповых данных (по умолчанию ['S_group', 'E_group', 'BB_group'])
    
    Returns:
    - Первый словарь содержит таблицы с данными, разбитыми по указанным заголовкам (headers)
    - Второй словарь содержит групповые данные по указанным индексам (group_table_indexes)
    
    Clarification:
    - Важно отметить, что функция сортирует данные в каждой таблице по убыванию и добавляет столбцы с индексами строк.
    """

    result = {}

    for i in range(0, len(headers)):
        table = data[data.columns[i + 1:i + 2]]
        table.insert(0, 'ID', range(1, len(table) + 1))

        table = table.sort_values(by=data.columns[i + 1], ascending=False)
        table.insert(0, 'п/п', range(1, len(table) + 1))

        result[headers[i]] = table

    group_table = {}

    for key in group_table_indexes:
        group_table[key] = data.at[0, key]

    return result, group_table
