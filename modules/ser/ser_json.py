import json


def read_json(file_path: str) -> dict:

    """
    Читает данные из JSON файла и возвращает их в виде словаря.

    Args:
    - file_path (str): Путь к JSON файлу, который нужно прочитать.

    Returns:
    - dict: Словарь, содержащий данные из JSON файла.

    Raises:
    - FileNotFoundError: Если файл не найден по указанному пути.
    - JSONDecodeError: Если происходит ошибка декодирования JSON данных.

    Example:
    >>> read_json('data.json')
    {'name': 'Alice', 'age': 30, 'city': 'New York'}
    """

    with open(file_path, 'r', encoding="utf-8") as file:
        data = json.load(file)
    return data
