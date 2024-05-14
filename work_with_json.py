import json


def read_json_file(json_file_name) -> dict:
    """
    Читает JSON файл и возвращает словарь с его данными
    """
    try:
        with open(json_file_name, "r", encoding="utf-8") as JSON_file:
            data = json.load(JSON_file)

        return data

    except FileNotFoundError:
        print(f"\nНеудача при открытии файла '{json_file_name}',"
              f" в функции read_csv_frequency\n")
    except Exception as e:
        print(f"\nПроизошла ошибка {e} в функции read_json_file\n")
