import pandas as pd

from pptx import Presentation
from pptx.util import Inches

import work_with_json

from work_with_exel import (get_data_from_sheet, edit_data_from_sheet,
                            find_and_save_img_from_exel)
from work_with_pptx import Present


def main():
    slides_for_stats = {'Статистика1': {'C_plus': 6, 'E_plus': 8, 'КУО': 10, 'group': 14},
                        'Статистика2': {'C_plus': 18, 'E_plus': 19, 'КУО': 20, 'group': 21},
                        'Статистика3': {'C_plus': 25, 'E_plus': 26, 'КУО': 27, 'group': 28},
                        'Статистика4': {'C_plus': 32, 'E_plus': 33, 'КУО': 34, 'group': 35},
                        'Статистика5': {'C_plus': 39, 'E_plus': 40, 'КУО': 41, 'group': 42},
                        'Статистика6': {'C_plus': 46, 'E_plus': 47, 'КУО': 48, 'group': 49}}
    slides_for_graphs = {'Статистика1': [15, 16],
                         'Статистика2': [22, 23],
                         'Статистика3': [29, 30],
                         'Статистика4': [36, 37],
                         'Статистика5': [43, 44],
                         'Статистика6': [50, 51]}

    json_data = work_with_json.read_json_file("./files.json")

    pptx_file = json_data["pptx_out_file"]
    template_file = json_data["template_file"]
    xlsx_file = json_data["exel_in_file"]

    prs = Present(pptx_file, 10, template_file)
    prs.save_titul_slide("Какой-то 11й группы")

    for key in slides_for_stats.keys():
        data = get_data_from_sheet(xlsx_file, key)
        images_paths = find_and_save_img_from_exel(json_data["exel_in_file"], key)
        data = edit_data_from_sheet(data)
        for stat_key in data.keys():
            prs.add_table_to_slide(data[stat_key], slides_for_stats[key][stat_key])

        ind = 0
        for path in images_paths:
            prs.add_image_to_slide(path, slides_for_graphs[key][ind])
            ind += 1

    prs.save()


if __name__ == '__main__':
    main()
