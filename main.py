import pandas as pd
from pptx import Presentation
from work_with_exel import get_data_from_sheet, edit_data_from_sheet
from pptx.util import Inches
from work_with_pptx import Present


def main():
    slides_for_stats = {'Статистика1': {'C_plus': 6, 'E_plus': 8, 'КУО': 10, 'group': 14},
                        'Статистика2': {'C_plus': 18, 'E_plus': 19, 'КУО': 20, 'group': 21},
                        'Статистика3': {'C_plus': 25, 'E_plus': 26, 'КУО': 27, 'group': 28},
                        'Статистика4': {'C_plus': 32, 'E_plus': 33, 'КУО': 34, 'group': 35},
                        'Статистика5': {'C_plus': 39, 'E_plus': 40, 'КУО': 41, 'group': 42},
                        'Статистика6': {'C_plus': 46, 'E_plus': 47, 'КУО': 48, 'group': 49}}

    pptx_file = 'result.pptx'
    template_file = 'pptx files/template.pptx'
    xlsx_file = 'exel files/data_socio.xlsx'
    prs = Present(pptx_file, 10, template_file)
    prs.save_titul_slide("Какой-то 11й группы")
    for key in slides_for_stats.keys():
        data = get_data_from_sheet(xlsx_file, key)
        data = edit_data_from_sheet(data)
        for stat_key in data.keys():
            prs.add_table_to_slide(data[stat_key], slides_for_stats[key][stat_key])
    prs.save()


if __name__ == '__main__':
    main()
