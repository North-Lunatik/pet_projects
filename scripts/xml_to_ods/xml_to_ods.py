import os
from datetime import datetime

import odsgenerator
from lxml import etree
from progressbar import ProgressBar

#
# Конвертер xml файла с прикрепленным населением в таблицу формата ods
# Файл собирается не особенно шустро, где-то примерно 5000 строк / минуту
#

HEADER = ['Фамилия', 'Имя', 'Отчество', 'Дата рождения', 'ЕНП']
FIELDS = ['FAM', 'IM', 'OT', 'DR', 'NPOLIS']
WIDTH_LIST = [2500, 2500, 2900, 2500, 3500]

output_filename = ''
for filename in os.listdir(os.getcwd()):
    if filename.upper().startswith('PRKS') and filename.upper().endswith('XML'):
        tree = etree.parse(os.path.join(os.getcwd(), filename))
        root = tree.getroot()
        sheetdate = root.find('ZGLV').find('DATE').text
        output_filename = f"Население на {datetime.fromisoformat(sheetdate).strftime('%d.%m.%Y')}"
        
        rows = []

        persons = root.findall('PERS')
        bar = ProgressBar(min_value=1, max_value=len(persons))
        
        print('Обрабатываем данные')
        for i, pers in enumerate(persons, start = 1):
            rows.append(
                {
                    'row': ['' if pers.find(field) is None else pers.find(field).text for field in FIELDS],
                    'style': 'grid_06pt'
                }
            )
            bar.update(i)
        
        bar.finish()
        
        print('Ожидаем сборки файла...')
        raw = odsgenerator.ods_bytes(
            [
                {
                    "name": output_filename,
                    "width": WIDTH_LIST,
                    "table": [
                        {
                            "row": HEADER,
                            "style": "bold_center_grid_06pt",
                        },
                        *rows
                    ],
                }
            ]
        )
        
        print('Файл собран, ожидаем запись...')
        with open(os.path.join(os.getcwd(), f'{output_filename}.ods'), "wb") as f:
            f.write(raw)
            print('Готово.')
        break
else:
    print('Ошибка: При запуске не найдено ни одного файла для обработки')
