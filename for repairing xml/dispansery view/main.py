import os
from datetime import datetime
from itertools import cycle
from pathlib import Path

import xlrd
from lxml import etree
from lxml.etree import tostring


def get_data_from_report(filename):
    """Возвращает записи отчета, в которых указана дата последней явки."""
    report_data = {}

    with xlrd.open_workbook_xls(str(filename)) as book:
        sheet = book.sheet_by_index(0)
        row_count = sheet.nrows
        
        for row_index in range(21, row_count):
            fio = sheet.cell(row_index, 1).value
            dr = sheet.cell(row_index, 3).value
            date_of_appearance = sheet.cell(row_index, 10).value.split('\n')[0]
            ds = str(sheet.cell(row_index, 2).value).split('. ')[0]
            if fio == '' and dr == '':
                break
            
            if date_of_appearance:
                report_data.setdefault(
                    (fio, datetime.strptime(dr, '%d.%m.%Y')), set()
                ).add(
                    (datetime.strptime(date_of_appearance, '%d.%m.%Y'), ds)
                )
        
    return report_data

def check_duplicates(file_path: str) -> str:
    """Возвращает сводку по наличию дубликатов в итоговой xml по параметрам fio, dr, ds"""
    result = {}

    tree = etree.parse(xml_file_path)
    root = tree.getroot()
    for zap in root.findall('ZAP'):
        fio = f"{zap.find('FAM').text} {zap.find('IM').text} {zap.find('OT').text if zap.find('OT') is not None else ''}".strip()
        dr = datetime.fromisoformat(zap.find('DR').text)
        ds = zap.find('DS').text
        result.setdefault((fio, dr), []).append(ds)

    return f'Дубликатов: {len([x for x, y in result.items() if len(y) > 1])}'

if __name__ == '__main__':
    report_data = None
    xml_file_path = None
    for filename in Path(os.getcwd()).iterdir():
        if filename.name.endswith('_list_disp_observ_pg.xls'):
            report_data = get_data_from_report(str(filename))
        elif filename.name.endswith('.xml'):
            xml_file_path = str(filename)
        
        if report_data and xml_file_path:
            break

    prepared_report_data = {x: cycle(y) for x, y in report_data.items()}

    tree = etree.parse(xml_file_path)
    root = tree.getroot()
    for zap in root.findall('ZAP'):
        if zap.find('DISP_TYP').text != '3':
            zap.getparent().remove(zap)
        else:
            fio = f"{zap.find('FAM').text} {zap.find('IM').text} {zap.find('OT').text if zap.find('OT') is not None else ''}".strip()
            dr = datetime.fromisoformat(zap.find('DR').text)
            dr_str = dr.strftime('%Y-%m-%d')
            prev = datetime.fromisoformat(zap.find('DAT_PREV').text).date() if zap.find('DAT_PREV').text else None

            if not zap.find('DS').text:
                rd = prepared_report_data.get((fio, dr), None)
                if rd:
                    date_prev, ds = next(rd)
                    zap.find('DS').text = ds
                    zap.find('DAT_PREV').text = date_prev.strftime('%Y-%m-%d')
                else:
                    print(f'Для {fio} {dr_str} не найдено данных в отчете')

    result_path = os.path.join(os.getcwd(), 'result.xml')
    with open(result_path, "w", encoding='cp1251', errors=None, newline='\r\n') as f:
        f.write(tostring(root, pretty_print=True, encoding='Windows-1251', xml_declaration=True).decode('cp1251'))

    print(check_duplicates(result_path))
    print('Готово.')