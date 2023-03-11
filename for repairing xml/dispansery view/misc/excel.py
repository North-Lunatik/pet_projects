import xlrd
from datetime import datetime
from typing import Tuple


def get_last_show_up_date(sheet: xlrd.sheet.Sheet, row_index: int) -> str:
    """
    Возвращает дату последней явки пациента.
    
    Дата взятия на учет (A), Дата последней явки (B)
    Если В - пусто, берем А, если В есть, берем максимальное B.
    """
    first_show_up_date = datetime.strptime(sheet.cell(row_index, 9).value, '%d.%m.%Y')
    string_show_up_dates = sheet.cell(row_index, 11).value
    show_up_dates = []
    if string_show_up_dates:
        show_up_dates = [datetime.strptime(x, '%d.%m.%Y') for x in sheet.cell(row_index, 11).value.split('\n')]

    if show_up_dates:
        date = max(show_up_dates)
        return date.strftime('%d.%m.%Y')
    else:
        return first_show_up_date.strftime('%d.%m.%Y')


def get_data_from_report(filename: str) -> Tuple[dict]:
    """Возвращает записи отчета, в которых указана дата последней явки и вспомогательные данные."""
    report_data = {}
    phone_data = {}

    with xlrd.open_workbook_xls(str(filename)) as book:
        sheet = book.sheet_by_index(0)
        row_count = sheet.nrows
        
        for row_index in range(21, row_count-2):
            fio = sheet.cell(row_index, 1).value
            dr = sheet.cell(row_index, 3).value
            phone = sheet.cell(row_index, 5).value
            date_of_appearance = get_last_show_up_date(sheet, row_index)
            ds = str(sheet.cell(row_index, 2).value).split('. ')[0]
            if fio == '' and dr == '':
                break
            
            if date_of_appearance:
                report_data.setdefault(
                    (fio, datetime.strptime(dr, '%d.%m.%Y')), set()
                ).add(
                    (datetime.strptime(date_of_appearance, '%d.%m.%Y'), ds)
                )
                # Cохраняем дополнительную информацию.
                phone_data.setdefault(
                    (fio, datetime.strptime(dr, '%d.%m.%Y')), set()
                ).add(phone)
        
    return report_data, phone_data
