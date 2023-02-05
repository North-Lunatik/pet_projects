import os
from datetime import datetime
from itertools import cycle
from tkinter import BOTH, END, Tk, filedialog, messagebox
from tkinter.ttk import Button, Entry, Frame, Label
from typing import Tuple

import xlrd
from lxml import etree
from lxml.etree import Element, tostring


def get_data_from_report(filename) -> Tuple[dict]:
    """Возвращает записи отчета, в которых указана дата последней явки и вспомогательные данные."""
    report_data = {}
    phone_data = {}

    with xlrd.open_workbook_xls(str(filename)) as book:
        sheet = book.sheet_by_index(0)
        row_count = sheet.nrows
        
        for row_index in range(21, row_count):
            fio = sheet.cell(row_index, 1).value
            dr = sheet.cell(row_index, 3).value
            phone = sheet.cell(row_index, 5).value
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
                # сохраняем дополнительную информацию
                phone_data.setdefault(
                    (fio, datetime.strptime(dr, '%d.%m.%Y')), set()
                ).add(phone)
        
    return report_data, phone_data

def check_duplicates(file_path: str) -> Tuple[str, int]:
    """Возвращает сводку по наличию дубликатов в итоговой xml по параметрам fio, dr, ds"""
    result = {}

    tree = etree.parse(file_path)
    root = tree.getroot()
    for zap in root.findall('ZAP'):
        fio = f"{zap.find('FAM').text} {zap.find('IM').text} {zap.find('OT').text if zap.find('OT') is not None else ''}".strip()
        dr = datetime.fromisoformat(zap.find('DR').text)
        ds = zap.find('DS').text
        result.setdefault((fio, dr, ds), []).append(ds)

    duplicates = [(x, y) for x, y in result.items() if len(y) > 1]
    
    return f'Дубликатов в xml: {len(duplicates)}', len(duplicates)

def remove_duplicates(file_path: str) -> None:
    """Удаляет дубликаты из конечного файла с результатом."""
    result = {}

    tree = etree.parse(file_path)
    root = tree.getroot()
    for zap in root.findall('ZAP'):
        fio = f"{zap.find('FAM').text} {zap.find('IM').text} {zap.find('OT').text if zap.find('OT') is not None else ''}".strip()
        dr = datetime.fromisoformat(zap.find('DR').text)
        ds = zap.find('DS').text
        if (fio, dr, ds) not in result:
            result[(fio, dr, ds)] = [ds]
        else:
            root.remove(zap)

    with open(file_path, "w", encoding='cp1251', errors=None, newline='\r\n') as f:
        f.write(tostring(root, pretty_print=True, encoding='Windows-1251', xml_declaration=True).decode('cp1251'))
    print('Дубликаты удалены.')

def get_ot_data(ot_obj: Element) -> str:
    """Возвращает обработанное отчество."""
    if ot_obj is None:
        return ''

    ot = ot_obj.text
    if ot.upper() == 'НЕТ':
        return ''
    
    return ot

class Application(Frame):

    def __init__(self) -> None:
        super().__init__()
        self.xml_filepath = None
        self.report_filepath = None
        self.initUI()

    def initUI(self):
        """Инициализирует UI."""
        self.master.title('Костыль для подстановки диагнозов')
        self.pack(fill=BOTH, expand=True)

        # Путь до отчета
        report_field_label = Label(self, text="Файл отчета")
        report_field_label.grid(row=1, column=1, padx=5, pady=5)
        self.report_field = Entry(self, width=60)
        self.report_field.grid(row=1, column=2)
        select_report_button = Button(self, text="...", command=self.open_report_file)
        select_report_button.grid(row=1, column=3)

        # Путь до XML
        xml_field_label = Label(self, text="Файл XML")
        xml_field_label.grid(row=2, column=1, padx=5, pady=5)
        self.xml_field = Entry(self, width=60)
        self.xml_field.grid(row=2, column=2)
        select_xml_button = Button(self, text="...", command=self.open_xml_file)
        select_xml_button.grid(row=2, column=3)

        run_button = Button(self, text="Преобразовать", command=self.rebuild_xml)
        run_button.grid(row=3, column=2)

    def open_report_file(self):
        """Устанавливает путь к отчету."""
        self.report_filepath = filedialog.askopenfilename()
        self.report_field.delete("0", END)
        self.report_field.insert(0, self.report_filepath)

    def open_xml_file(self):
        """Устанавливает путь к xml."""
        self.xml_filepath = filedialog.askopenfilename()
        self.xml_field.delete("0", END)
        self.xml_field.insert(0, self.xml_filepath)

    def rebuild_xml(self):
        """Обрабатываем файл."""
        if not self.report_filepath:
            messagebox.showerror("Ошибка", "Не выбран файл с отчетом.")
        if not self.xml_filepath:
            messagebox.showerror("Ошибка", "Не выбран xml файл.")
        report_data = None
        phone_data = None
        xml_file_path = self.xml_filepath
        report_data, phone_data = get_data_from_report(str(self.report_filepath))

        # фиксируем позицию элементов
        prepared_report_data = {}
        for x, y in report_data.items():
            data = list(y)
            data.sort()
            prepared_report_data[x] = cycle(data)

        tree = etree.parse(xml_file_path)
        root = tree.getroot()
        for zap in root.findall('ZAP'):
            if zap.find('DISP_TYP').text != '3':
                zap.getparent().remove(zap)
            else:
                fio = f"{zap.find('FAM').text} {zap.find('IM').text} {get_ot_data(zap.find('OT'))}".strip()
                dr = datetime.fromisoformat(zap.find('DR').text)
                dr_str = dr.strftime('%Y-%m-%d')
                prev = datetime.fromisoformat(zap.find('DAT_PREV').text).date() if zap.find('DAT_PREV').text else None

                if not zap.find('DS').text:
                    rd = prepared_report_data.get((fio, dr), None)
                    if rd:
                        date_prev, ds = next(rd)
                        zap.find('DS').text = ds
                        zap.find('DAT_PREV').text = date_prev.strftime('%Y-%m-%d')
                        phones = [x for x in phone_data.get((fio, dr), []) if x != '']
                        # Если телефон указан в отчете
                        if phones:
                            zap.find('PHONE').text = phones[0]
                    else:
                        print(f'Для {fio} {dr_str} не найдено данных в отчете')

        result_path = os.path.join(os.getcwd(), 'result.xml')
        with open(result_path, "w", encoding='cp1251', errors=None, newline='\r\n') as f:
            f.write(tostring(root, pretty_print=True, encoding='Windows-1251', xml_declaration=True).decode('cp1251'))

        result_text, count = check_duplicates(result_path)
        print(result_text)
        if count > 0:
            remove_duplicates(result_path)

        print('Готово.')


def main():

    root = Tk()
    app = Application()
    root.mainloop()


if __name__ == '__main__':
    main()