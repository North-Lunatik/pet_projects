import os
from datetime import datetime
from itertools import cycle
from tkinter import (BOTH, END, Checkbutton, IntVar, Tk, W, filedialog,
                     messagebox)
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import Button, Entry, Frame, Label
from typing import List, Tuple, Union

from lxml import etree
from lxml.etree import tostring

from misc.excel import get_data_from_report
from misc.utils import clean_patronymic, clean_phone

help = """
Программа предназначена для исправления данных в некорректно сформированном xml по ДН.

Информацию предполагается брать из отчета: `Список пациентов, запланированных для диспансерного наблюдения`

Файл отчета xls формируется несколько некорректно. 
Для исправления можно выполнить "Сохранить как" в формате excel 98/2003 и после этого выбирать для обработки.
    
"""


def check_duplicates(file_path: str) -> Tuple[str, int]:
    """Возвращает сводку по наличию дубликатов в итоговой xml по параметрам fio, dr, ds."""
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

class Application(Frame):

    def __init__(self) -> None:
        super().__init__()
        self.xml_filepath = None
        self.report_filepath = None
        self.console = None
        self.ds_from_168n = None
        self.initUI()

    def initUI(self) -> None:
        """Инициализирует UI."""
        self.master.title('Костыль для подстановки недостающих данных по ДН.')
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

        # Признак для отключения удаления лишних типов данных
        self.is_not_remove_other_data = IntVar(value=0)
        self.cb = Checkbutton(
            self, text="Оставить в xml записи помимо ДН?",
            variable=self.is_not_remove_other_data,
        )
        self.cb.grid(row=3, column=1, columnspan=2, sticky=W, ipadx=30)

        # Признак для фильтрации диагнозов по приазу 168Н
        self.filtered_by_ds_from168n = IntVar(value=0)
        self.cb2 = Checkbutton(
            self, text="Оставить для ДН только записи с диагнозами из приказа 168Н",
            variable=self.filtered_by_ds_from168n,
        )
        self.cb2.grid(row=4, column=1, columnspan=2, sticky=W, ipadx=30)

        # Имя файла
        package_number_field_label = Label(self, text="Имя файла")
        package_number_field_label.grid(row=5, column=1, padx=5, pady=5)
        self.package_number_field = Entry(self, width=60)
        self.package_number_field.grid(row=5, column=2)
        package_number_example_label = Label(self, text="D-M<Код МО>-F35-<Год>-<Номер пакета>, пример: D-M352530-F35-2023-1")
        package_number_example_label.grid(row=6, column=1, columnspan=2)

        run_button = Button(self, text="Преобразовать", command=self.rebuild_xml)
        run_button.grid(row=7, column=2)
        help_button = Button(self, text="Справка", command=self.show_help)
        help_button.grid(row=7, column=3)

        # Виджет для отображаения результатов обработки
        self.console = ScrolledText(self, height=10)
        self.console.grid(row=8, column=1, columnspan=3)

    def to_console(self, msgs: Union[str, List[str]]) -> None:
        """Добавляет строку в виджет вывода результатов на экран."""
        if self.console:
            if isinstance(msgs, str):
                msgs = [msgs]
            
            self.console.update_idletasks()
            for msg in msgs:
                self.console.insert("end", msg + '\n')

    def open_report_file(self) -> None:
        """Устанавливает путь к отчету."""
        self.report_filepath = filedialog.askopenfilename()
        self.report_field.delete("0", END)
        self.report_field.insert(0, self.report_filepath)

    def open_xml_file(self) -> None:
        """Устанавливает путь к xml."""
        self.xml_filepath = filedialog.askopenfilename()
        self.xml_field.delete("0", END)
        self.xml_field.insert(0, self.xml_filepath)

    def show_help(self) -> None:
        """Выводит диалоговое окно со справкой по использованию."""
        messagebox.showinfo('Справка по использованию.', help)
    
    def is_allow_remove(self) -> bool:
        """Возвращает признак отражающий разрешение удалять данные с DISP_TYPE == 3."""
        return bool(not self.is_not_remove_other_data.get())

    def rebuild_xml(self) -> None:
        """Обрабатывает файл."""
        if not self.report_filepath:
            messagebox.showerror("Ошибка", "Не выбран файл отчета.")
            return
        if not self.xml_filepath:
            messagebox.showerror("Ошибка", "Не выбран xml файл.")
            return
        
        report_data = None
        phone_data = None
        xml_file_path = self.xml_filepath

        self.to_console('Обработка файла отчета...')
        report_data, phone_data = get_data_from_report(str(self.report_filepath))
        self.to_console('Завершено.')

        # фиксируем позицию элементов
        prepared_report_data = {}
        for x, y in report_data.items():
            data = list(y)
            data.sort()
            prepared_report_data[x] = cycle(data)

        self.to_console('Обработка xml файла...')
        tree = etree.parse(xml_file_path)
        root = tree.getroot()
        for zap in root.findall('ZAP'):
            if zap.find('DISP_TYP').text != '3':
                if self.is_allow_remove():
                    zap.getparent().remove(zap)
            else:
                fio = f"{zap.find('FAM').text} {zap.find('IM').text} {clean_patronymic(zap.find('OT'))}".strip()
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
                            zap.find('PHONE').text = clean_phone(phones[0])
                    else:
                        self.to_console(f'Для {fio} {dr_str} не найдено данных в отчете')

        if self.filtered_by_ds_from168n.get():
            self.to_console(f'Фильтруем записи по приказу 168 Н...')

            if self.ds_from_168n is None:
                from decr_168n_15_03_2022 import get_all_diagnoses
                self.ds_from_168n = get_all_diagnoses()

            removing_counter = 0
            records = root.findall('ZAP')
            for zap in records:
                if zap.find('DISP_TYP').text == '3':
                    if str(zap.find('DS').text).strip().upper() not in self.ds_from_168n:
                        zap.getparent().remove(zap)
                        removing_counter += 1

            self.to_console(f'Фильтрация завершена. Удалено {removing_counter} записей из {len(records)}')
        
        custom_filename = self.package_number_field.get()
        if custom_filename:
            custom_filename = custom_filename.upper()
            
            # D-M352530-F35-2023-1
            filename_elements = custom_filename.split('-')

            result_path = os.path.join(os.getcwd(), f'{custom_filename}.xml')
            root.find('ZGLV').find('FILENAME').text = custom_filename
            
            build_date = datetime.now()
            root.find('ZGLV').find('DATA').text = build_date.strftime('%Y-%m-%d')

            root.find('ZGLV').find('CODE_MO').text = filename_elements[1][1:]

            root.find('ZGLV').find('YEAR').text = filename_elements[3]
            root.find('ZGLV').find('R').text = filename_elements[4]
        else:
            result_path = os.path.join(os.getcwd(), 'result.xml')
        
        with open(result_path, "w", encoding='cp1251', errors=None, newline='\r\n') as f:
            f.write(tostring(root, pretty_print=True, encoding='Windows-1251', xml_declaration=True).decode('cp1251'))
        
        self.to_console(['Завершено.', '', 'Поиск дубликатов в xml...'])
        result_text, count = check_duplicates(result_path)
        if count > 0:
            self.to_console(result_text)
            remove_duplicates(result_path)
            self.to_console('Дубликаты удалены.\n')

        self.to_console('Поиск записей без диагноза:')
        i = 0
        for zap in root.findall('ZAP'):
            if zap.find('DISP_TYP').text != '3' and not zap.find('DS').text:
                zap.getparent().remove(zap)
                i += 1
        if i:
            self.to_console(f'Найдено и удалено {i} записей.')
        else:
            self.to_console(f'Записей не найдено.\n')

        self.to_console(f'Результат находится в `{result_path}`')
        self.to_console(['', 'Готово.'])


def start_app() -> None:
    """Стартует приложение tk."""
    root = Tk()
    app = Application()
    root.mainloop()

if __name__ == '__main__':
    start_app()
