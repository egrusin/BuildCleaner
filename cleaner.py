# Подстановщик черточек в номер здания

import re
import click
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.reader.excel import load_workbook

def format_build(cell: Cell) -> None:
    """Преобразует ячейку [1а] к виду [1/а] или ничего не делает"""
    value = str(cell.value).strip()
    regexp = r'[0-9]+[а-яА-Я]+'
    res = re.fullmatch(regexp, value)
    if res is None:
        return
    else:
        cell.value = value[:-1] + '/' + value[-1]
          
@click.command()
@click.argument('path_to_workbook')
@click.option('--worksheet', '-s', default='active')
@click.option('--build-column', '-c', default='')
def main(path_to_workbook, worksheet, build_column):
    # Открываем книгу эксель:
    book: Workbook = load_workbook(path_to_workbook)

    # Определяем рабочий лист:
    sheet: Worksheet = book[worksheet] if worksheet != 'active' else book.active

    # Определяем индекс столбца со значениями номера дома:
    column_index: int = [i.value for i in sheet['1']].index(build_column) if build_column != '' else 0

    # Форматируем значение ячеек в столбце:
    i = 1
    while True:
        row = sheet[str(i)]
        build = row[column_index]

        # Остановка цикла
        if build.value is None:
            break
        
        # Меняем значения
        format_build(build)
        i += 1
    
    # Сохраняем книгу
    book.save(path_to_workbook)

if __name__ == "__main__":
    main()