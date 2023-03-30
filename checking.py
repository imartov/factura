import os
from openpyxl import load_workbook


def checking_files():
    path = f"{os.getcwd()}\\xlsx"
    files_list = os.listdir(path)
    # print(f"\nВсего найдено файлов: {len(files_list)}")

    none_rows = ''
    for file_name in files_list:
        if file_name[0] == '~':
            print(f'Пропущен временный файл {file_name}')
            continue
        else:
            print(f"Извлечен файл {file_name}")

            # getting sheets in every excel file
            wb = load_workbook(filename=f"{path}\\{file_name}")
            print(f"Всего найдено листов: {len(wb.sheetnames)}")

            for sheet_name in wb.sheetnames:
                sheet_active = wb[sheet_name]
                print(f'Проверка листа {sheet_name}\n')

                i = 0
                for row in sheet_active.iter_rows(min_row=12, max_col=3):
                    products_name_sheet = row[0]
                    count_sheet = row[1]
                    price_sheet = row[2]

                    if not products_name_sheet.value and not count_sheet.value and not price_sheet.value:
                        break
                    elif products_name_sheet.value and count_sheet.value and price_sheet.value:
                        pass
                    else:
                        print(f'Недопустимые значения строки {row}')
                        none_rows += f'Файл: {file_name}, лист: {sheet_name}, строка: ' + str(row).replace('(', '').replace(')', ';\n')
                    i += 1

    if none_rows:
        with open(f'{os.getcwd()}\\none_cells.txt', 'w', encoding='utf-8') as file:
            file.write(none_rows)
        print(f'Найдены пустые значения в следующих строках:\n{none_rows}')
        print('\nЗаполните поля и введите "next" для повторной проверки')
        if input_cell() == 'next':
            checking_files()
    else:
        print('Проверка документа пройдена успешна, пустые поля отсутствуют.')
        return


def input_cell():
    menu = ['next']
    input_data = input('Место для ввода: ')
    if input_data not in menu:
        print(f'\nВведены недопустимые данные\nВозможные варианты для ввода:')
        print(', '.join(menu))
        input_cell()
    return input_data


def waiting_for_checking():
    print('Выполнение программы приостановлено.\nДля возобновления введите "next"\nДля остановки введите "stop"')
    input_data = input()
    if input_data != 'next' and input_data != 'stop':
        print('Проверьте корректность введенных данных и повторите ввод')
        waiting_for_checking()
    return input_data


def main():
    checking_files()
    # input_cell()
    # waiting_for_checking()


if __name__ == "__main__":
    main()
