import os
from openpyxl import load_workbook


def checking_files():
    path = f"{os.getcwd()}\\xlsx"
    files_list = os.listdir(path)
    print(f"\nВсего найдено файлов: {len(files_list)}")

    file = open(f'{os.getcwd()}\\none_cells.txt', 'w', encoding='utf-8')
    file.close()

    rows_list = []
    for file_name in files_list:
        print(f"\nИзвлечен файл {file_name}")

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

                if products_name_sheet.value is None and count_sheet.value is None and price_sheet.value is None:
                    break
                else:
                    rows_list.append(row)
                    with open(f'{os.getcwd()}\\none_cells.txt', 'a', encoding='utf-8') as file:
                        file.write(f'Файл: {file_name}, строка: {str(row)}\n')
                i += 1

    if rows_list:
        if input_cell() == 'next':
            checking_files()
    else:
        print('Проверка документа пройдена успешна, пустые поля отсутствуют.')
        return


def input_cell():
    menu = ['stop', 'help', 'next']
    input_data = input('Место для ввода: ')
    if input_data not in menu:
        print(f'\nВведены некорректные данные\nВозможные варианты для ввода:')
        print(*menu)
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
