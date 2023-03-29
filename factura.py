from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.support.ui import Select
import os
from openpyxl import load_workbook
from fake_useragent import UserAgent
from fake_useragent import UserAgent
import keys
from login_facturowo import input_login, get_saved_login_data, get_temp_login_data
from tqdm import tqdm, trange
from googletrans import Translator
from checking import *


def get_factura():
    if "saved_login.json" in os.listdir(os.getcwd()):
        pass
    else:
        input_login()

    # user_agent = UserAgent()

    # options = webdriver.ChromeOptions()
    # options.add_argument(f"user-agent={user_agent.random}")
    try:
        # driver = webdriver.Chrome(options=options)
        driver = webdriver.Chrome()
        driver.maximize_window()
        url = "https://www.fakturowo.pl/logowanie"
        driver.get(url)
    except Exception as ex:
        print(f"{ex}\n{keys.error_message_driver}\n{keys.error_message_stop_running}.")
        return

    sleep(3)

    # login
    login_mail = driver.find_element(By.XPATH, f"//input[@name='email']")
    login_mail.send_keys(Keys.CONTROL, "a")
    login_mail.send_keys(Keys.DELETE)

    login_password = driver.find_element(By.XPATH, f"//input[@name='haslo']")
    login_password.send_keys(Keys.CONTROL, "a")
    login_password.send_keys(Keys.DELETE)

    if "saved_login.json" in os.listdir(os.getcwd()):
        email, password = get_saved_login_data()
        login_mail.send_keys(email)
        login_password.send_keys(password)

    else:
        email, password = get_temp_login_data()
        login_mail.send_keys(email)
        login_password.send_keys(password)

    button_login = driver.find_element(By.XPATH, f"//button[@name='login']")
    button_login.click()

    sleep(2)

    # if entered login data are wrong
    err = False
    while not err:
        try:
            driver.find_element(By.XPATH, "//div[@class='alert alert-danger']")
            print("\nThe entered login data are wrong. Please repeat to enter login data")
            input_login()
            sleep(1)

            login_mail = driver.find_element(By.XPATH, f"//input[@name='email']")
            login_mail.send_keys(Keys.CONTROL, "a")
            login_mail.send_keys(Keys.DELETE)

            login_password = driver.find_element(By.XPATH, f"//input[@name='haslo']")
            login_password.send_keys(Keys.CONTROL, "a")
            login_password.send_keys(Keys.DELETE)

            if "saved_login.json" in os.listdir(os.getcwd()):
                email, password = get_saved_login_data()
                login_mail.send_keys(email)
                login_password.send_keys(password)
            else:
                email, password = get_temp_login_data()
                login_mail.send_keys(email)
                login_password.send_keys(password)

            button_login = driver.find_element(By.XPATH, f"//button[@name='login']")
            button_login.click()
        except Exception as ex:
            err = True

    # delete temp login and password
    try:
        os.remove(f"{os.getcwd()}\\temp_login.json")
    except Exception:
        pass

    # add button for creating new document
    new_document_button = driver.find_element(By.XPATH, f"//a[@class='btn btn-xs btn-primary']")
    new_document_button.click()

    sleep(1.5)

    # getting input excel files
    path = f"{os.getcwd()}\\xlsx"
    files_list = os.listdir(path)
    print(f"\nВсего найдено файлов: {len(files_list)}.")

    for file in files_list:
        print(f"\nИзвлечен файл {file}.")

        # getting sheets in every excel file
        wb = load_workbook(filename=f"{path}\\{file}")
        print(f"Всего найдено листов: {len(wb.sheetnames)}.")

        for sheet_name in wb.sheetnames:

            # select active sheet
            sheet_active = wb[sheet_name]

            # select data from sheet from excel file
            document_type_sheet = str(sheet_active['B1'].value)
            buyer_name_sheet = str(sheet_active['B2'].value)
            buyer_nip_sheet = str(sheet_active['B3'].value)
            buyer_address_sheet = str(sheet_active['B4'].value)
            buyer_place_sheet = str(sheet_active['B5'].value)
            tax_sheet = str(sheet_active['B6'].value)
            currency_sheet = str(sheet_active['B7'].value)
            measure_sheet = str(sheet_active['B8'].value)

            # filling field 'Dokument'
            document_type_select = Select(driver.find_element(By.ID, 'rodzaj'))
            document_type_select.select_by_visible_text(document_type_sheet)

            # filling fields about buyer
            buyer_name_field = driver.find_element(By.NAME, "nabywca[nazwa]")
            buyer_name_field.send_keys(Keys.CONTROL, "a")
            buyer_name_field.send_keys(Keys.DELETE)
            buyer_name_field.send_keys(buyer_name_sheet)

            sleep(0.5)

            buyer_nip_field = driver.find_element(By.NAME, "nabywca[nip]")
            buyer_nip_field.send_keys(Keys.CONTROL, "a")
            buyer_nip_field.send_keys(Keys.DELETE)
            buyer_nip_field.send_keys(buyer_nip_sheet)

            sleep(0.5)

            buyer_address_field = driver.find_element(By.ID, 'ulica_nabywca')
            buyer_address_field.send_keys(Keys.CONTROL, "a")
            buyer_address_field.send_keys(Keys.DELETE)
            buyer_address_field.send_keys(buyer_address_sheet)

            sleep(0.5)

            buyer_place_field = driver.find_element(By.ID, 'miasto_nabywca')
            buyer_place_field.send_keys(Keys.CONTROL, "a")
            buyer_place_field.send_keys(Keys.DELETE)
            buyer_place_field.send_keys(buyer_place_sheet)

            sleep(0.5)

            # filling field 'Waluta'
            currency_field = Select(driver.find_element(By.ID, 'waluta'))
            currency_field.select_by_visible_text(currency_sheet)

            sleep(1)

            # for row in tqdm(range(1, sheet_active.max_column, 1), desc=sheet_name, unit="product", dynamic_ncols=True):

            new_position = True
            i = 0
            for row in sheet_active.iter_rows(min_row=12, max_col=3, values_only=True):
                if row[0]:

                    # select data about products
                    product_name_sheet_ru = str(row[0])
                    translator = Translator()
                    product_name_sheet_pl = translator.translate(text=product_name_sheet_ru, src='ru', dest='pl')

                    count_sheet = str(row[1])
                    price_sheet = str(row[2])

                    # filling field about product
                    product_name_site = driver.find_element(By.ID, f"nazwa_{i}")
                    product_name_site.send_keys(Keys.CONTROL, "a")
                    product_name_site.send_keys(Keys.DELETE)
                    product_name_site.send_keys(product_name_sheet_pl.text.capitalize())

                    sleep(0.5)

                    # select option of measure
                    select_measure = Select(driver.find_element(By.ID, f"jm_{i}"))
                    select_measure.select_by_visible_text(measure_sheet)

                    sleep(0.5)

                    # select field of count
                    count_site = driver.find_element(By.ID, f"ilosc_{i}")
                    count_site.send_keys(Keys.CONTROL, "a")
                    count_site.send_keys(Keys.DELETE)
                    count_site.send_keys(count_sheet)

                    sleep(0.5)

                    # select field of price
                    price_netto_site = driver.find_element(By.XPATH, f"//input[@id='cena_netto_{i}']")
                    price_netto_site.send_keys(Keys.CONTROL, "a")
                    price_netto_site.send_keys(Keys.DELETE)
                    price_netto_site.send_keys(price_sheet)

                    sleep(0.5)

                    # filling field 'Stawka VAT'
                    tax_field = Select(driver.find_element(By.ID, f'stawka_vat_{i}'))
                    if document_type_sheet == 'Eksport towarów (poza UE)':
                        tax_value = '0'
                        tax_field.select_by_value(tax_value)
                    else:
                        tax_field.select_by_value(tax_sheet)

                    sleep(0.5)

                    add_product_button = driver.find_element(By.XPATH, "//a[@onclick='addRow(arr)']")
                    add_product_button.click()

                    sleep(0.5)
                    i += 1
                else:
                    delete_product_button = driver.find_element(By.XPATH, f"//a[@onclick='deleteRow({i},0,arr);']")
                    delete_product_button.click()
                    break


            # just click for filling field 'Wartość netto'
            try:
                just_click = driver.find_element(By.XPATH, "//input[@name='wartosc_netto[0]']")
                just_click.click()
            except Exception as ex:
                print(ex)

            if waiting_for_checking() == 'stop':
                driver.close()
                driver.quit()
                return print('Принудительное завершение программы.')

            # clicking on button for create document while it will be True
            err = False
            while not err:
                try:
                    create_factura = driver.find_element(By.XPATH, "//button[@id='pobierz_i_zapisz']")
                    driver.execute_script("arguments[0].click();", create_factura)
                    sleep(10)
                except Exception:
                    err = True

            sleep(2)

            # go back
            driver.back()

        sleep(1)

    driver.close()
    driver.quit()
    return print("\nProgram execution completed successfully")


def main():
    get_factura()


if __name__ == "__main__":
    main()
