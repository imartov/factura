from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.support.ui import Select
import os
from openpyxl import load_workbook
from fake_useragent import UserAgent
import keys
from login_facturowo import input_login, get_saved_login_data, get_temp_login_data
from tqdm import tqdm, tqdm_gui, tnrange


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
    os.remove(f"{os.getcwd()}\\temp_login.json")

    # create new document
    new_document_button = driver.find_element(By.XPATH, f"//a[@class='btn btn-xs btn-primary']")
    new_document_button.click()

    sleep(1.5)

    # getting input excel files
    path = "D:\\python_projects\\factura\\xlsx"
    files_list = os.listdir(path)
    print(f"\nВсего найдено файлов: {len(files_list)}.")

    for file in files_list:
        print(f"Извлечен файл {file}.\n")

        # getting sheets in every excel file
        wb = load_workbook(filename=f"{path}\\{file}")
        print(f"Всего найдено листов: {len(wb.sheetnames)}.")

        for sheet in wb.sheetnames:

            # filling fields about companies
            supply_field = driver.find_element(By.NAME, "sprzedawca[nazwa]")
            supply_field.send_keys(Keys.CONTROL, "a")
            supply_field.send_keys(Keys.DELETE)
            supply_field.send_keys("Supply_company")

            sleep(0.5)

            supply_nip_field = driver.find_element(By.NAME, "sprzedawca[nip]")
            supply_nip_field.send_keys(Keys.CONTROL, "a")
            supply_nip_field.send_keys(Keys.DELETE)
            supply_nip_field.send_keys("9856325783")

            sleep(0.5)

            buyer_field = driver.find_element(By.NAME, "nabywca[nazwa]")
            buyer_field.send_keys(Keys.CONTROL, "a")
            buyer_field.send_keys(Keys.DELETE)
            buyer_field.send_keys("Buyer_company")

            sleep(0.5)

            buyer_nip_field = driver.find_element(By.NAME, "nabywca[nip]")
            buyer_nip_field.send_keys(Keys.CONTROL, "a")
            buyer_nip_field.send_keys(Keys.DELETE)
            buyer_nip_field.send_keys("9687653259")

            sleep(0.5)

            # select active sheet
            sheet_active = wb[sheet]
            print(f"Извлечение данных из листа: {sheet}.")
            print(f"Всего найдено позиций: {sheet_active.max_column - 1}.")

            for row in range(1, sheet_active.max_column, 1):

                # select data about products
                product_name_sheet = sheet_active[1][row].value
                count_sheet = sheet_active[3][row].value
                price_netto_sheet = sheet_active[4][row].value

                sleep(0.5)

                product_name_site = driver.find_element(By.ID, f"nazwa_{row-1}")
                product_name_site.clear()
                product_name_site.send_keys(product_name_sheet)

                sleep(0.5)

                select_site = Select(driver.find_element(By.ID, f"jm_{row-1}"))
                select_site.select_by_index(row)

                sleep(0.5)

                count_site = driver.find_element(By.ID, f"ilosc_{row-1}")
                count_site.send_keys(Keys.CONTROL, "a")
                count_site.send_keys(Keys.DELETE)
                count_site.send_keys(count_sheet)

                sleep(0.5)

                price_netto_site = driver.find_element(By.XPATH, f"//input[@id='cena_netto_{row-1}']")
                price_netto_site.send_keys(Keys.CONTROL, "a")
                price_netto_site.send_keys(Keys.DELETE)
                price_netto_site.send_keys(price_netto_sheet)

                sleep(0.5)

                if row < sheet_active.max_column - 1:
                    add_button_site = driver.find_element(By.XPATH, "//a[@onclick='addRow(arr)']")
                    add_button_site.click()

                    sleep(0.5)
                else:
                    pass

            try:
                just_click = driver.find_element(By.XPATH, "//input[@name='wartosc_netto[0]']")
                just_click.click()
            except Exception as ex:
                print(ex)

            # clicking on button for create document while it will be True
            err = False
            while not err:
                try:
                    create_factura = driver.find_element(By.XPATH, "//button[@id='pokaz_i_zapisz']")
                    driver.execute_script("arguments[0].click();", create_factura)
                    sleep(1)
                except Exception:
                    err = True

            sleep(2)

            # go back
            driver.back()

        sleep(1)

    driver.close()
    driver.quit()
    return print("Program execution completed successfully")


def main():
    get_factura()


if __name__ == "__main__":
    main()
