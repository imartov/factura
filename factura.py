from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.support.ui import Select
from random import sample


def get_factura():
    product_list = ["Товар_1", "Товар_2", "Товар_3", "Товар_4", "Товар_5"]
    list_count = sample(range(1, 101, 1), len(product_list))
    price_netto_list = sample(range(100, 1001, 1), len(product_list))

    driver = webdriver.Chrome()
    driver.maximize_window()
    url = "https://www.fakturowo.pl/wystaw"
    driver.get(url)

    sleep(3)

    for index, product in enumerate(product_list, start=0):
        product_name = driver.find_element(By.ID, f"nazwa_{index}")
        product_name.clear()
        product_name.send_keys(product)

        select = Select(driver.find_element(By.ID, f"jm_{index}"))
        select.select_by_index(index + 1)

        sleep(1)

        count = driver.find_element(By.ID, f"ilosc_{index}")
        count.clear()
        count.send_keys(list_count[index])

        sleep(1)

        price_netto = driver.find_element(By.XPATH, f"//input[@id='cena_netto_{index}']")
        price_netto.send_keys(Keys.CONTROL, "a")
        price_netto.send_keys(Keys.DELETE)
        price_netto.send_keys(price_netto_list[index])

        sleep(1)

        add_button = driver.find_element(By.XPATH, "//a[@onclick='addRow(arr)']")
        add_button.click()

        sleep(1)

    sleep(10)


def main():
    get_factura()


if __name__ == "__main__":
    main()