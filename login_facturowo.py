import json
import os


def input_yes_no():
    input_saved_login = input("Сохранить логин и пароль? Введите 'yes' or 'no': ")
    if input_saved_login != "yes" and input_saved_login != "no":
        print("You need to enter only 'yes' or 'no'")
        input_yes_no()
    return input_saved_login


def input_login():
    input_email = input("\nВведите логин от https://www.fakturowo.pl: ")
    input_password = input("Введите пароль от https://www.fakturowo.pl: ")
    saved_login = input_yes_no()
    if saved_login == "yes":
        saved_data = {
            "email": input_email,
            "password": input_password
        }

        try:
            with open(f"{os.getcwd()}\\saved_login.json", "w", encoding="utf-8") as file:
                json.dump(saved_data, file, indent=4, ensure_ascii=False)
        except Exception:
            pass
    elif saved_login == "no":
        saved_data = {
            "email": input_email,
            "password": input_password
        }

        with open(f"{os.getcwd()}\\temp_login.json", "w", encoding="utf-8") as file:
            json.dump(saved_data, file, indent=4, ensure_ascii=False)


def get_saved_login_data():
    with open(f"{os.getcwd()}\\saved_login.json", encoding="utf-8") as file:
        login_data = json.load(file)
        return login_data["email"], login_data["password"]


def get_temp_login_data():
    with open(f"{os.getcwd()}\\temp_login.json", encoding="utf-8") as file:
        login_data = json.load(file)
        return login_data["email"], login_data["password"]


def main():
    input_login()
    input_yes_no()
    get_saved_login_data()
    get_temp_login_data()


if __name__ == "main__":
    main()
