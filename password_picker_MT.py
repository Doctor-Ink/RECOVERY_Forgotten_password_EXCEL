import itertools
import time
from datetime import datetime
import win32com.client as client
from string import digits, punctuation, ascii_letters
from threading import Thread
import pythoncom
import datetime
from password_picker import time_track


PATH = r'C:\Users\Zver\PycharmProjects\RECOVERY_Forgotten_password_EXCEL\book.xlsx'

# разделим на 3 потока
# В 3-ёxпоточном режиме
#                       [INFO] ---------- Password is: 101
#                               Скрипт отработал - 195.68 секунды

class Picker(Thread):
    def __init__(self, PATH, list_length_password, need_stop, possible_symbols, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.PATH = PATH
        self.list_length_password = list_length_password
        self.possible_symbols = possible_symbols
        self.need_stop = need_stop


    def run(self):
        for pass_length in range(self.list_length_password[0], self.list_length_password[-1] + 1):
            for password in itertools.product(self.possible_symbols, repeat=pass_length):
                password = "".join(password)
                if self.need_stop:
                    break
                self.password_entry(password=password)

        print('Не удалось найти пароль, возможно вы ввели неверные данные!!!')

    def password_entry(self, password):
        # функция запускает клиент и проверяет пароль
        try:
            # Сразу перед инициализацией DCOM в run()
            pythoncom.CoInitializeEx(0)
            # brute excel doc
            open_doc = client.Dispatch("Excel.Application")
            open_doc.Workbooks.Open(PATH, False, True, None, password)
            time.sleep(0.1)
            print(f"[INFO] ---------- Password is: {password}")
            with open('password.txt', mode='w', encoding='utf-8') as file:
                file.write(password)
            self.need_stop = True
        except:
            print(f"Incorrect {password}")

def input_initial_data():
    # функция запрашивает исходные данные

    while True:
        password_length = input("Введите длину пароля, от скольки - до скольки символов, например 3 - 7: ")
        if ('-' in password_length) and (password_length.replace('-', '').isdigit()):
            password_length = [int(item) for item in password_length.split('-')]
            list_length_password = [item for item in range(password_length[0], password_length[1] + 1)]
        else:
            print('некорректные данные')
            continue

        choice = input("Если пароль содержит только цифры, введите: 1\nЕсли пароль содержит только буквы, введите: 2\n"
                       "Если пароль содержит цифры и буквы введите: 3\n"
                       "Если пароль содержит цифры, буквы и спецсимволы введите: 4\n------------>   ")

        dict_value = {
            '1': digits,  # 0123456789
            '2': ascii_letters,  # abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ
            '3': digits + ascii_letters,  # 0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ
            '4': digits + ascii_letters + punctuation,  # !"#$%&'()*+,-./:;<=>?@[\]^_`{|}~
        }

        if choice in dict_value.keys():
            possible_symbols = dict_value[choice]
            return list_length_password, possible_symbols
        else:
            print('Введите корректные данные!!!')


def time_running_script(min_characters, max_characters, possible_symbols):
    total_count = 0
    for step in range(min_characters, max_characters + 1):
        print(f'Для пароля из {step} символа(ов) - \n{len(possible_symbols) ** step} комбинаций')
        total_count += len(possible_symbols) ** step
    try:
        time_format = str(datetime.timedelta(seconds= (total_count * 0.18)))
        print(f"Общее число комбинаций - {total_count}\n "
              f"Расчётное время работы - {time_format} секунд")
    except Exception as exc:
        print('Python не переведёт это число в дни и годы')


@time_track
def main():
    # шаг 1 запрос исходных данных
    list_length_password, possible_symbols = input_initial_data()
    time_running_script(
        min_characters=list_length_password[0],
        max_characters=list_length_password[-1],
        possible_symbols=possible_symbols
    )

    first = Picker(PATH=PATH, list_length_password=list_length_password[:-2], possible_symbols=possible_symbols, need_stop=False)
    second = Picker(PATH=PATH, list_length_password=list_length_password[-2:-1], possible_symbols=possible_symbols, need_stop=False)
    third = Picker(PATH=PATH, list_length_password=[list_length_password[-1]], possible_symbols=possible_symbols, need_stop=False)
    first.start()
    second.start()
    third.start()
    while True:
        if first.need_stop or second.need_stop or third.need_stop:
            first.need_stop = True
            second.need_stop = True
            third.need_stop = True
            break
    first.join()
    second.join()
    third.join()


if __name__ == '__main__':
    main()