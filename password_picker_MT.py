import itertools
import time
from datetime import datetime
import win32com.client as client
from string import digits, punctuation, ascii_letters
from threading import Thread
import pythoncom


PATH = r'C:\Users\Zver\PycharmProjects\RECOVERY_Forgotten_password_EXCEL\book.xlsx'

# разделим на 4 потока, пароли будем забирать из генератора

def combination_generator_first(password_length, possible_symbols):
    # в этом генероторе проходим все конбинации кроме последней
    for pass_length in range(password_length[0], password_length[1]):
            for password in itertools.product(possible_symbols, repeat=pass_length):
                password = "".join(password)
                yield password

def combination_generator_second(password_length, possible_symbols):
    # этот генеротор проходит все конбинации последнего числа длины пароля "123456" - все комбинации из 6 символов
    for pass_length in range(password_length[1], password_length[1] + 1):
        for password in itertools.product(possible_symbols, repeat=pass_length):
            password = "".join(password)
            yield password


class Picker(Thread):
    def __init__(self, PATH, password, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.PATH = PATH
        self.password = password

    def run(self):
        # Сразу перед инициализацией DCOM в run()
        pythoncom.CoInitializeEx(0)
        # brute excel doc
        count = 0
        while self.trigger:
            for password in self.list:
                self.password_entry(password=password)

    def password_entry(self):
        # функция запускает клиент и проверяет пароль

        open_doc = client.Dispatch("Excel.Application")
        try:
            open_doc.Workbooks.Open(
                PATH,
                False,
                True,
                None,
                self.password
            )
            time.sleep(0.1)
            print(f"[INFO] ---------- Password is: {self.password}")
            with open('password.txt', mode='w', encoding='utf-8') as file:
                file.write(self.password)
            return False
        except:
            print(f"Incorrect {self.password}")
        return True

def input_initial_data():
    # функция запрашивает исходные данные

    while True:
        password_length = input("Введите длину пароля, от скольки - до скольки символов, например 3 - 7: ")
        if ('-' in password_length) and (password_length.replace('-', '').isdigit()):
            password_length = [int(item) for item in password_length.split('-')]
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
            return password_length, possible_symbols
        else:
            print('Введите корректные данные!!!')


def get_list_150():
    my_list_150 = []
    with open('150_russian_password.txt') as file:
        for line in file.readlines():
            my_list_150.append(line[:-1:])
    return my_list_150


def get_list_10K():
    my_list_10K = []
    with open('10000_world_password.txt') as file:
        file.readline()
        for line in file.readlines():
            my_list_10K.append(line[:-1:])
    return my_list_10K


def main():
    # шаг 1 запрос исходных данных
    input_initial_data()
    start_timestamp = time.time()
    print(f"Started at - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")

    # шаг делим исходный массив данных на 4 потока
    # исходный массив для первого потока:
    my_list_150 = get_list_150()




    # шаг 2 - в многопоточном стиле перебираем два списка

    my_list_10K = get_list_10K()
    print(my_list_150)


    thread_one = Picker(list=my_list_150)
    thread_two = Picker(list=my_list_10K)

    thread_one.start()
    thread_two.start()

    thread_one.join()
    thread_two.join()

    # # шаг 3 - перебор списка из 10 000 популярных паролей в мире
    #
    # recovery_excel_password(list=my_list_10K)
    #
    # # шаг 4 - бональный перебор всех возможных комбинаций
    # all_variant_list = get_all_variant(possible_symbols=possible_symbols, list_len=password_length)
    # recovery_excel_password(list=all_variant_list)

    print(f"Finished at - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")
    print(f"Password cracking time - {time.time() - start_timestamp}")


if __name__ == '__main__':
    main()