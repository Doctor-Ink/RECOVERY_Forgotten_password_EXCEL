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

# разделим на 4 потока, пароли будем забирать из генератора
# В 4-ёxпоточном режиме 100 паролей перебираются   за 18 секунд
#                       1000 паролей перебираются  за 180 секунд (3 минуты)
#                       10000 паролей перебираются за 1800 секунд (30 минут)

def combination_generator(password_length, possible_symbols):
    # в этом генероторе проходим все конбинации кроме последней
    for pass_length in range(password_length[0], password_length[1]):
            for password in itertools.product(possible_symbols, repeat=pass_length):
                password = "".join(password)
                print(password)
                yield password


class Picker(Thread):
    def __init__(self, PATH, password, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.PATH = PATH
        self.password = password

    def run(self):
        while True:
            self.password_entry()

    def password_entry(self):
        # функция запускает клиент и проверяет пароль

        # Сразу перед инициализацией DCOM в run()
        pythoncom.CoInitializeEx(0)
        # brute excel doc
        open_doc = client.Dispatch("Excel.Application")
        try:
            open_doc.Workbooks.Open(PATH, False, True, None, self.password)
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
        # print(f'Для пароля из {step} символа(ов) - \n{len(possible_symbols) ** step} комбинаций')
        total_count += len(possible_symbols) ** step
    try:
        time_format = str(datetime.timedelta(seconds= (total_count * 0.18)))
        print(f"Общее число комбинаций - {total_count}\n "
              f"Расчётное время работы - {time_format} секунд")
    except Exception as exc:
        print('Python не переведёт это число в дни и годы')

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

@time_track
def main():
    # шаг 1 запрос исходных данных
    list_length_password, possible_symbols = input_initial_data()
    time_running_script(
        min_characters=list_length_password[0],
        max_characters=list_length_password[-1],
        possible_symbols=possible_symbols
    )

    # шаг делим исходный массив данных на 4 потока
    # исходный массив для первого потока:
    first = Picker(
        PATH=PATH,
        password=combination_generator(password_length=list_length_password[0], possible_symbols=possible_symbols)
    )
    second = Picker(
        PATH=PATH,
        password=combination_generator(password_length=pass_length, possible_symbols=possible_symbols)
    )
    third = Picker(
        PATH=PATH,
        password=combination_generator(password_length=pass_length, possible_symbols=possible_symbols)
    )
    fourth = Picker(
        PATH=PATH,
        password=combination_generator(password_length=pass_length, possible_symbols=possible_symbols)
    )
    first.start()
    second.start()
    third.start()
    fourth.start()

    first.join()
    second.join()
    third.join()
    fourth.join()





    # шаг 2 - в многопоточном стиле перебираем два списка

    # my_list_10K = get_list_10K()
    # print(my_list_150)
    #
    #
    # thread_one = Picker(list=my_list_150)
    # thread_two = Picker(list=my_list_10K)
    #
    # thread_one.start()
    # thread_two.start()
    #
    # thread_one.join()
    # thread_two.join()

    # # шаг 3 - перебор списка из 10 000 популярных паролей в мире
    #
    # recovery_excel_password(list=my_list_10K)
    #
    # # шаг 4 - бональный перебор всех возможных комбинаций
    # all_variant_list = get_all_variant(possible_symbols=possible_symbols, list_len=password_length)
    # recovery_excel_password(list=all_variant_list)


if __name__ == '__main__':
    main()