import itertools
import time
from datetime import datetime
import win32com.client as client
from string import digits, punctuation, ascii_letters
from threading import Thread
import pythoncom



class Picker(Thread):
    def __init__(self, list, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.list = list
        self.password_length = None
        self.possible_symbols = None

    def run(self):
        # Сразу перед инициализацией DCOM в run()
        pythoncom.CoInitializeEx(0)
        # brute excel doc
        count = 0
        for password in self.list:
            open_doc = client.Dispatch("Excel.Application")
            count +=1
            try:
                open_doc.Workbooks.Open(
                    r'C:\Users\Zver\PycharmProjects\RECOVERY_Forgotten_password_EXCEL\НАТА.xlsx',
                    False,
                    True,
                    None,
                    password
                )
                time.sleep(0.1)
                print(f"Attempt #{count} Password is: {password}")
                return True
            except:
                print(f"Attempt #{count} Incorrect {password}")
        return False


def input_initial_data():
    print("***Hello friend!***")
    try:
        password_length = input("Введите длину пароля, от скольки - до скольки символов, например 3 - 7: ")
        password_length = [int(item) for item in password_length.split("-")]
    except Exception:
        print('Проверьте введённые данные')

    print("Если пароль содержит только цифры, введите: 1\nЕсли пароль содержит только буквы, введите: 2\n"
          "Если пароль содержит цифры и буквы введите: 3\nЕсли пароль содержит цифры, буквы и спецсимволы введите: 4")

    try:
        choice = int(input(": "))

        if choice == 1:
            possible_symbols = digits
        elif choice == 2:
            possible_symbols = ascii_letters
        elif choice == 3:
            possible_symbols = digits + ascii_letters
        elif choice == 4:
            possible_symbols = digits + ascii_letters + punctuation
        else:
            print('Введите корректные данные!!!')
        print(possible_symbols)
    except:
        print('Введите корректные данные....')
    return possible_symbols, password_length


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


# def get_all_variant(password_length):
#     all_variant_list = []
#     for pass_length in range(password_length[0], password_length[1] + 1):
#         for password in itertools.product(possible_symbols, repeat=pass_length):
#             password = "".join(password)
#             recovery_excel_password()
#     print(all_variant_list)
#     return all_variant_list


def main():
    # шаг 1 запрос исходных данных
    input_initial_data()
    start_timestamp = time.time()
    print(f"Started at - {datetime.utcfromtimestamp(time.time()).strftime('%H:%M:%S')}")

    # шаг 2 - в многопоточном стиле перебираем два списка
    my_list_150 = get_list_150()
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