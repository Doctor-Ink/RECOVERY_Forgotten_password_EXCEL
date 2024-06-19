import itertools
import time
import datetime
import win32com.client as win32
from string import digits, punctuation, ascii_letters

PATH = r'C:\Users\Professional\Desktop\pythonProjects\RECOVERY_Forgotten_password_EXCEL\book.xlsx'
print("***Hello friend!***")


def input_initial_data():
    # функция запрашивает исходные данные

    while True:
        password_length = input("Введите длину пароля, от скольки - до скольки символов, например 3-7: ")
        if ('-' in password_length) and (password_length.replace('-', '').isdigit()):
            password_length = [int(item) for item in password_length.split('-')]
        else:
            print('некорректные данные')
            continue

        choice = input("Если пароль содержит только цифры, введите: 1\n"
                       "Если пароль содержит только буквы, введите: 2\n"
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


def password_entry(path, password, count):
    #
    # Open an existing workbook
    #
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = 0
    try:
        excel.Workbooks.Open(
            Filename=path,
            UpdateLinks=False,
            ReadOnly=True,
            Format=None,
            Password=password
        )
        print(f"[INFO] ---------- Password is: {password}")
        with open('password.txt', mode='w', encoding='utf-8') as file:
            file.write(password)
        return False
    except Exception as exc:
        # print(exc)
        time.sleep(0.2)
        print(f"Attempt {count} Incorrect {password}")
    return True


def enumeration_all_variant(password_length, possible_symbols, count):
    for pass_length in range(password_length[0], password_length[1] + 1):
        for password in itertools.product(possible_symbols, repeat=pass_length):
            password = "".join(password)
            count += 1
            result = password_entry(path=PATH, password=password, count=count)
            if result is False:
                return False
    print('Не удалось найти пароль, возможно вы ввели неверные данные!!!')
    return False


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


def time_track(func):
    # функция-декаратор, которая считает время работы
    def surogate(*args, **kwargs):
        start_time = time.time()

        result = func(*args, **kwargs)

        end_time= time.time()
        result_time = end_time - start_time
        print(f'Скрипт отработал - {round(result_time, 2)} секунды')
        return result
    return surogate

@time_track
def main():
    # шаг 1 запрос исходных данных
    pass_length, possible_symbols = input_initial_data()
    time_running_script(
        min_characters=pass_length[0],
        max_characters=pass_length[1],
        possible_symbols=possible_symbols
    )

    count = 0
    while True:
        result = enumeration_all_variant(
            password_length=pass_length,
            possible_symbols=possible_symbols,
            count=count
        )
        if result is False:
            break


if __name__ == '__main__':
    main()