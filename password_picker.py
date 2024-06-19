import itertools
import time
import datetime
import win32com.client as win32
from string import digits, punctuation, ascii_letters

PATH = r'C:\Users\Professional\Desktop\pythonProjects\RECOVERY_Forgotten_password_EXCEL\book.xlsx'

print("***Hello friend!***")

# В однопоточном режиме 100 паролей перебираются   за 18 секунд
#                       1000 паролей перебираются  за 180 секунд (3 минуты)
#                       10000 паролей перебираются за 1800 секунд (30 минут)
#                       [INFO] ---------- Password is: 101
#                                               Скрипт отработал - 40.51 секунды


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


def input_initial_data():
    # функция запрашивает исходные данные

    while True:
        password_length = input("Введите длину пароля, от скольки - до скольки символов, например 3 - 7: ")
        if ('-' in password_length) and (password_length.replace('-', '').isdigit()):
            password_length = [int(item) for item in password_length.split('-')]
        else:
            print('некорректные данные')
            continue

        choice = input("Если пароль содержит только цифры, введите: 1\n"
                       "Если пароль содержит только буквы, введите: 2\n"
                       "Если пароль содержит цифры и буквы введите: 3\n"
                       "Если пароль содержит цифры, буквы и спецсимволы введите: 4\n"
                       "------------>   ")

        dict_value = {
            '1': digits,  # 0123456789
            '2': ascii_letters,  # abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ
            '3': ascii_letters + digits,  # abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789
            '4': ascii_letters + digits + punctuation,  # !"#$%&'()*+,-./:;<=>?@[\]^_`{|}~
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
def enumeration_all_variants(password_length, possible_symbols, count):
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
        print(f'Для пароля из {step} символа(ов) - \n{len(possible_symbols) ** step} комбинаций')
        total_count += len(possible_symbols) ** step
    try:
        time_format = str(datetime.timedelta(seconds= (total_count / 6)))
        print(f"Общее число комбинаций - {total_count}\n "
              f"Расчётное время работы - {time_format} секунд")
    except Exception as exc:
        print('Python не переведёт это число в дни и годы')


def main():
    # шаг 1 запрос исходных данных
    pass_length, possible_symbols = input_initial_data()
    time_running_script(min_characters=pass_length[0], max_characters=pass_length[-1], possible_symbols=possible_symbols)
    count = 0
    while True:
        # шаг 2 - перебираем из 150 популярных паролей в России
        # my_list_150 = get_list_150()
        # for item in my_list_150:
        #     count += 1
        #     result = password_entry(password=item, count=count)
        #     if result is False:
        #         break
        # if result is False:
        #     break

        # шаг 3 - перебор списка из 10 000 популярных паролей в мире
        # my_list_10K = get_list_10K()
        # for item in my_list_10K:
        #     count += 1
        #     result = password_entry(password=item, count=count)
        #     if result is False:
        #         break
        # if result is False:
        #     break

        # шаг 4 - перебор списка из 1 000 000 популярных паролей в мире
        # my_list_1M = get_list_1M()
        # for item in my_list_1M:
        #     count += 1
        #     result = password_entry(password=item, count=count)
        #     if result is False:
        #         break
        # if result is False:
        #     break

        # шаг 5 - перебор списка из 10 000 000 популярных паролей в мире
        # my_list_10M = get_list_10M()
        # for item in my_list_10M:
        #     count += 1
        #     result = password_entry(password=item, count=count)
        #     if result is False:
        #         break
        # if result is False:
        #     break

        # шаг 4 - бональный перебор всех возможных комбинаций
        result = enumeration_all_variants(password_length=pass_length, possible_symbols=possible_symbols, count=count)
        if result is False:
            break


if __name__ == '__main__':
    main()