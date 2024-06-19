import os
import time
import datetime
import win32com.client as client

PATH = r'C:\Users\Zver\PycharmProjects\RECOVERY_Forgotten_password_EXCEL\Таблицы по загазованности .xlsx'
print("***Hello friend!***")

###  в данном скрипте будем перебирать варианты паролей из имеющихся текстовых файлов около 10 миллионов вариантов

def password_entry(password, count):
    # функция запускает клиент и проверяет пароль

    open_doc = client.Dispatch("Excel.Application")
    try:
        open_doc.Workbooks.Open(
            PATH,
            False,
            True,
            None,
            password
        )
        print(f"[INFO] ---------- Password is: {password}")
        with open('password.txt', mode='w', encoding='utf-8') as file:
            file.write(password)
        return False
    except:
        print(f"Attempt {count} Incorrect {password}")
    return True


def open_file_get_list(path):
    my_list = []
    try:
        with open(path, encoding='cp1251') as file:
            for line in file.readlines():
                my_list.append(line[:-1:])
        return my_list
    except Exception as exc:
        print(f'Ошибка кодировки - {path}')
        with open(path, encoding='utf-8') as file:
            for line in file.readlines():
                my_list.append(line[:-1:])
        return my_list


def time_running_script():
    total_count = 10_000_000 + 150 + 1000
    try:
        time_format = str(datetime.timedelta(seconds= (total_count / 6)))
        print(f"Общее число комбинаций - {total_count}\n "
              f"Расчётное время работы - {time_format} секунд")
    except Exception as exc:
        print('Python не переведёт это число в дни и годы')


def time_track(func):
    # функция-декаратор, которая считает время работы
    print('Function is starting!')
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

    time_running_script()
    count = 0
    result = True
    while result:
        for item in os.listdir('10M'):
            print(f'Current file is ---{item}')
            current_list = open_file_get_list(path=os.path.join('10M/', item))
            for password in current_list:
                count +=1
                result = password_entry(password=password, count=count)
                if result is False:
                    break
            if result is False:
                break
        break


if __name__ == '__main__':
    main()