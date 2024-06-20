import itertools
import time
from datetime import datetime
import win32com.client as win32
from string import digits, punctuation, ascii_letters
from threading import Thread
import pythoncom
import datetime
from queue import Queue
from password_picker import time_track, input_initial_data, time_running_script


PATH = r"C:\Users\Professional\Desktop\pythonProjects\RECOVERY_Forgotten_password_EXCEL\Book1.xlsx"


# разделим на 4 потока
# В 4-ёxпоточном режиме
#                       [INFO] ---------- see log

class Picker(Thread):
    def __init__(self, path, queue, need_stop, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.path = path
        self.queue = queue
        self.need_stop = need_stop


    def run(self):
        while not self.need_stop[0] or not self.queue.empty():
            password_list = self.queue.get()
            for password in password_list:
                if self.need_stop[0]:
                    break
                self.password_entry(password)
            self.queue.task_done()
        if not self.need_stop[0]:
            print('Не удалось найти пароль, возможно вы ввели неверные данные!!!')


    def password_entry(self, password):
        # функция запускает клиент и проверяет пароль
        try:
            # Сразу перед инициализацией DCOM в run()
            pythoncom.CoInitializeEx(0)
            # brute excel doc
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = 0
            excel.Workbooks.Open(self.path, False, True, None, password)
            time.sleep(0.1)
            print(f"[INFO] ---------- Password is: {password}")
            with open('password.txt', mode='w', encoding='utf-8') as file:
                file.write(password)
            self.need_stop[0] = True
        except Exception as exc:
            # print(exc)
            print(f"Incorrect {password}")


def generator_passw(password_length, possible_symbols):
    count = 0
    result = []
    for pass_length in range(password_length[0], password_length[1] + 1):
        for password in itertools.product(possible_symbols, repeat=pass_length):
            password = "".join(password)
            result.append(password)
            count += 1
            if count % 1000 == 0:
                yield result
                result = []
    if result:
        yield result

@time_track
def main():
    # шаг 1 запрос исходных данных
    list_length_password, possible_symbols = input_initial_data()
    need_stop = [False]
    queue = Queue()
    time_running_script(
        min_characters=list_length_password[0],
        max_characters=list_length_password[-1],
        possible_symbols=possible_symbols
    )

    for lst_psw in generator_passw(password_length=list_length_password, possible_symbols=possible_symbols):
        queue.put(lst_psw)

    threads = []
    for _ in range(4):
        thread = Picker(path=PATH, queue=queue, need_stop=need_stop)
        threads.append(thread)
        thread.start()

    queue.join()
    for thread in threads:
        thread.join()

if __name__ == '__main__':
    main()