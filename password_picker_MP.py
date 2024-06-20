import itertools
import time
from datetime import datetime
import win32com.client as win32
from multiprocessing import Process, Manager, Queue
import pythoncom
from string import digits, punctuation, ascii_letters
import datetime
from password_picker import time_track, input_initial_data, time_running_script
from password_picker_MT import generator_passw

PATH = r'C:\Users\Professional\Desktop\pythonProjects\RECOVERY_Forgotten_password_EXCEL\Book1.xlsx'

# разделим на 4 процесса
# В 4-ёxппроцессорном режиме
#                       [INFO] ---------- see log


class Picker(Process):
    def __init__(self, path, queue, need_stop, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.path = path
        self.queue = queue
        self.need_stop = need_stop

    def run(self):
        while not self.need_stop.value or not self.queue.empty():
            password_list = self.queue.get()
            for password in password_list:
                if self.need_stop.value:
                    break
                self.password_entry(password)
            self.queue.task_done()
        if not self.need_stop.value:
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
            self.need_stop.value = True
        except Exception as exc:
            # print(exc)
            print(f"Incorrect {password}")


@time_track
def main():
    # шаг 1 запрос исходных данных
    list_length_password, possible_symbols = input_initial_data()
    time_running_script(
        min_characters=list_length_password[0],
        max_characters=list_length_password[-1],
        possible_symbols=possible_symbols
    )

    manager = Manager()
    need_stop = manager.Value('i', False)
    queue = manager.Queue()

    # Generate password lists and add them to the queue
    for password_list in generator_passw(password_length=list_length_password, possible_symbols=possible_symbols):
        queue.put(password_list)

    # Create and start processes
    processes = []
    for _ in range(4):
        process = Picker(path=PATH, queue=queue, need_stop=need_stop)
        processes.append(process)
        process.start()

    # Wait for all processes to complete
    for process in processes:
        process.join()


if __name__ == '__main__':
    main()