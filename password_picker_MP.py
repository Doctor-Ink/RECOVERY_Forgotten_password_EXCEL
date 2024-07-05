import time
import win32com.client as win32
from multiprocessing import Process, Manager, Queue, Event
import pythoncom
from password_picker import time_track, input_initial_data, time_running_script
from password_picker_MT import generator_passw

PATH = r'C:\Users\Professional\Desktop\pythonProjects\RECOVERY_Forgotten_password_EXCEL\book0007.xlsx'

# разделим на 4 процесса
# В 4-ёxппроцессорном режиме
#                       [INFO] ---------- see log


class Picker(Process):
    def __init__(self, path, queue, stop_event, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.path = path
        self.queue = queue
        self.stop_event = stop_event

    def run(self):
        # Сразу перед инициализацией DCOM в run()
        pythoncom.CoInitializeEx(0)
        while not self.stop_event.is_set() or not self.queue.empty():
            try:
                password_list = self.queue.get()
            except Exception:
                continue
            if password_list is None:
                break

            for password in password_list:
                if self.stop_event.is_set():
                    break
                self.password_entry(password)
            self.queue.task_done()
        if not self.stop_event.is_set():
            print('Не удалось найти пароль, возможно вы ввели неверные данные!!!')

    def password_entry(self, password):
        # функция запускает клиент и проверяет пароль
        try:
            # brute excel doc
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = 0
            excel.Workbooks.Open(self.path, False, True, None, password)
            time.sleep(0.1)
            print(f"[INFO] ---------- Password is: {password}")
            with open('password.txt', mode='w', encoding='utf-8') as file:
                file.write(password)
            self.stop_event.set()
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
    stop_event = Event()
    queue = manager.Queue(maxsize=4)

    # Create and start processes
    processes = []
    for _ in range(4):
        process = Picker(path=PATH, queue=queue, stop_event=stop_event)
        processes.append(process)
        process.start()

    # Generate password lists and add them to the queue
    for password_list in generator_passw(password_length=list_length_password, possible_symbols=possible_symbols):
        if stop_event.is_set():
            break
        queue.put(password_list)

    for _ in range(4):
        queue.put(None)

    # Wait for all processes to complete
    for process in processes:
        process.join()


if __name__ == '__main__':
    main()