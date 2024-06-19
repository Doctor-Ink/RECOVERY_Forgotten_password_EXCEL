import win32com.client as win32
import time


PATH = r'C:\Users\Professional\Desktop\pythonProjects\RECOVERY_Forgotten_password_EXCEL\book.xlsx'

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

for pasw in ['000', '001', '002', '003', '004', '005', '006', '007', '008', ]:
    password_entry(path=PATH, password=pasw, count=1)