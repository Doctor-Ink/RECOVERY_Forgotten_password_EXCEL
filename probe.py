import itertools
import time
import win32com.client as client
from string import digits, punctuation, ascii_letters



PATH = r'C:\Users\Zver\PycharmProjects\RECOVERY_Forgotten password_EXCEL\book.xlsx'

dict_value = {
    '1': digits,
    '2': ascii_letters,
    '3': digits + ascii_letters,
    '4': digits + ascii_letters + punctuation,
}

possible_symbols = dict_value['4']
def enumeration_all_variant(possible_symbols):
    list_password = []
    for pass_length in range(5, 6):
        for password in itertools.product(possible_symbols, repeat=pass_length):
            password = "".join(password)
            list_password.append(password)
    print(len(list_password))


enumeration_all_variant(possible_symbols=possible_symbols)