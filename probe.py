import itertools
import time
import win32com.client as client
from string import digits, punctuation, ascii_letters



# PATH = r'C:\Users\Zver\PycharmProjects\RECOVERY_Forgotten password_EXCEL\book.xlsx'
#
# dict_value = {
#     '1': digits,    #   0123456789
#     '2': ascii_letters,    #    abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ
#     '3': digits + ascii_letters,      #     0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ
#     '4': digits + ascii_letters + punctuation,      #   !"#$%&'()*+,-./:;<=>?@[\]^_`{|}~
# }
#
# # possible_symbols = dict_value['4']
# possible_symbols = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-_'
# begin = 0
# end = 15
# print(len(possible_symbols))
# print(dict_value['1'])
# print(dict_value['2'])
# print(dict_value['3'])
# print(dict_value['4'])
# # количество комбинаций
# total_count = 0
# for step in range(begin, end + 1):
#     print(f'Для пароля из {step} символа - \n{len(possible_symbols) ** step} комбинаций')
#     total_count += len(possible_symbols) ** step
#     print(f'Общее число комбинаций - \n{total_count}')
# print(f'Делим на 4 \n{total_count//4}')

some_time = 365 * 24 * 3600 + 3600 * 2 + 60 * 15 + 45  #  365 * 24 * 3600 = 31536000
print(some_time)
days = some_time // (24 * 3600)
print(days)

hour = some_time % 3600
print(hour)
# sec %= 3600 min = sec // 60 sec %= 60


