from string import digits, punctuation, ascii_letters

print("***Hello friend!***")


def input_initial_data():
    # функция запрашивает исходные данные

    while True:
        password_length = input("Введите длину пароля, от скольки - до скольки символов, например 3 - 7: ")
        if ('-' in password_length) and (password_length.replace('-', '').isdigit()):
            password_length = [int(item) for item in password_length.split('-')]
        else:
            print('некорректные данные')
            continue

        choice = input("Если пароль содержит только цифры, введите: 1\nЕсли пароль содержит только буквы, введите: 2\n"
              "Если пароль содержит цифры и буквы введите: 3\nЕсли пароль содержит цифры, буквы и спецсимволы введите: 4\n "
                       "------------>   ")

        dict_value = {
            '1': digits,
            '2': ascii_letters,
            '3': digits + ascii_letters,
            '4': digits + ascii_letters + punctuation,
        }

        if choice in dict_value.keys():
            possible_symbols = dict_value[choice]
            return password_length, possible_symbols
        else:
            print('Введите корректные данные!!!')


password_length, possible_symbols = input_initial_data()
print(password_length, possible_symbols)
print(password_length[0])
print(password_length[1])
print(type(password_length[1]))
print(type(password_length[0]))