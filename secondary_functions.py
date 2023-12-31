# Функции, которые вызываются в основных функциях файла "main.py"

import person
from datetime import datetime


# Функция получает от пользователя информацию, необходимую для новой записи в базу данных
# возвращает объект класса Person
def get_data_from_operator():
    while True:
        lastname = input("Введите фамилию (или нажмите 'Enter' что бы продолжить): ")
        name = input("Введите имя: ")
        surname = input("Введите отчество (или нажмите 'Enter' что бы продолжить): ")
        gender = input("Введите пол (мужчина/женщина): ")
        birth_date = input("Введите дату рождения: ")
        death_date = input("Введите дату смерти (или нажмите 'Enter' что бы продолжить): ")

        if name and gender and birth_date:
            return person.Person(name, birth_date, gender, lastname, surname, death_date)

        print("Вы ввели неверные данные! Имя, пол и дата рождения обязательны для ввода! Попробуйте снова")


# Возвращает номер первой пустой строки в таблице
def find_empty_row(sheet):
    for row in (sheet.max_row, sheet.max_row + 1):
        if not sheet.cell(row, 1).value:
            return row


# Принимает дату рождения и дату смерти (или сегодняшнюю дату) из базы данных
# возвращает возраст
def def_age(birth_date, last_day):
    day, month, year = birth_date.split(".")
    day_last_day, month_last_day, year_last_day = last_day.split(".")

    age = int(year_last_day) - int(year)

    if int(month_last_day) == int(month):
        if int(day) > int(day_last_day):
            age -= 1
    elif int(month_last_day) < int(month):
        age -= 1

    form = " года"
    if str(age)[-2:] in ("12", "13", "14") or not str(age)[-1] in "234":
        form = " лет"

    return str(age) + form


# Функция возвращает дату в нужном для работы программы формате
def date(day):
    if day:
        if isinstance(day, str):
            right_date = day
        else:
            year, month, day, *a = str(day).replace(" ", "-").split("-")
            right_date = f"{day}.{month}.{year}"
    else:
        year, month, day, *a = str(datetime.today()).replace(" ", "-").split("-")
        right_date = f"{day}.{month}.{year}"

    return right_date
