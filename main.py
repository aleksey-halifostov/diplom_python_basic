import openpyxl
import secondary_functions as function


# Главная исполняемая функция
# выполняет роль главного меню
# запрашивает у пользователя действие и вызывает соответствующие функции
def main():
    func_to_call = {1: new_person, 2: find_in_data}
    while True:
        print("Выберите действие (введите номер)")
        print("1. Ввести новую запись")
        print("2. Поиск в БД")
        print("3. Загрузить данные в БД из файла")
        print("4. Сохранить данные из БД в файл")
        print("5. Экспортировать данные в json формат")
        print('-' * 50)
        print("0. Выход")
        choice = input("Ваш выбор: ")

        if not choice.isdigit() or int(choice) > 5:
            print("Не корректный ввод")
            continue
        elif int(choice) == 0:
            break

        func_to_call[int(choice)]()


def open_and_closing_data(func):
    def dec_func():
        book = openpyxl.load_workbook("data.xlsx")
        sheet = book["Sheet"]

        try:
            func(sheet)
        finally:
            book.save("data.xlsx")
            book.close()

    return dec_func


# Записывает информацию о новом человеке в базу данных
@open_and_closing_data
def new_person(sheet):
    # Создание объекта класса Person, подготовка к записи
    new_obj = function.get_data_from_operator()
    row = function.find_empty_row(sheet)

    # Запись новых данных, сохранение и закрытие файла
    for ind, elem in enumerate(new_obj.__dict__.values()):
        sheet.cell(row, ind + 1).value = elem


# Выполняет поиск в базе данных по введенному оператором запросу
# печатает все результаты, которые удовлетворяют поиск
@open_and_closing_data
def find_in_data(sheet):
    to_find = input("Введите ФИО: ").upper()
    print("-" * 50)
    male_or_female = {"мужчина": ("Родился:", "Умер:"), "женщина": ("Родилась:", "Умерла:")}

    for row in range(1, sheet.max_row + 1):
        if to_find in sheet.cell(row, 1).value.upper():
            death_str = ""
            if sheet.cell(row, 4).value:
                death_str = f" {sheet.cell(row, 2).value[1]} {sheet.cell(row, 4).value}"
            print(
                f"{sheet.cell(row, 1).value} {function.def_age(sheet.cell(row, 3).value, sheet.cell(row, 4).value)}, " +
                f"{sheet.cell(row, 2).value}. {male_or_female[sheet.cell(row, 2).value][0]} {sheet.cell(row, 3).value}"
                + death_str)


if __name__ == "__main__":
    main()