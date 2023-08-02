# Класс создания объекта для форматирования и краткосрочного хранения информации

class Person:

    def __init__(self, name, birth_date, gender, lastname="", surname="", death_date=""):
        self.fullname = f"{lastname.title()} {name.title()} {surname.title()}".strip()
        self.gender = gender.lower()
        self.birth_date = birth_date.replace(" ", ".").replace("/", ".")
        self.death_date = death_date.replace(" ", ".").replace("/", ".")
