# Класс создания объекта для форматирования и краткосрочного хранения информации

class Person:

    def __init__(self, name, birth_date, gender, lastname="", surname="", death_date=""):
        self.name = f"{lastname.title()} {name.title()} {surname.title()}".strip()
        self.gender = gender.lower()
        self.birth_date = birth_date.replace(" ", ".").replace("/", ".")

        if death_date:
            self.death_date = death_date.replace(" ", ".").replace("/", ".")
        else:
            self.death_date = death_date