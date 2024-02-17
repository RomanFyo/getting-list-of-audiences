import xml.etree.ElementTree as ET
import re

import pandas as pd


class Audiences:
    def __init__(self, path_to_file, to_keep):
        # Приватные поля
        self.__iter__ = ET.parse(path_to_file).getroot().iter()  # итератор
        self.__xmlns__ = "{urn:schemas-microsoft-com:office:spreadsheet}"  # пространство имен (приписывается к ячейкам)
        self.__element__ = None  # текущий элемент
        self.__to_keep__ = to_keep  # список столбцов, которые надо оставить в датафрейме

        # Публичные поля
        self.df = None

        self.get_audiences()  # заполнение датафрейма с аудиториями
        self.format_df()  # удаление из датафрейма ненужных столбцов

    def get_audiences(self):
        self.__element__ = self.get_next()

        # Пропускаем все ячейки до первой ячейки с тегом "ряд"
        while self.__element__.tag != f"{self.__xmlns__}Row":
            self.__element__ = self.get_next()

        # Непосредственно заполнение датафрейма
        while self.__element__ is not None:
            current_row = self.get_row()
            if self.df is None:  # если датафрейм None, то это первая строка файла, в
                self.df = pd.DataFrame(columns=current_row)  # которой записаны столбцы -> можно определить df
            else:
                self.df.loc[len(self.df)] = current_row  # запись строки файла в датафрейм

    def get_row(self):
        current_row = []

        while True:
            self.__element__ = next(self.__iter__, None)
            if self.__element__ is None or self.__element__.tag == f"{self.__xmlns__}Row":
                break  # выход, если новый ряд или итератор закончился
            elif self.__element__.tag == f"{self.__xmlns__}Data":  # ячейки с тегом Data хранят всю отображаемую
                current_row.append(self.__element__.text)  # в таблице файла информацию
            elif self.__element__.tag == f"{self.__xmlns__}strong":  # Специальное условие для 2 ячеек (подбором)
                del current_row[-1]  # Удаляется запись текста из последней встреченной Data-ячейки
                current_row.append(self.__element__.text)

        return current_row

    def format_df(self):
        all_columns = self.df.columns  # получение списка всех столбцов датафрейма

        # Удаление столбцов датафрейма
        self.df = self.df.drop(
            columns=[x for x in all_columns if x not in self.__to_keep__]
        )  # генерация столбцов для удаления
        self.df = self.df[self.__to_keep__]  # расстановка столбцов в нужном порядке
        if "ПК/ММ" in self.__to_keep__:  # разделение строки, перечисляющей характеристики аудитории
            self.df["ПК/ММ"] = self.df["ПК/ММ"].apply(func=separate_characteristics)

    def get_next(self):
        return next(self.__iter__, None)


def separate_characteristics(string):
    if string is None:
        return None
    else:
        return re.split(", |/", string)


if __name__ == "__main__":
    # Название файла
    file_name = "Помещения.xml"
    # Столбцы, которые НЕ нужно удалять из файла
    columns = ["Здание", "Название", "ПК/ММ", "Компьютерная аудитория", "Кол-во ПК"]
    audiences = Audiences(file_name, columns)
    print(audiences.df.to_string())
