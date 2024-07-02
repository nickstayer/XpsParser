from XpsParser import XpsParser
from Word import Word
import os
from pathlib import Path


def main():
    path_xps = os.path.join(os.getcwd(), "data")
    template = os.path.join(os.getcwd(), "template.docx")
    input(f"Поместите xps файлы в папку {path_xps} и нажмите Enter")
    if os.path.exists(path_xps):
        files = list(Path(path_xps).glob("*"))
        word = Word()
        counter = 0
        for file in files:
            if file.suffix == ".xps":
                counter += 1
                parser = XpsParser(file)
                xps = parser.parse()
                if os.path.exists(template):
                    word.open(template)
                    word.insert_text_after_line("Подразделение:", xps.name)
                    word.insert_text_in_table(1, 1, 2, xps.name)
                    word.insert_text_in_table(1, 2, 2, xps.passw)
                    word.insert_text_in_table(1, 3, 2, xps.passw_phrase)
                    new_file = os.path.join(os.path.dirname(file), xps.name)
                    word.save_as(new_file)
                    word.close()
                else:
                    print(f"Файл {template} отсутствует")
        word.quit()
        input(f"Работа программы завершена! Обработано файлов: {counter}")
    else:
        print(f"Папка {path_xps} не существует")


if __name__ == "__main__":
    main()
