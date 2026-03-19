import os
import re
from openpyxl import Workbook


def extract_first_number(filename: str):
    """
    Извлекает из строки filename первое трёх- или четырёхзначное число,
    исключая числа в диапазоне 2020–2025.
    Возвращает число в виде строки или None, если ничего не найдено.
    """
    # Ищем все последовательности из 3 или 4 цифр
    candidates = re.findall(r'\d{3,4}', filename)
    for cand in candidates:
        num = int(cand)
        # Проверяем, что число трёх- или четырёхзначное
        if (100 <= num <= 999) or (1000 <= num <= 9999):
            # Исключаем диапазон 2020–2025
            if 2020 <= num <= 2025:
                continue
            return str(num)  # Возвращаем первое подходящее
    return None


def main():
    folder = ("путь к папке").strip()  # УКАЗАТЬ ПУТЬ

    if not os.path.isdir(folder):
        print("Ошибка: указанная папка не существует.")
        return

    files = []
    for entry in os.listdir(folder):
        full_path = os.path.join(folder, entry)
        if os.path.isfile(full_path):
            files.append(entry)

    if not files:
        print("В папке нет файлов.")
        return

    # Подготовка данных: каждая строка = [имя_файла, число_или_пусто]
    rows = []
    for file in files:
        name_without_ext = os.path.splitext(file)[0]
        number = extract_first_number(name_without_ext)
        rows.append([name_without_ext, number if number is not None else ""])

    # Создаём книгу Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Имена файлов"
    ws.append(["Имя файла", "ID"])

    for row in rows:
        ws.append(row)

    # Сохраняем в той же папке
    output_path = os.path.join(folder, "_свод.xlsx")
    try:
        wb.save(output_path)
        print(f"Файл успешно создан: {output_path}")
    except Exception as e:
        print(f"Ошибка при сохранении файла: {e}")


if __name__ == "__main__":
    main()
