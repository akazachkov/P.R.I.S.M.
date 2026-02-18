# app/core/elements/working_with_folders.py

import subprocess
import re
from pathlib import Path


def ensure_directory_exists(base_path, subfolder):
    """
    Создает директорию, если она не существует

    Args:
        - base_path - Путь к базовой директории
        - subfolder - Имя новой поддиректории
    """
    directory = Path(f"{base_path}/{subfolder}")
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def open_file_and_folder(path: str | Path) -> bool:
    """
    Открывает указанную директорию или файл в Проводнике Windows

    Args:
        - path - Строка или Path-объект, указывающий на директорию

    Returns:
        - bool - True, если команда была успешно запущена, иначе False
    """
    try:
        # Нормализуем путь, преобразуя его в Windows-совместимый вид
        target_path = Path(path).resolve()

        # Проверяем, существует ли путь
        if not target_path.exists():
            print(f"Ошибка: путь '{target_path}' не найден.")
            return False

        # Самая простая и надежная команда
        subprocess.run(["explorer", str(target_path)])
        return True

    except subprocess.CalledProcessError as e:
        print(f"Не удалось запустить 'explorer'. Ошибка: {e}")
        return False
    except Exception as e:
        print(f"Ошибка: {e}")
        return False


def parse_file_path(file_path):
    """
    Парсит путь к файлу для извлечения базовой директории, числа (номера) и
    аббревиатуры
    """
    try:
        path_obj = Path(file_path)

        # 1. Извлекаем базовую директорию
        base_dir = str(path_obj.parent)

        # 2. Извлекаем имя файла без пути
        filename = path_obj.name

        # 3. Ищем ПЕРВОЕ число из 1-4 цифр и дополняем ведущими нулями
        number_match = re.search(r'(?:_|\b)(\d{1,4})(?:_|\b)', filename)
        if not number_match:
            # Попробуем другой вариант - ищем просто последовательность цифр
            number_match = re.search(r'(\d{1,4})', filename)
        number = number_match.group(1).zfill(4) if number_match else None

        # 4. Ищем аббревиатуру ВСО или ВДО
        abbrev_match = re.search(r'(?:^|_|\s)(ВСО|ВДО)(?:$|_|\s)', filename)
        if not abbrev_match:
            # Если не нашли с разделителями, ищем просто в строке
            abbrev_match = re.search(r'(ВСО|ВДО)', filename)
        abbrev = abbrev_match.group(1) if abbrev_match else None
        return base_dir, number, abbrev
    except Exception as e:
        print(f"Ошибка при парсинге пути: {e}")
        return None, None, None
