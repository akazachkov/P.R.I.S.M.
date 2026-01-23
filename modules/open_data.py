import tkinter as tk
import tomllib
import subprocess
from pathlib import Path

from core.module_loader import BaseModule
from config.app_config import CONFIG_PATHS_NAME


# Функция для открытия файла и директории (универсальная)
def open_file_and_folder(path: str | Path) -> bool:
    """
    Открывает указанную директорию или файл в Проводнике Windows.

    Args:
        path: Строка или Path-объект, указывающий на директорию.

    Returns:
        bool: True, если команда была успешно запущена, иначе
        False.
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


class NameNewModule(BaseModule):
    """
    Открытие папок и файлов.
    """
    name = "open_data"
    button_text = "Панель кнопок для открытия файлов и папок"
    module_label = (
        "Описание функционала модуля - быстрая навигация по папкам и файлам"
    )

    @classmethod
    def initialize_frame(cls, parent_frame: tk.Frame) -> None:
        """
        parent_frame - фрейм для содержимого модуля.
        """
        # Открываем файл и загружаем его содержимое в словарь `config_p_n`
        with open(CONFIG_PATHS_NAME, "rb") as f:  # 'rb' (read binary)
            config_p_n = tomllib.load(f)

        button_open_data = tk.Button(
            parent_frame,
            text='Открыть папку',
            command=lambda: open_file_and_folder(
                f"{config_p_n["path_data"]}")  # Указываем папку для открытия
        )
        button_open_data.pack(pady=10)

        button_rrp_tracker = tk.Button(
            parent_frame,
            text='Открыть файл',
            command=lambda: open_file_and_folder(
                f"{config_p_n["readmy"]}")  # Указываем файл для открытия
        )
        button_rrp_tracker.pack(pady=10)
