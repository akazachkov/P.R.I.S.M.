"""app/gui/elements/logo_manager.py"""

import tkinter as tk

from PIL import Image, ImageTk

from config.app_config import APP_LOGO


def setup_window_logo(window: tk.Tk):
    """
    Устанавливает логотип для окна Tkinter.

    Args:
    :param window: главное окно Tkinter
    """
    try:
        # 1. Загружаем изображение
        logo_image = Image.open(APP_LOGO)  # APP_LOGO - путь к файлу логотипа

        # 2. Преобразуем в PhotoImage
        logo_photo = ImageTk.PhotoImage(logo_image)

        # 3. Устанавливаем иконку окна
        window.wm_iconphoto(True, logo_photo)

        # Сохраняем ссылку на изображение, чтобы оно не было удалено сборщиком
        # мусора
        window.logo_photo = logo_photo
    except Exception as e:  # noqa: BLE001
        print(f"Ошибка загрузки логотипа: {e}")  # noqa: T201
