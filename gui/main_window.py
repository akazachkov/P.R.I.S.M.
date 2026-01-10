# app/gui/main_window.py

import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk

from core.app_controller import AppController
from config.app_config import APP_LOGO


class MainWindow(tk.Tk):
    """
    Управляет параметрами основного окна приложения.
    """
    def __init__(self, controller: AppController):
        super().__init__()
        self.controller = controller
        self.title(controller.app_title)
        self.geometry("650x250")
        self.minsize(650, 300)
        self.maxsize(1000, 900)

        # --- УСТАНОВКА ЛОГОТИПА В ЗАГОЛОВОК ОКНА ---
        try:
            # 1. Загружаем изображение
            logo_image = Image.open(APP_LOGO)

            # 2. Преобразуем в PhotoImage
            logo_photo = ImageTk.PhotoImage(logo_image)

            # 3. Устанавливаем иконку окна
            self.wm_iconphoto(True, logo_photo)  # True означает, что иконка
            # будет использоваться для всех окон
        except Exception as e:
            print(f"Ошибка загрузки логотипа: {e}")

        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        self._create_widgets()
        # Убедимся, что AppController будет знать о MainWindow после его
        # создания
        self.controller.set_main_window(self)

        # --- НАСТРОЙКА СОБЫТИЯ И ПОДПИСКИ ---
        # 1. Создаем имя для нашего виртуального события
        self._resize_event = "<<ModuleListChanged>>"

        # 2. "Подписываем" `_resize_to_content` на это событие.
        # Теперь, когда окно получит `<<ModuleListChanged>>`, оно вызовет
        # `_resize_to_content`.
        self.bind(self._resize_event, self._resize_to_content)

    def _create_widgets(self):
        # Notebook для вкладок
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)

        # Вкладка: "Подключаемые модули"
        self.modules_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.modules_tab, text="Подключаемые модули")

        # Фрейм для кнопок модулей
        self.modules_frame = ttk.Frame(self.modules_tab)
        self.modules_frame.pack(side='left', fill='y', padx=10, pady=10)

        # Фрейм для содержимого модулей
        self.content_frame = ttk.Frame(self.modules_tab)
        self.content_frame.pack(side='right', fill='both',
                                expand=True, padx=10, pady=10)

        # Запускаем создание UI из Controller
        self.controller.create_ui(self.modules_frame, self.content_frame)

    def _resize_to_content(self, event=None):
        """
        Изменяет размер основного окна, чтобы оно соответствовало высоте
        содержимого в Canvas, где находятся фреймы модулей.
        """

        # 0. Предварительные проверки
        if not (hasattr(self.controller, 'ui_handler')
                and self.controller.ui_handler
                and self.controller.ui_handler.canvas):
            return

        # 1. Убеждаемся, что все виджеты обновили свою геометрию.
        # Это нужно, чтобы `scrollregion` стала актуальной.
        # Мы должны обновлять именно Canvas, так как именно он содержит
        # контент.
        self.controller.ui_handler.canvas.update_idletasks()

        # 2. Получаем `scrollregion` из Canvas.
        try:
            scrollregion_bbox = self.controller.ui_handler.canvas.bbox("all")
            required_content_height = scrollregion_bbox[3]  # y2 (высота всего
            # содержимого)
        except (tk.TclError, TypeError, IndexError):
            required_content_height = 0

        # 3. Рассчитываем новый размер окна.
        # Теперь используем правильный, простой и надежный способ получения
        # `minsize` и `maxsize`: просто вызываем метод без аргументов.
        # `m = self.minsize()` вернет кортеж (ширина, высота), например (600,
        # 300).
        try:
            min_width, min_height = self.minsize()
        except (tk.TclError, ValueError):
            min_height = 300

        try:
            max_width, max_height = self.maxsize()
        except (tk.TclError, ValueError):
            max_height = 900

        # 4. Рассчитываем новую высоту окна.
        current_width = self.winfo_width()

        if required_content_height > 0:
            # Получаем текущую высоту окна и `content_frame` для расчета
            # высоты "шапки"
            current_window_height = self.winfo_height()
            content_frame_height = self.content_frame.winfo_height()

            # Расчет высоты "шапки" окна (заголовок + вкладки)
            header_height = current_window_height - content_frame_height

            # Требуемая высота окна = высота всего контента + высота "шапки"
            required_window_height = (
                required_content_height + header_height + 10
                )  # `+10` для отступов

            # Ограничиваем требуемую высоту окна `minsize` и `maxsize`
            new_height = min(
                max(required_window_height, min_height),
                max_height
                )
        else:
            # Если модулей нет, возвращаемся к минимальной высоте
            new_height = min_height

        # 5. Устанавливаем новый размер окна, сохраняя текущую ширину.
        if new_height != self.winfo_height():
            self.geometry(f"{current_width}x{new_height}")

            # Для отладки:
            print(f"[Отладка] Высота содержимого: {required_content_height}")
            print(f"[Отладка] Требуемая высота окна: {required_window_height}")
            print(f"[Отладка] Новый размер окна: {current_width}x{new_height}")

    def on_closing(self):
        """Вызывается при закрытии окна"""
        self.controller.on_app_close()
        self.destroy()
