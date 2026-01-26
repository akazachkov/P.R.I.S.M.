# app/gui/main_window.py

import tkinter as tk
from tkinter import ttk

from core.app_controller import AppController
from config.app_config import (
    GEOMETRY_MAIN_WINDOW, MAIN_WINDOW_MAXSIZE, MAX_CONCURRENT_MODULES
)

# На перспективу:
# from gui.elements.logo_manager import setup_window_logo


class MainWindow(tk.Tk):
    """
    Управляет параметрами основного окна приложения.
    """
    def __init__(self, controller: AppController):
        super().__init__()
        self.controller = controller
        self.title(controller.app_title)

        # Устанавливаем геометрию окна.
        self.geometry(GEOMETRY_MAIN_WINDOW)

        def _parse_geometry(geometry_str: str) -> tuple:
            """
            Разбирает строку геометрии в формате "WxH" и возвращает кортеж
            (ширина, высота).
            """
            return tuple(map(int, geometry_str.split('x')))

        # Устанавливаем minsize и maxsize окна, распаковывая кортежи.
        # Минимальный размер = начальному.
        self.minsize(*_parse_geometry(GEOMETRY_MAIN_WINDOW))
        self.maxsize(*_parse_geometry(MAIN_WINDOW_MAXSIZE))

        # Установка лого в заголовок окна (на перспективу).
        # setup_window_logo(self)

        # Создаём обработчик события закрытия окна.
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Запускаем создание виджетов
        self._create_widgets()
        # Убедимся, что AppController будет знать о MainWindow после его
        # создания.
        self.controller.set_main_window(self)

        # --- Настройка масштабирования фреймов подключаемых модулей ---
        # Добавляем привязку к событию "ручного" изменения размера окна.
        self.bind("<Configure>", self._on_window_resize)
        # Добавляем переменную для хранения текущей ширины.
        width, height = map(int, GEOMETRY_MAIN_WINDOW.split('x'))
        self.current_width = width

        # --- Настройка изменения высоты окна при открытии/закрытии модулей ---
        # 1. Создаем имя для нашего виртуального события.
        self._resize_event = "<<ModuleListChanged>>"

        # 2. "Подписываем" `_resize_to_content` на это событие.
        # Теперь, когда окно получит `<<ModuleListChanged>>`, оно вызовет
        # `_resize_to_content`.
        self.bind(self._resize_event, self._resize_to_content)

    def _create_widgets(self):
        # Notebook для вкладок.
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)

        # Вкладка: "Подключаемые модули".
        self.modules_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.modules_tab, text="Подключаемые модули")

        # Фрейм для кнопок модулей.
        self.modules_frame = ttk.Frame(self.modules_tab)
        self.modules_frame.pack(side='left', fill='y', padx=10, pady=10)

        # Фрейм для содержимого модулей.
        self.content_frame = ttk.Frame(self.modules_tab)
        self.content_frame.pack(side='right', fill='both',
                                expand=True, padx=10, pady=10)

        # Запускаем создание UI из Controller.
        self.controller.create_ui(self.modules_frame, self.content_frame)

    def _on_window_resize(self, event=None):
        """
        Обработчик изменения размера окна.
        Обновляет ширину динамических фреймов модулей.
        """
        new_width = self.winfo_width()
        if new_width != self.current_width:
            self.current_width = new_width
            # Обновляем ширину Canvas.
            if (hasattr(self.controller, 'ui_handler') and
                    self.controller.ui_handler):
                self.controller.ui_handler.update_canvas_width(new_width)

    def _resize_to_content(self, event=None):
        """
        Изменяет размер основного окна, чтобы оно соответствовало высоте
        содержимого в Canvas, где находятся фреймы модулей.
        """
        # 0. Предварительные проверки.
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
            # Получаем ограничивающую рамку (bounding box) для всех элементов
            # на Canvas. bbox("all") возвращает кортеж из 4 чисел:
            # (x1, y1, x2, y2)
            # [0] x1, [1] y1 - координаты верхнего левого угла.
            # [2] x2, [3] y2 - координаты нижнего правого угла.
            scrollregion_bbox = self.controller.ui_handler.canvas.bbox("all")
            required_content_height = scrollregion_bbox[3]  # [3] y2
        # В случае ошибки, окно будет установлено в минимальную высоту.
        except (tk.TclError, TypeError, IndexError):
            required_content_height = 0

        # 3. Рассчитываем новый размер окна.
        # Получаем `minsize` и `maxsize`.
        # `m = self.minsize()` вернет кортеж (ширина, высота), из которого нам
        # нужна только высота. С maxsize - аналогично.
        try:
            min_width, min_height = self.minsize()
        except (tk.TclError, ValueError):
            min_height = 650

        try:
            max_width, max_height = self.maxsize()
        except (tk.TclError, ValueError):
            max_height = 1500

        # 4. Сохраняем текущую ширину окна перед изменением высоты.
        current_width = self.winfo_width()

        # Получаем количество свободных слотов для модулей.
        available_slots = self.controller.get_available_slots()

        # Рассчитываем новую высоту окна.
        if available_slots != MAX_CONCURRENT_MODULES:
            # Если есть запущенные модули, то получаем текущую высоту окна и
            # `content_frame` для расчета высоты "шапки".
            current_window_height = self.winfo_height()
            content_frame_height = self.content_frame.winfo_height()

            # Расчет высоты "шапки" окна (заголовок окна и панель вкладок).
            header_height = current_window_height - content_frame_height

            # Требуемая высота окна = высота всего контента + высота "шапки".
            required_window_height = (
                required_content_height + header_height + 10
            )  # `+10` для отступов

            # Ограничиваем требуемую высоту окна `minsize` и `maxsize`.
            new_height = min(
                max(required_window_height, min_height),
                max_height
            )

        else:
            # Если модули закрыты, то возвращаемся к начальным размерам окна.
            new_height = min_height
            current_width = min_width

        # 5. Устанавливаем новый размер окна.
        self.geometry(f"{current_width}x{new_height}")

        # Для отладки:
        print(f"[Отладка] Высота содержимого: {required_content_height}")
        print(f"[Отладка] Новый размер окна: {current_width}x{new_height}")

    def on_closing(self):
        """
        Вызывается при закрытии окна.
        """
        self.controller.on_app_close()
        self.destroy()
