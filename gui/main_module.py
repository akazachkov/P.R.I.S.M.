# app/gui/main_module.py

import tkinter as tk
from tkinter import ttk

from config.app_config import GEOMETRY_MAIN_WINDOW, WRAPLENGHT_BUTTON
from core.module_loader import BaseModule


class MainModuleUI:
    """Управляет интерфейсом на вкладке 'Подключаемые модули'."""
    def __init__(
            self, modules_frame: ttk.Frame,
            content_frame: ttk.Frame,
            controller=None
    ):
        """
        Инициализирует класс UI.

        Args:
        :param modules_frame: фрейм для размещения кнопок модулей
        :param content_frame: фрейм для размещения содержимого модулей
        :param controller: ссылка на AppController для управления слотами
        """
        self.modules_frame = modules_frame
        self.content_frame = content_frame
        self._parent_controller = controller
        self._setup_scrollable_area()

    def _setup_scrollable_area(self):
        """Настраивает Canvas и Scrollbar для прокрутки содержимого модулей."""
        # Настройка скроллинга
        self.canvas = tk.Canvas(self.content_frame)
        scrollbar = ttk.Scrollbar(
            self.content_frame,
            orient="vertical",
            command=self.canvas.yview
        )
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.scrollable_frame_container = ttk.Frame(self.canvas)
        # Создаем прокси-окно на Canvas для нашего фрейма
        self.canvas_frame_id = self.canvas.create_window(
            (0, 0),
            window=self.scrollable_frame_container,
            anchor="nw",
            tags="frame"
        )

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Привязываем прокрутку колесом мыши
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        # Начальная настройка области прокрутки
        self._update_scrollregion()

    def _on_mousewheel(self, event: tk.Event):
        """Обработчик прокрутки колеса мыши."""
        # Прокручиваем Canvas, даже если мышь находится над дочерним фреймом.
        # event.delta обычно равен 120 (вверх) или -120 (вниз) на Windows
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _update_scrollregion(self):
        """Обновляет область прокрутки Canvas."""
        if self.canvas:
            self.canvas.update_idletasks()
            # Устанавливаем ширину прокси-окна равной ширине Canvas
            actual_width = self.canvas.winfo_width()
            if self.canvas_frame_id:
                self.canvas.itemconfigure(
                    self.canvas_frame_id,
                    width=actual_width - 10
                )
            # Теперь обновляем `scrollregion` на основе нового содержимого
            self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def update_canvas_width(self, new_width):
        """Обновляет ширину canvas при изменении размера окна."""
        if self.canvas and self.canvas_frame_id:
            self.canvas.itemconfigure(
                self.canvas_frame_id,
                width=self.canvas.winfo_width()
            )
            self._update_scrollregion()

    def create_module_button(
            self,
            module_name: str,
            module_class: type[BaseModule],
            on_click_handler
    ):
        """Создаёт кнопку модуля на левой панели."""
        button_text = getattr(module_class, 'button_text', module_name.title())

        # --- Создание лейбла над кнопкой открытия модуля ---
        # Проверяем наличие "button_label" в файле модуля
        if hasattr(module_class, 'button_label') and module_class.button_label:
            # Если "button_label" есть, то создаём лейбл
            label = tk.Label(
                self.modules_frame,
                text=module_class.button_label,
                anchor='w',
                wraplength=WRAPLENGHT_BUTTON,  # Максимальная ширина лейбла
                justify=tk.CENTER,  # Выравнивание текста по центру
                relief='groove'  # Стиль бордюра лейбла
            )
            label.pack(pady=(5, 0))

        button = tk.Button(
            self.modules_frame,
            text=button_text,
            wraplength=WRAPLENGHT_BUTTON,  # Максимальная ширина кнопки
            command=lambda m=module_class: on_click_handler(m)
        )
        button.pack(pady=5)

    def create_module_frame(self, module_class: type[BaseModule]):
        """
        Создаёт фрейм для нового модуля.
        Возвращает кортеж (container, body_frame, header_frame).
        """
        title = getattr(module_class, 'module_label', module_class.__name__)

        # Определяем тип модуля - с фиксированной шириной или динамической
        if module_class.width_frame is not None:
            # Фиксированная ширина
            container = ttk.Frame(self.scrollable_frame_container)
            container.pack(fill='x', pady=5)

            fixed_frame = ttk.Frame(
                container,
                width=module_class.width_frame,
                relief="solid"
            )
            fixed_frame.pack(expand=False, anchor='nw')
        else:
            # Динамическая ширина
            container = ttk.Frame(
                self.scrollable_frame_container,
                relief="solid",
                padding=5
            )
            container.pack(fill='x', pady=5)
            fixed_frame = container  # Для единообразия

        # Заголовок с кнопкой закрытия
        header_frame = ttk.Frame(fixed_frame)
        header_frame.pack(fill='x', padx=5, pady=5)

        close_button = ttk.Button(
            header_frame,
            text="✕",
            width=2,
            command=lambda: self._remove_module_frame_with_slot(container)
        )
        close_button.pack(side='right', anchor='ne')

        # Вычисляем ширину для переноса текста заголовка
        try:
            window_width = int(GEOMETRY_MAIN_WINDOW.split('x')[0])
            wrap_len = round(window_width / 1.5, -1) - 100
        except (ValueError, IndexError):
            wrap_len = 200  # Значение по умолчанию

        ttk.Label(
            header_frame,
            text=title,
            wraplength=(
                wrap_len if module_class.width_frame is None
                else module_class.width_frame
            ),
            justify='left'
        ).pack(side='left', fill='x')

        # Создаем тело модуля
        body_frame = ttk.Frame(fixed_frame)
        body_frame.pack(fill='both', padx=5, pady=5)

        # Для фиксированных фреймов вычисляем высоту по содержимому
        if module_class.width_frame is not None:
            fixed_frame.update_idletasks()
            required_height = (
                header_frame.winfo_reqheight()
                + body_frame.winfo_reqheight()
                + 20
            )
            fixed_frame.config(height=required_height)

        # Обновляем область прокрутки
        self._update_scrollregion()
        return container, body_frame, header_frame

    def _remove_module_frame_with_slot(self, frame_to_remove: ttk.Frame):
        """Удаляет фрейм модуля и освобождает слот в семафоре"""
        # Удаляем фрейм из UI
        try:
            frame_to_remove.destroy()
        except tk.TclError:
            pass
        finally:
            self._update_scrollregion()

        # Освобождаем слот в семафоре
        if self._parent_controller:
            self._parent_controller._remove_pinned_frame(frame_to_remove)
