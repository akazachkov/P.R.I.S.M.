# app/gui/main_module.py

import tkinter as tk
from tkinter import ttk
from typing import Type

from core.module_loader import BaseModule


class MainModuleUI:
    """
    Управляет пользовательским интерфейсом на вкладке "Подключаемые модули".
    """
    def __init__(self, modules_frame: ttk.Frame, content_frame: ttk.Frame):
        """
        Инициализирует класс UI.

        :param modules_frame: Фрейм для размещения кнопок модулей.
        :param content_frame: Фрейм для размещения содержимого модулей.
        """
        self.modules_frame = modules_frame
        self.content_frame = content_frame

        # Виджеты для скроллинга
        self.canvas = None
        self.scrollable_frame_container = None
        self.canvas_frame_id = None

        self._setup_scrollable_area()

    def _setup_scrollable_area(self):
        """
        Настраивает Canvas и Scrollbar для прокрутки содержимого модулей.
        """
        # --- НАСТРОЙКА СКРОЛЛИНГА ---
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

        # Упаковываем элементы Canvas
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Привязываем прокрутку колесом мыши
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # Начальная настройка области прокрутки
        self._update_scrollregion()

    def _update_scrollregion(self):
        """
        Обновляет область прокрутки Canvas.
        """
        if self.canvas:
            self.canvas.update_idletasks()

            # Устанавливаем ширину прокси-окна равной ширине Canvas
            actual_canvas_width = self.canvas.winfo_width()
            if self.canvas_frame_id:  # Проверка, что прокси-окно создано
                self.canvas.itemconfigure(
                    self.canvas_frame_id, width=actual_canvas_width)

            # Теперь обновляем `scrollregion` на основе нового содержимого
            self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def _on_mousewheel(self, event: tk.Event):
        """
        Обработчик прокрутки колеса мыши.
        """
        # Прокручиваем Canvas, даже если мышь находится над дочерним фреймом.
        # event.delta обычно равен 120 (вверх) или -120 (вниз) на Windows
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def create_module_button(self, module_name: str,
                             module_class: Type[BaseModule],
                             on_button_click_handler):
        """
        Создает кнопку и метку для модуля на левой панели.
        Вызывается AppController'ом.
        """
        label_text = getattr(module_class, 'button_label', module_name.title())
        button_text = getattr(module_class, 'button_text', label_text)

        # --- СОЗДАНИЕ ЛЕЙБЛА (МЕТКИ) СЛЕВА ---
        label = tk.Label(
            self.modules_frame,
            text=label_text,
            anchor='w',
            wraplength=200,  # Максимальная ширина блока лейбла
            justify=tk.CENTER,  # Выравнивание текста по центру
            relief='groove'  # Стиль бордюра лейбла
            )
        label.pack(pady=(5, 0))

        # --- СОЗДАНИЕ КНОПКИ ---
        button = tk.Button(
            self.modules_frame,
            text=button_text,
            command=lambda m=module_class: on_button_click_handler(m)
            )
        button.pack(pady=5)

    def create_module_frame(self, module_class: Type[BaseModule]):
        """
        Создает фрейм для нового модуля и возвращает его.
        Вызывается AppController'ом.

        :return: Кортеж (new_module_frame, body_frame)
        """
        # --- ШАГ 1: СОЗДАНИЕ КОНТЕЙНЕРА ---
        new_module_frame = ttk.Frame(
            self.scrollable_frame_container,
            borderwidth=2,
            relief="groove"
            )
        new_module_frame.pack(
            fill="none",
            expand=False,
            pady=5,
            padx=5,
            anchor='nw'
            )

        # --- ШАГ 2: СОЗДАНИЕ ЗАГОЛОВКА ---
        header_frame = ttk.Frame(new_module_frame)
        header_frame.pack(fill="x", pady=5, padx=5)

        module_label_text = getattr(
            module_class, 'module_label', module_class.__name__)
        ttk.Label(
            header_frame,
            text=module_label_text,
            font=('Arial', 10),
            anchor='w',
            wraplength=320,  # Максимальная ширина блока лейбла
            justify=tk.LEFT
            ).pack(side='left')

        # --- ШАГ 3: Возвращаем фреймы для дальнейшего использования ---
        # AppController сам создаст кнопку "Закрыть" и body_frame,
        # но мы даем ему базовый контейнер `new_module_frame`

        body_frame = ttk.Frame(new_module_frame)
        body_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Обновляем область прокрутки
        self._update_scrollregion()

        return new_module_frame, body_frame, header_frame

    def remove_module_frame(self, frame_to_remove: ttk.Frame):
        """
        Удаляет фрейм модуля из UI.
        """
        try:
            frame_to_remove.destroy()
        except tk.TclError:
            pass
        finally:
            self._update_scrollregion()
