# app/core/app_controller.py

import tkinter as tk
import threading
import queue
from tkinter import ttk
from typing import Dict, Type, List

from .module_loader import import_modules, BaseModule
from config.app_config import APP_TITLE, MODULES_DIR
from gui.main_module import MainModuleUI


class AppController:
    def __init__(self):
        self.app_title = APP_TITLE
        self.loaded_modules: Dict[str, Type[BaseModule]] = {}

        # --- АТРИБУТЫ ДЛЯ УПРАВЛЕНИЯ СЕМАФОРОМ ---
        # 1. Семафор с тремя "разрешениями" на запуск
        self.module_semaphore = threading.Semaphore(3)
        # 2. Очередь для передачи команд в GUI-поток
        self.command_queue: queue.Queue = queue.Queue()
        # 3. Список для хранения всех активных фреймов модулей
        self.pinned_module_frames: List[ttk.Frame] = []

        # 4. Ссылка на класс, отвечающий за UI
        self.ui_handler: MainModuleUI | None = None
        self.main_window: tk.Tk | None = None  # Явно инициализируем

        # Начинаем обработку очереди в фоновом режиме
        self._start_queue_processor()

    def _start_queue_processor(self):
        """
        Запускаем внутренний цикл обработки команд из очереди.
        """
        def _process():
            while True:
                command_func, *command_args = self.command_queue.get()
                if command_func is None:
                    break
                try:
                    command_func(*command_args)
                except Exception as e:
                    print(f"Ошибка в UI-команде: {e}")
                finally:
                    self.command_queue.task_done()
        threading.Thread(target=_process, daemon=True).start()

    def set_main_window(self, window: tk.Tk) -> None:
        """
        Вызывается из MainWindow после его инициализации.
        Это важно, чтобы мы могли безопасно взаимодействовать с окном.
        """
        self.main_window = window

    def create_ui(self, modules_frame: tk.Frame, content_frame: tk.Frame):
        """
        Создает UI для модулей.
        """
        # Создаем делегат UI и передаем ему фреймы
        self.ui_handler = MainModuleUI(modules_frame, content_frame)

        self.loaded_modules = import_modules(MODULES_DIR)
        for module_name, module_class in self.loaded_modules.items():
            # Используем делегат для создания UI-компонентов
            # Передаем ему ссылку на метод-обработчик, чтобы кнопка знала, что
            # вызывать
            self.ui_handler.create_module_button(
                module_name, module_class, self._handle_module_click
                )

    def _handle_module_click(self, module_class: Type[BaseModule]):
        """
        Обработчик нажатия кнопки модуля.
        """
        def _try_open():
            acquired = self.module_semaphore.acquire(blocking=True, timeout=0)
            if not acquired:
                print("[Ошибка] Нет свободного слота для модуля.")
                return

            # Если слот захвачен, ставим команду на создание GUI.
            # Передаем весь класс, который внутри себя будет содержать gui_run
            # (это нужно, если модуль его использует)
            self.command_queue.put((self._open_module_ui, module_class))

        threading.Thread(target=_try_open, daemon=True).start()

    def _open_module_ui(self, module_class: Type[BaseModule]):
        """
        Выполняется в GUI-потоке для создания и отображения фрейма модуля.
        """
        if not self.ui_handler:
            print(
                "[Ошибка] `ui_handler` не инициализирован в `_open_module_ui`."
                )
            return

        # Используем делегат UI для создания фрейма и получения ссылок на его
        # части
        new_module_frame, body_frame, header_frame = (
            self.ui_handler.create_module_frame(module_class)
            )

        # Добавляем фрейм в список для отслеживания
        self.pinned_module_frames.append(new_module_frame)

        # Создаем кнопку "Закрыть"
        close_button = ttk.Button(
            header_frame,
            text="✕",
            width=3,
            command=lambda f=new_module_frame: self._remove_pinned_frame(f)
            )
        close_button.pack(side='right')

        try:
            # Вызываем логику отрисовки самого модуля
            module_class.initialize_frame(body_frame)
        except Exception as e:
            tk.Label(
                body_frame,
                text=f"Ошибка загрузки модуля:\n{e}",
                fg="red"
                ).pack()

        # --- ИНИЦИАЦИЯ СОБЫТИЯ ПЕРЕРАСЧЕТА РАЗМЕРА ОКНА ---
        # Проверяем, что у нас есть доступ к main_window и `_resize_event`
        if self.main_window and hasattr(self.main_window, '_resize_event'):
            # Генерируем событие, которое было создано в MainWindow
            self.main_window.event_generate(self.main_window._resize_event)

    def _remove_pinned_frame(self, frame_to_remove: ttk.Frame):
        """
        Вспомогательный метод для удаления фрейма и освобождения слота.
        """
        # Проверка на ui_handler:
        if frame_to_remove in self.pinned_module_frames and self.ui_handler:
            # Используем делегат UI для физического удаления виджета
            self.ui_handler.remove_module_frame(frame_to_remove)

            # Удаляем фрейм из списка и освобождаем слот
            self.pinned_module_frames.remove(frame_to_remove)
            self.module_semaphore.release()
            print(
                f"[Инфо] Модуль закрыт. Доступно слотов: {
                    self.module_semaphore._value}"
                )

        # --- ИНИЦИАЦИЯ СОБЫТИЯ ПЕРЕРАСЧЕТА РАЗМЕРА ОКНА ---
        # Точно так же генерируем событие после удаления модуля
        if self.main_window and hasattr(self.main_window, '_resize_event'):
            self.main_window.event_generate(self.main_window._resize_event)

    def on_app_close(self):
        """
        Вызывается при закрытии приложения.
        """
        self.command_queue.put((None,))
        print("Приложение закрывается.")
