# app/core/app_controller.py

import tkinter as tk
import threading
import queue
from tkinter import ttk, messagebox
from typing import Dict, Type, List

from .module_loader import import_modules, BaseModule
from .module_api import ModuleAPI
from config.app_config import APP_TITLE, MODULES_DIR, MAX_CONCURRENT_MODULES
from gui.main_module import MainModuleUI


class AppController:
    def __init__(self):
        self.app_title = APP_TITLE
        self.loaded_modules: Dict[str, Type[BaseModule]] = {}

        # --- Атрибуты для управления семафором ---
        # 1. Семафор с определенным количеством "разрешений" на запуск
        self.module_semaphore = threading.Semaphore(MAX_CONCURRENT_MODULES)
        # 2. Очередь для передачи команд в GUI-поток
        self.command_queue: queue.Queue = queue.Queue()
        # 3. Список для хранения всех активных фреймов модулей
        self.pinned_module_frames: List[ttk.Frame] = []
        # 4. Ссылка на класс, отвечающий за UI
        self.ui_handler: MainModuleUI | None = None
        self.main_window: tk.Tk | None = None
        self.api: ModuleAPI | None = None
        self.notebook = None
        self.opened_tabs: List[ttk.Frame] = []

    def set_main_window(self, window: tk.Tk) -> None:
        """Вызывается из MainWindow после его создания."""
        self.main_window = window
        self.api = ModuleAPI(self, window)
        self._start_queue_processor()

    def _start_queue_processor(self):
        """Периодическая проверка очереди команд в главном потоке."""
        def process():
            try:
                while True:
                    item = self.command_queue.get_nowait()
                    if item is None:  # Сигнал завершения
                        break
                    func, *args = item
                    try:
                        func(*args)
                    except Exception as e:
                        print(f"Ошибка в UI-команде: {e}")
                    self.command_queue.task_done()
            except queue.Empty:
                pass
            finally:
                if self.main_window and self.main_window.winfo_exists():
                    self.main_window.after(100, process)

        self.main_window.after(100, process)

    def create_ui(self, modules_frame: tk.Frame, content_frame: tk.Frame):
        """Создаёт интерфейс модулей на левой панели."""
        # Создаем делегат UI и передаем ему фреймы
        self.ui_handler = MainModuleUI(modules_frame, content_frame, self)
        self.loaded_modules = import_modules(MODULES_DIR)
        for name, cls in self.loaded_modules.items():
            # Используем делегат для создания UI-компонентов.
            self.ui_handler.create_module_button(
                name, cls, self._handle_module_click
            )

    def set_notebook(self, notebook: ttk.Notebook):
        """Сохраняет ссылку на главный Notebook."""
        self.notebook = notebook

    def _handle_module_click(self, module_class: Type[BaseModule]):
        """Обработчик нажатия кнопки модуля."""
        def _try_open():
            acquired = self.module_semaphore.acquire(blocking=False)
            if not acquired:
                self.command_queue.put(
                    (lambda: messagebox.showwarning(
                        "Внимание",
                        "Достигнуто максимальное количество открытых модулей."
                    ),)
                )
                return

            # В зависимости от типа модуля, вызываем соответствующий метод
            if getattr(module_class, 'is_tab', False):
                self.command_queue.put((self._open_module_tab, module_class))
            else:
                self.command_queue.put((self._open_module_ui, module_class))

        threading.Thread(target=_try_open, daemon=True).start()

    def _open_module_tab(self, module_class: Type[BaseModule]):
        """Создаёт новую вкладку для модуля (выполняется в GUI-потоке)."""
        if not self.notebook:
            self.module_semaphore.release()
            return

        # Создаём фрейм для содержимого вкладки
        tab_frame = ttk.Frame(self.notebook)
        title = getattr(module_class, 'module_label', module_class.__name__)

        # --- Заголовок с кнопкой закрытия ---
        header_frame = ttk.Frame(tab_frame)
        header_frame.pack(fill='x', padx=5, pady=5)

        close_btn = ttk.Button(
            header_frame,
            text="✕",
            width=2,
            command=lambda: self._remove_tab(tab_frame)
        )
        close_btn.pack(side='right', anchor='ne')

        ttk.Label(
            header_frame,
            text=title,
            font=('TkDefaultFont', 10, 'bold')
        ).pack(side='left')

        # --- Тело модуля ---
        body_frame = ttk.Frame(tab_frame)
        body_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # Добавляем вкладку в Notebook
        self.notebook.add(
            tab_frame,
            text=getattr(module_class, 'button_text', title)
        )

        try:
            module_class.initialize_frame(body_frame, api=self.api)
        except Exception as e:
            error_msg = str(e)
            # При ошибке удаляем вкладку, освобождаем слот и показываем
            # сообщение
            self.notebook.forget(tab_frame)
            self.module_semaphore.release()
            self.command_queue.put(
                (lambda: messagebox.showerror(
                    "Ошибка модуля",
                    f"Не удалось загрузить модуль {title}:\n{error_msg}"
                ),)
            )
            return

        # Сохраняем в список открытых вкладок
        self.opened_tabs.append(tab_frame)

        # Переключаемся на новую вкладку
        self.notebook.select(tab_frame)

    def _open_module_ui(self, module_class: Type[BaseModule]):
        """
        Создаёт фрейм модуля и запускает его инициализацию (выполняется в
        GUI-потоке).
        """
        if not self.ui_handler:
            self.module_semaphore.release()
            return

        # Создаём контейнер для модуля
        container, body_frame, header_frame = (
            self.ui_handler.create_module_frame(module_class)
        )

        try:
            # Передаём API в модуль
            module_class.initialize_frame(body_frame, api=self.api)
        except Exception as e:
            # При ошибке показываем сообщение и удаляем фрейм
            tk.Label(
                body_frame,
                text=f"Ошибка загрузки модуля:\n{e}",
                fg="red"
            ).pack()
            # Не добавляем в список, освобождаем слот и удаляем контейнер
            self.module_semaphore.release()
            container.destroy()
            return

        # Успешно создан – добавляем в список и генерируем событие изменения
        # размера
        self.pinned_module_frames.append(container)
        if self.main_window and hasattr(self.main_window, '_resize_event'):
            self.main_window.event_generate(self.main_window._resize_event)

    def get_available_slots(self):
        """Возвращает количество свободных слотов."""
        return self.module_semaphore._value

    def _remove_pinned_frame(self, frame_to_remove: ttk.Frame):
        """Удаляет фрейм из списка активных и освобождает слот."""
        if frame_to_remove in self.pinned_module_frames:
            self.pinned_module_frames.remove(frame_to_remove)
            self.module_semaphore.release()

        # После удаления модуля, инициируем событие перерасчета размера окна
        if self.main_window and hasattr(self.main_window, '_resize_event'):
            self.main_window.event_generate(self.main_window._resize_event)

    def _remove_tab(self, tab_frame: ttk.Frame):
        """Закрывает вкладку и освобождает слот."""
        if tab_frame in self.opened_tabs:
            self.opened_tabs.remove(tab_frame)
            self.module_semaphore.release()
        try:
            self.notebook.forget(tab_frame)
        except tk.TclError:
            pass  # вкладка уже могла быть удалена

    def on_app_close(self):
        """Вызывается при закрытии приложения."""
        self.command_queue.put((None,))
        print("Приложение закрывается.")
