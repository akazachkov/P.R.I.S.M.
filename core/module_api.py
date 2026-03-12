# app/core/module_api.py

import datetime
import tomllib
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple, Union

from config.app_config import CONFIG_PATHS_NAME
from core.elements.background_task import BackgroundTaskManager
from core.elements.convert_register_to_list import transform_excel_list
from core.elements.copy_files import copy_files
from core.elements.excel_reader import ExcelReader
from core.elements.pdf_finder import PDFFinder
from core.elements.working_with_folders import (
    ensure_directory_exists,
    open_file_and_folder,
    parse_file_path
)


class ModuleAPI:
    """
    Единый интерфейс для взаимодействия модулей с приложением.
    """
    def __init__(self, controller, main_window):
        self._controller = controller
        self._main_window = main_window
        self._bg_manager = BackgroundTaskManager(self)
        self._config = None

    # ------------------------------------------------------------------
    # Конфигурация
    # ------------------------------------------------------------------
    def get_config(self, key: str = None) -> Any:
        """Возвращает значение конфигурации по ключу или весь словарь."""
        if self._config is None:
            with open(CONFIG_PATHS_NAME, "rb") as f:
                self._config = tomllib.load(f)
        if key is None:
            return self._config
        return self._config.get(key)

    # ------------------------------------------------------------------
    # Работа с папками и файлами
    # ------------------------------------------------------------------
    def ensure_directory_exists(
            self,
            base_path: Union[str, Path],
            subfolder: str
    ) -> Path:
        """Создаёт поддиректорию, если её нет."""
        return ensure_directory_exists(base_path, subfolder)

    def open_file_and_folder(self, path: Union[str, Path]) -> bool:
        """Открывает файл или папку в проводнике."""
        return open_file_and_folder(path)

    def parse_file_path(
            self,
            file_path: Union[str, Path]
    ) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """
        Разбирает путь к файлу на базовую директорию, номер и аббревиатуру.
        """
        return parse_file_path(file_path)

    # ------------------------------------------------------------------
    # Копирование файлов
    # ------------------------------------------------------------------
    def copy_files(
        self,
        source_dirs: List[Tuple[Union[str, Path], Optional[List[str]]]],
        target_dir: Union[str, Path],
        overwrite: bool = True,
        exclude_patterns: Optional[List[str]] = None
    ) -> List[str]:
        """Копирует файлы из нескольких источников в целевую папку."""
        return copy_files(source_dirs, target_dir, overwrite, exclude_patterns)

    def download_pdfs(
        self,
        files: List[Path],
        progress_callback: Optional[Callable[[int, int], None]] = None
    ) -> Tuple[List[Path], Path]:
        """
        Скачивает PDF-файлы в папку 'Загрузки' с созданием подпапки по дате.
        """
        from core.elements.copy_files import download_pdfs
        return download_pdfs(files, self, progress_callback)

    # ------------------------------------------------------------------
    # Работа с файлами Excel
    # ------------------------------------------------------------------
    def transform_excel_list(
            self,
            input_fn: Union[str, Path],
            output_file: Union[str, Path]
    ) -> None:
        """Преобразует ведомость в список оборудования."""
        transform_excel_list(input_fn, output_file)

    def read_excel_mapping(
        self,
        file_path: Union[str, Path],
        sheet_name: str,
        key_col: int,
        value_col: int,
        start_row: int = 2,
        validator: Optional[Callable] = None,
        default: Any = None
    ) -> Dict[str, Any]:
        """
        Читает Excel-файл (поддерживаются .xlsx и .xlsb) и возвращает
        словарь соответствия из двух столбцов.

        :param file_path: путь к файлу
        :param sheet_name: имя листа
        :param key_col: номер столбца с ключом (1-based)
        :param value_col: номер столбца со значением (1-based)
        :param start_row: строка начала данных (по умолчанию 2 – после
            заголовка)
        :param validator: функция, принимающая значение и возвращающая True,
            если его следует использовать (иначе подставляется default)
        :param default: значение по умолчанию для не прошедших валидацию

        :return: словарь {ключ: значение}
        """
        with ExcelReader(file_path) as reader:
            return reader.get_mapping(
                sheet_name=sheet_name,
                key_col=key_col,
                value_col=value_col,
                start_row=start_row,
                validator=validator,
                default=default
            )

    # ------------------------------------------------------------------
    # Работа с файлами PDF
    # ------------------------------------------------------------------
    def get_pdf_finder(self) -> PDFFinder:
        """Возвращает настроенный экземпляр PDFFinder для поиска PDF-файлов."""
        config = self.get_config()
        return PDFFinder(
            root_folder=config.get("folder_ktd", ""),
            year_labels=config.get("years", []),
            subfolder_name=config.get("subfolder_name", "")
        )

    # ------------------------------------------------------------------
    # Управление слотами и окном
    # ------------------------------------------------------------------
    def get_available_slots(self) -> int:
        """Возвращает количество свободных слотов для модулей."""
        return self._controller.get_available_slots()

    def generate_resize_event(self):
        """Генерирует событие пересчёта высоты главного окна."""
        if self._main_window and hasattr(self._main_window, '_resize_event'):
            self._main_window.event_generate(self._main_window._resize_event)

    # ------------------------------------------------------------------
    # Работа с GUI-потоком
    # ------------------------------------------------------------------
    def schedule_gui_task(self, func, *args):
        """Планирует выполнение функции в главном потоке GUI."""
        self._controller.command_queue.put((func, *args))

    def run_in_background(
        self,
        target: Callable,
        on_success: Optional[Callable] = None,
        on_error: Optional[Callable] = None,
        *args,
        **kwargs
    ) -> None:
        """
        Запускает целевую функцию в фоновом потоке и гарантирует,
        что колбэки on_success/on_error будут выполнены в главном потоке.

        Параметры:
        :param target: вызываемый объект (функция), который будет выполнен в
            фоне. Может возвращать любой результат.
        :param on_success: опциональный колбэк, вызываемый в главном потоке при
            успешном завершении target. Получает результат target в качестве
            аргумента.
        :param on_error: опциональный колбэк, вызываемый в главном потоке при
            возникновении исключения в target. Получает строку с описанием
            ошибки.
        :param *args: позиционные аргументы, передаваемые в target.
        :param **kwargs: именованные аргументы, передаваемые в target.

        Примечания:
        - Исключения, возникшие в target, автоматически логируются через
        api.log.
        - Если on_error не указан, пользователю будет показано стандартное
        диалоговое окно с ошибкой (через messagebox.showerror).
        - Метод не блокирует вызывающий поток и возвращает управление
        немедленно.
        """
        self._bg_manager.run(target, on_success, on_error, *args, **kwargs)

    # ------------------------------------------------------------------
    # Логирование
    # ------------------------------------------------------------------
    def log(self, message: str, level: str = "info"):
        """Записывает сообщение в лог (пока только в консоль)."""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] [API {level}] {message}")
