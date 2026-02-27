# app/core/module_api.py

import tomllib
from pathlib import Path
from typing import Any, List, Tuple, Union, Optional

from config.app_config import CONFIG_PATHS_NAME
from core.elements.convert_register_to_list import transform_excel_list
from core.elements.copy_files import copy_files
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

    # ------------------------------------------------------------------
    # Конвертация ведомости в список
    # ------------------------------------------------------------------
    def transform_excel_list(
            self,
            input_fn: Union[str, Path],
            output_file: Union[str, Path]
    ) -> None:
        """Преобразует ведомость в список оборудования."""
        transform_excel_list(input_fn, output_file)

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

    # ------------------------------------------------------------------
    # Логирование
    # ------------------------------------------------------------------
    def log(self, message: str, level: str = "info"):
        """Записывает сообщение в лог (пока в консоль)."""
        print(f"[API {level}] {message}")
