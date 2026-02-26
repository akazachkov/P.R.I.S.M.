# app/core/module_loader.py

import importlib.util
import inspect
from pathlib import Path
from typing import Dict, Type


# Определяем BaseModule здесь, чтобы избежать циклических импортов и чтобы все
# модули могли легко его импортировать
class BaseModule:
    """
    Абстрактный базовый класс для всех подключаемых модулей.
    Все модули должны наследоваться от этого класса и реализовывать
    метод initialize_frame.
    """
    name: str = None
    label: str = None
    button_text: str = None
    module_label: str = None
    width: int = None  # Ширина для титула (динамический фрейм)
    width_frame: int = None  # Ширина для всего фрейма (фиксированная ширина)

    @classmethod
    def initialize_frame(cls, parent_frame, api):
        """
        Инициализирует интерфейс модуля внутри parent_frame.
        :param parent_frame: родительский фрейм (ttk.Frame)
        :param api: объект ModuleAPI для взаимодействия с приложением
        """
        raise NotImplementedError("Модуль должен реализовать initialize_frame")


def import_modules(modules_dir: Path) -> Dict[str, Type[BaseModule]]:
    """
    Сканирует папку с модулями и импортирует все Python-файлы.
    Возвращает словарь {имя_модуля: класс_модуля}.
    """
    loaded_modules = {}
    for file_path in modules_dir.glob("*.py"):
        # Исключаем служебные и тестовые файлы, начинающиеся на "__"
        if file_path.name.startswith('__'):
            continue
        try:
            # Используем importlib для динамической загрузки
            spec = importlib.util.spec_from_file_location(
                file_path.stem,
                file_path
            )
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)

            # Ищем класс, наследующийся от BaseModule
            for name, obj in inspect.getmembers(module, inspect.isclass):
                # isinstance не работает с базовыми классами из importlib,
                # поэтому используем `issubclass`, но осторожно
                if issubclass(obj, BaseModule) and obj is not BaseModule:
                    module_class = obj
                    # Устанавливаем имя, если оно не задано в классе
                    if not module_class.name:
                        module_class.name = file_path.stem
                    if not module_class.label:
                        module_class.label = module_class.name.replace(
                            '_',
                            ' '
                        ).title()

                    # Определяем тип модуля по имени файла
                    if file_path.stem.startswith('tab_'):
                        module_class.is_tab = True   # Модуль-вкладка
                    else:
                        module_class.is_tab = False  # Модуль-фрейм

                    loaded_modules[file_path.stem] = module_class
                    print(
                        f"Загружен модуль: {file_path.stem} (тип: "
                        f"{'вкладка' if module_class.is_tab else 'фрейм'})"
                    )
        except Exception as e:
            print(f"Ошибка загрузки модуля {file_path}: {e}")
    return loaded_modules
