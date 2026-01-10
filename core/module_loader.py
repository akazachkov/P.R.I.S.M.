# app/core/module_loader.py

import importlib.util
import inspect
from pathlib import Path
from typing import Dict, Type


# Определяем BaseModule здесь, чтобы избежать циклических импортов
# и чтобы все модули могли легко его импортировать.
class BaseModule:
    """
    Абстрактный базовый класс для всех подключаемых модулей.
    """
    name: str = None
    label: str = None
    button_text: str = None

    @classmethod
    def initialize_frame(cls, parent_frame):
        pass


def import_modules(modules_dir: Path) -> Dict[str, Type[BaseModule]]:
    """
    Сканирует папку с модулями и импортирует все Python-файлы.
    Возвращает словарь {имя_модуля: класс_модуля}.
    """
    if not modules_dir.exists():
        modules_dir.mkdir(exist_ok=True)

    loaded_modules = {}

    for file_path in modules_dir.glob("*.py"):
        # Исключаем __init__ и другие служебные файлы
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
                # поэтому используем `issubclass`, но осторожно.
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

                    loaded_modules[file_path.stem] = module_class
                    print(f"Загружен модуль: {file_path.stem}")
        except Exception as e:
            print(f"Ошибка загрузки модуля {file_path}: {e}")

    return loaded_modules
