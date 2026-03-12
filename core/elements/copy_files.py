# app/core/elements/copy_files.py

import shutil
import datetime
from pathlib import Path
from typing import List, Tuple, Optional, Union, Callable


def copy_files(
    source_dirs: List[Tuple[Union[str, Path], Optional[List[str]]]],
    target_dir: Union[str, Path],
    overwrite: bool = True,
    exclude_patterns: Optional[List[str]] = None
) -> List[str]:
    """
    Универсальная функция для копирования файлов из нескольких директорий.

    Args:
    :param source_dirs: список кортежей (путь_к_папке, [фильтры_имен_файлов]).
        Если фильтры None: копируются все файлы
    :param target_dir: целевая директория
    :param overwrite: разрешить перезапись существующих файлов
    :param exclude_patterns: шаблоны для исключения файлов (напр., ["~$*"])

    Returns:
        список сообщений о результатах копирования
    """
    target_dir = Path(target_dir)
    results = []

    # Устанавливаем шаблоны исключения по умолчанию
    if exclude_patterns is None:
        exclude_patterns = ["~$*"]  # Исключаем временные файлы MS Office

    # Создаем целевую директорию, если она не существует
    target_dir.mkdir(parents=True, exist_ok=True)

    for source_dir, name_filters in source_dirs:
        source_path = Path(source_dir)
        if not source_path.exists():
            results.append(f"Ошибка: Директория {source_path} не существует")
            continue

        # Определяем, какие файлы копировать
        if name_filters is None:
            # Копируем все файлы, кроме исключённых
            files_to_copy = [
                f for f in source_path.iterdir() if f.is_file() if
                f.is_file() and not any(
                    f.match(pattern) for pattern in exclude_patterns
                )
            ]
        else:
            # Копируем файлы, содержащие любое из указанных слов в имени
            files_to_copy = [
                f for f in source_path.iterdir()
                if f.is_file()
                and any(
                    filter_word.lower() in f.name.lower() for
                    filter_word in name_filters
                )
                and not any(f.match(pattern) for pattern in exclude_patterns)
            ]

        # Копируем файлы
        for file in files_to_copy:
            try:
                dest = target_dir / file.name
                if dest.exists() and not overwrite:
                    results.append(
                        f"Пропущен {file.name} (файл уже существует)"
                    )
                    continue
                shutil.copy2(file, dest)
                results.append(
                    f"Из {source_path.name} - {file.name}"
                )
            except Exception as e:
                results.append(f"Ошибка копирования {file.name}: {str(e)}")
    return results if results else ["Нет файлов для копирования"]


def download_pdfs(
    files: List[Path],
    api,
    progress_callback: Optional[Callable[[int, int], None]] = None
) -> Tuple[List[Path], Path]:
    """
    Копирование PDF-файлов в папку "Загрузки" в подпапку с текущей датой.
    При конфликтах имён добавляется числовой суффикс.

    Args:
    :param files: список путей к файлам для копирования
    :param api: объект ModuleAPI для логирования
    :param progress_callback: опциональная функция, принимающая
        (current, total)

    Returns:
        кортеж (список скопированных файлов, путь к созданной папке)
    """
    current_date = datetime.datetime.now().strftime("%Y.%m.%d")
    folder_name = f"{current_date} Скачанные PDF"
    downloads_dir = Path.home() / "Downloads" / folder_name

    counter = 1
    original_dir = downloads_dir
    while downloads_dir.exists():
        downloads_dir = (
            original_dir.parent / f"{original_dir.name} ({counter})"
        )
        counter += 1

    downloads_dir.mkdir(parents=True)
    api.log(f"Создана папка для скачивания: {downloads_dir}", "info")

    copied = []
    total = len(files)

    for i, filepath in enumerate(files):
        try:
            dest = downloads_dir / filepath.name
            # Если файл уже существует, добавляем суффикс
            file_counter = 1
            while dest.exists():
                stem = filepath.stem
                # Если имя уже содержит _цифры в конце, отрезаем их
                parts = stem.split('_')
                if parts[-1].isdigit():
                    base = '_'.join(parts[:-1])
                else:
                    base = stem
                dest = (
                    downloads_dir / f"{base}_{file_counter}{filepath.suffix}"
                )
                file_counter += 1

            shutil.copy2(filepath, dest)
            copied.append(dest)
            api.log(f"Скопирован: {dest}", "info")

            if progress_callback:
                progress_callback(i + 1, total)

        except Exception as e:
            api.log(f"Ошибка при копировании {filepath}: {e}", "error")

    return copied, downloads_dir
