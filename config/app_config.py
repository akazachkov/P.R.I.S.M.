# app/config/app_config.py

from pathlib import Path


# Определяем корневую директорию проекта и путь к папке с модулями.
PROJECT_ROOT = Path(__file__).resolve().parent.parent
MODULES_DIR = PROJECT_ROOT / "modules"

# Путь к файлу с конфигурацией.
CONFIG_PATHS_NAME = PROJECT_ROOT / "config" / "config_paths_name.toml"

# Название приложения, логотип в строке заголовка.
APP_TITLE = "The F.I.R. v.2 (File Interaction Runner)"
APP_LOGO = PROJECT_ROOT / "images" / "logo_small.png"

# Геометрия приложения.
GEOMETRY_MAIN_WINDOW = "700x300"  # Начальный/минимальный размер окна.
MAIN_WINDOW_MAXSIZE = "1500x900"  # Максимальный размер окна.
WRAPLENGHT_BUTTON = 200  # Максимальная ширина лейбла, над кнопкой модуля.

# Количество одновременно работающих подключаемых модулей.
MAX_CONCURRENT_MODULES = 3
