# app/config/app_config.py

from pathlib import Path


# Определяем корневую директорию проекта и путь к папке с модулями
PROJECT_ROOT = Path(__file__).resolve().parent.parent
MODULES_DIR = PROJECT_ROOT / "modules"

# Путь к файлу с конфигурацией
CONFIG_PATHS_NAME = PROJECT_ROOT / "config" / "config_paths_name.toml"

# Название приложения, логотип в строке заголовка
APP_TITLE = "P.R.I.S.M. v.1.2.2"
# APP_LOGO = PROJECT_ROOT / "images" / "logo.png"

# Геометрия приложения
GEOMETRY_MAIN_WINDOW = "800x850"  # Начальный/минимальный размер окна (ШхВ)
MAIN_WINDOW_MAXSIZE = "1920x1920"  # Максимальный размер окна
WRAPLENGHT_BUTTON = 200  # Максимальная ширина лейбла, над кнопкой модуля

# Количество одновременно работающих подключаемых модулей
MAX_CONCURRENT_MODULES = 6
