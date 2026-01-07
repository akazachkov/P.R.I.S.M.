from pathlib import Path


# Определяем корневую директорию проекта
PROJECT_ROOT = Path(__file__).resolve().parent.parent
MODULES_DIR = PROJECT_ROOT / "modules"
CONFIG_PATHS_NAME = PROJECT_ROOT / "config" / "config_paths_name.toml"

# --- Application Constants ---
APP_TITLE = "The F.I.R. 2.0a (File Interaction Runner)"
APP_LOGO = PROJECT_ROOT / "images" / "logo_small.png"
