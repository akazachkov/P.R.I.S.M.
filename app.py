"""app/app.py"""

import logging
import sys

from config.app_config import CONFIG_PATHS_NAME
from core.app_controller import AppController
from gui.main_window import MainWindow

logger = logging.getLogger(__name__)


def main():
    """Запускает приложение."""
    try:
        # Передаём путь к конфигурации при создании контроллера
        controller = AppController(config_path=CONFIG_PATHS_NAME)
        app = MainWindow(controller)

        # Запускаем GUI цикл
        app.mainloop()
    except Exception:
        logger.exception("Ошибка при запуске приложения")
        sys.exit(1)


if __name__ == "__main__":
    main()
