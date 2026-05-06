"""app/app.py"""

import sys
import traceback

from config.app_config import CONFIG_PATHS_NAME
from core.app_controller import AppController
from gui.main_window import MainWindow


def main():
    """Запускает приложение."""
    try:
        # Передаём путь к конфигурации при создании контроллера
        controller = AppController(config_path=CONFIG_PATHS_NAME)
        app = MainWindow(controller)

        # Запускаем GUI цикл
        app.mainloop()
    except Exception as e:  # noqa: BLE001
        print("Ошибка при запуске приложения:", e)  # noqa: T201
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
