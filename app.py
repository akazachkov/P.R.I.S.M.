# app/app.py

import sys
import traceback

from core.app_controller import AppController
from gui.main_window import MainWindow


def main():
    try:
        controller = AppController()
        app = MainWindow(controller)

        # Запускаем GUI цикл
        app.mainloop()
    except Exception as e:
        print("Ошибка при запуске приложения:", e)
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    """
    Запускает приложение.
    """
    main()
