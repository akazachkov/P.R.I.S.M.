# app/core/elements/background_task.py

import threading
import traceback
from typing import Callable, Optional, Any


class BackgroundTaskManager:
    """
    Менеджер для выполнения длительных задач в фоновом потоке
    и возврата результата в главный поток через API.
    """

    def __init__(self, api):
        """
        :param api: экземпляр ModuleAPI для доступа к методам
            schedule_gui_task и log.
        """
        self._api = api

    def run(
        self,
        target: Callable,
        on_success: Optional[Callable[[Any], None]] = None,
        on_error: Optional[Callable[[str], None]] = None,
        *args,
        **kwargs
    ) -> None:
        """
        Запускает целевую функцию в отдельном потоке.

        :param target: функция, выполняемая в фоне (может возвращать
            результат).
        :param on_success: колбэк, вызываемый в главном потоке при успешном
            завершении, получает результат target.
        :param on_error: колбэк, вызываемый в главном потоке при ошибке,
            получает строку с описанием ошибки.
        :param args: позиционные аргументы для target.
        :param kwargs: именованные аргументы для target.
        """
        def wrapper():
            try:
                result = target(*args, **kwargs)
                if on_success:
                    self._api.schedule_gui_task(on_success, result)
            except Exception as e:
                # Формируем подробное сообщение об ошибке
                error_msg = (
                    f"{type(e).__name__}: {e}\n{traceback.format_exc()}"
                )
                self._api.log(error_msg, "error")

                if on_error:
                    # Если передан пользовательский обработчик ошибок
                    self._api.schedule_gui_task(on_error, str(e))
                else:
                    # По умолчанию показываем всплывающее окно с ошибкой
                    self._api.schedule_gui_task(
                        lambda err: self._api.messagebox.showerror(
                            "Ошибка", err),
                        str(e)
                    )

        thread = threading.Thread(target=wrapper, daemon=True)
        thread.start()
