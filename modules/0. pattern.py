import tkinter as tk

from core.module_loader import BaseModule


class NameNewModule(BaseModule):
    """
    Шаблон для нового подключаемого модуля.
    """
    # Метаданные модуля.
    name = "pattern"  # Внутреннее имя (может использоваться для логирования,
    # идентификации)
    button_label = "pattern"  # Отображается над кнопкой
    button_text = "pattern"  # Отображается на кнопке
    module_label = "pattern"  # Описание - отображается в верхней части фрейма
    # модуля

    @classmethod
    def initialize_frame(cls, parent_frame: tk.Frame) -> None:
        """
        parent_frame - фрейм для содержимого модуля.
        """
        None
