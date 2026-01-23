import tkinter as tk

from core.module_loader import BaseModule


class NameNewModule(BaseModule):
    """
    Шаблон для нового подключаемого модуля.
    """
    # Метаданные модуля.
    name = "pattern"  # Внутреннее имя (может использоваться для логирования,
    # идентификации).
    button_label = "pattern"  # Отображается над кнопкой (опционально, если
    # поле не требуется - закомментировать строку).
    button_text = "pattern"  # Отображается на кнопке.
    module_label = "pattern"  # Описание - отображается в верхней части фрейма
    # модуля.
    width_frame = 350  # Указываем ширину для фрейма модуля (опционально, если
    # закомментировать строку, то ширина фрейма модуля будет равной ширине
    # основного окна).

    @classmethod
    def initialize_frame(cls, parent_frame: tk.Frame) -> None:
        """
        parent_frame - фрейм для содержимого модуля.
        """
        None
