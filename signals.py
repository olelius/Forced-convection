from PyQt5.QtCore import QObject, pyqtSignal


class ButtonTextChangedSignal(QObject):
    button_text_changed = pyqtSignal(str)
