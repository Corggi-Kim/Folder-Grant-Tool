from collections.abc import Callable
from typing import Any

from PyQt5.QtCore import QObject, pyqtSignal, pyqtSlot


class ApiTask(QObject):
    succeeded = pyqtSignal(object)
    failed = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, operation: Callable[[], Any]):
        super().__init__()
        self.operation = operation

    @pyqtSlot()
    def run(self) -> None:
        try:
            self.succeeded.emit(self.operation())
        except Exception as exc:
            self.failed.emit(str(exc))
        finally:
            self.finished.emit()
