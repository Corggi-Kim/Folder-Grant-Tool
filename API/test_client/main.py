import sys

from PyQt5.QtWidgets import QApplication

from .window import TestClientWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Folder Grant API Test Client")
    window = TestClientWindow()
    window.show()
    return app.exec_()


if __name__ == "__main__":
    raise SystemExit(main())
