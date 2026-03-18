import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QFont
from ui.main_window import MainWindow

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("PPT 场景合成工具")
    app.setStyle("Fusion")
    # Use system font (SF Pro on macOS)
    f = app.font(); f.setPointSize(13); app.setFont(f)

    window = MainWindow(templates_dir=os.path.join(BASE_DIR, "templates"))
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
