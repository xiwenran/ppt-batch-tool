"""PPT 批量导出图片工具 — 入口。"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt6.QtWidgets import QApplication
from ui.main_window import MainWindow

APP_NAME = "PPT转图片"

try:
    from _build_info import BUILD
except ImportError:
    BUILD = "dev"


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setStyle("Fusion")
    f = app.font()
    f.setPointSize(13)
    app.setFont(f)

    window = MainWindow(build=BUILD)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
