"""PPT 批量导出图片 — 主窗口。"""

import os
import subprocess
import sys
from typing import List, Optional

from PyQt6.QtCore import Qt, QSettings
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QApplication, QComboBox, QFileDialog, QHBoxLayout, QLabel, QMainWindow,
    QMessageBox, QProgressBar, QPushButton, QScrollArea, QSpinBox,
    QLineEdit, QTextEdit, QVBoxLayout, QWidget,
)

from core.converter import (
    ConvertResult, ConvertWorker, detect_backends, backend_display_name,
    BACKEND_LIBREOFFICE, BACKEND_WPS_COM, BACKEND_PPT_COM,
    BACKEND_PPT_MAC, _find_libreoffice,
)
from core.scanner import scan_ppt_files

# ---------------------------------------------------------------------------
# 主题色（WeChat 风格）
# ---------------------------------------------------------------------------
_WIN = "#F7F7F7"
_CARD = "#FFFFFF"
_INPUT = "#F0F0F0"
_SEP = "#E5E5E5"
_TEXT = "#191919"
_TEXT2 = "#888888"
_GREEN = "#07C160"
_RED = "#FA5151"
_ORANGE = "#FF9500"


def _global_qss() -> str:
    return f"""
    QMainWindow, QWidget#central {{
        background: {_WIN};
    }}
    QWidget#card {{
        background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                        stop:0 #FFFFFF, stop:1 #FAFAFA);
        border-radius: 16px;
    }}
    QLabel {{
        color: {_TEXT};
    }}
    QLabel#h2 {{
        font-size: 16px;
        font-weight: bold;
        border-left: 4px solid {_GREEN};
        padding-left: 8px;
    }}
    QLabel#hint {{
        color: {_TEXT2};
        font-size: 12px;
    }}
    QLabel#badge_green {{
        color: white;
        background: {_GREEN};
        border-radius: 10px;
        padding: 2px 10px;
        font-size: 12px;
    }}
    QLabel#badge_red {{
        color: white;
        background: {_RED};
        border-radius: 10px;
        padding: 2px 10px;
        font-size: 12px;
    }}
    QLabel#badge_orange {{
        color: white;
        background: {_ORANGE};
        border-radius: 10px;
        padding: 2px 10px;
        font-size: 12px;
    }}
    QLineEdit {{
        background: {_INPUT};
        border: 1px solid {_SEP};
        border-radius: 8px;
        padding: 6px 10px;
        color: {_TEXT};
    }}
    QSpinBox {{
        background: {_INPUT};
        border: 1px solid {_SEP};
        border-radius: 8px;
        padding: 6px 10px;
        color: {_TEXT};
    }}
    QComboBox {{
        background: {_INPUT};
        border: 1px solid {_SEP};
        border-radius: 8px;
        padding: 6px 10px;
        color: {_TEXT};
    }}
    QPushButton {{
        border: none;
        border-radius: 18px;
        padding: 8px 20px;
        font-size: 13px;
        background: {_INPUT};
        color: {_TEXT};
    }}
    QPushButton:hover {{
        background: {_SEP};
    }}
    QPushButton#primary {{
        background: {_GREEN};
        color: white;
        border-radius: 22px;
        padding: 12px 40px;
        font-size: 15px;
        font-weight: bold;
        min-height: 44px;
    }}
    QPushButton#primary:hover {{
        background: #06AD56;
    }}
    QPushButton#primary:disabled {{
        background: #A8E6C1;
    }}
    QPushButton#danger {{
        background: {_RED};
        color: white;
        border-radius: 18px;
    }}
    QPushButton#danger:hover {{
        background: #E04545;
    }}
    QProgressBar {{
        background: {_INPUT};
        border: none;
        border-radius: 8px;
        height: 20px;
        text-align: center;
        color: {_TEXT};
    }}
    QProgressBar::chunk {{
        background: {_GREEN};
        border-radius: 8px;
    }}
    QTextEdit {{
        background: {_INPUT};
        border: 1px solid {_SEP};
        border-radius: 8px;
        padding: 6px;
        color: {_TEXT};
        font-size: 12px;
    }}
    QScrollArea {{
        border: none;
        background: transparent;
    }}
    """


class MainWindow(QMainWindow):

    def __init__(self, build: str = "dev"):
        super().__init__()
        self.setWindowTitle(f"PPT 转图片  [{build}]")
        self.setMinimumSize(640, 700)
        self.resize(700, 820)

        self._settings = QSettings("ppt2img", "PPT2Img")
        self._worker: Optional[ConvertWorker] = None
        self._ppt_files: List[str] = []
        self._backends: List[str] = []  # 所有检测到的后端
        self._soffice_path: Optional[str] = None

        self._build_ui()
        self._detect_engine()

    # ------------------------------------------------------------------
    # UI 构建
    # ------------------------------------------------------------------
    def _build_ui(self):
        self.setStyleSheet(_global_qss())

        central = QWidget()
        central.setObjectName("central")
        self.setCentralWidget(central)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_body = QWidget()
        self._layout = QVBoxLayout(scroll_body)
        self._layout.setContentsMargins(24, 24, 24, 24)
        self._layout.setSpacing(16)
        scroll.setWidget(scroll_body)

        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

        self._build_engine_card()
        self._build_folder_card()
        self._build_settings_card()
        self._build_action_card()
        self._build_result_card()
        self._build_log_card()

        self._layout.addStretch()

    def _make_card(self) -> tuple:
        card = QWidget()
        card.setObjectName("card")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(10)
        return card, layout

    # --- Card 0: 转换引擎状态 ---
    def _build_engine_card(self):
        card, lay = self._make_card()
        h = QLabel("转换引擎")
        h.setObjectName("h2")
        lay.addWidget(h)

        self._engine_badge = QLabel("检测中...")
        self._engine_badge.setObjectName("badge_green")
        lay.addWidget(self._engine_badge, alignment=Qt.AlignmentFlag.AlignLeft)

        self._engine_hint = QLabel("")
        self._engine_hint.setObjectName("hint")
        self._engine_hint.setWordWrap(True)
        lay.addWidget(self._engine_hint)

        # 引擎选择下拉框
        self._engine_row = QWidget()
        engine_lay = QHBoxLayout(self._engine_row)
        engine_lay.setContentsMargins(0, 0, 0, 0)
        engine_lay.addWidget(QLabel("选择引擎:"))
        self._engine_combo = QComboBox()
        self._engine_combo.currentIndexChanged.connect(self._on_engine_changed)
        engine_lay.addWidget(self._engine_combo, 1)
        self._engine_row.hide()
        lay.addWidget(self._engine_row)

        # LibreOffice 手动选择行（默认隐藏）
        self._lo_row = QWidget()
        lo_lay = QHBoxLayout(self._lo_row)
        lo_lay.setContentsMargins(0, 0, 0, 0)
        lo_lbl = QLabel("LibreOffice 路径:")
        self._lo_path = QLineEdit()
        self._lo_path.setReadOnly(True)
        self._lo_path.setPlaceholderText("未选择")
        lo_btn = QPushButton("浏览...")
        lo_btn.clicked.connect(self._browse_libreoffice)
        lo_lay.addWidget(lo_lbl)
        lo_lay.addWidget(self._lo_path, 1)
        lo_lay.addWidget(lo_btn)
        self._lo_row.hide()
        lay.addWidget(self._lo_row)

        self._layout.addWidget(card)

    # --- Card 1: 选择文件夹 ---
    def _build_folder_card(self):
        card, lay = self._make_card()
        h = QLabel("步骤 1: 选择 PPT 文件夹")
        h.setObjectName("h2")
        lay.addWidget(h)

        hint = QLabel("递归扫描文件夹中所有 PPT 文件（.ppt .pptx .pps .ppsx）")
        hint.setObjectName("hint")
        lay.addWidget(hint)

        row = QWidget()
        rl = QHBoxLayout(row)
        rl.setContentsMargins(0, 0, 0, 0)
        self._folder_input = QLineEdit()
        self._folder_input.setReadOnly(True)
        self._folder_input.setPlaceholderText("点击右侧按钮选择文件夹")
        browse = QPushButton("选择文件夹")
        browse.clicked.connect(self._browse_folder)
        rl.addWidget(self._folder_input, 1)
        rl.addWidget(browse)
        lay.addWidget(row)

        self._scan_badge = QLabel("")
        self._scan_badge.setObjectName("badge_green")
        self._scan_badge.hide()
        lay.addWidget(self._scan_badge, alignment=Qt.AlignmentFlag.AlignLeft)

        self._layout.addWidget(card)

    # --- Card 2: 导出设置 ---
    def _build_settings_card(self):
        card, lay = self._make_card()
        h = QLabel("步骤 2: 导出设置")
        h.setObjectName("h2")
        lay.addWidget(h)

        row1 = QWidget()
        r1l = QHBoxLayout(row1)
        r1l.setContentsMargins(0, 0, 0, 0)
        r1l.addWidget(QLabel("每个 PPT 最多导出前"))
        self._pages_spin = QSpinBox()
        self._pages_spin.setRange(1, 999)
        self._pages_spin.setValue(17)
        self._pages_spin.setSuffix(" 页")
        r1l.addWidget(self._pages_spin)
        r1l.addStretch()
        lay.addWidget(row1)

        row2 = QWidget()
        r2l = QHBoxLayout(row2)
        r2l.setContentsMargins(0, 0, 0, 0)
        r2l.addWidget(QLabel("输出文件夹:"))
        self._output_input = QLineEdit()
        self._output_input.setReadOnly(True)
        self._output_input.setPlaceholderText("默认为 PPT 文件夹下的「导出图片」子目录")
        out_btn = QPushButton("修改")
        out_btn.clicked.connect(self._browse_output)
        r2l.addWidget(self._output_input, 1)
        r2l.addWidget(out_btn)
        lay.addWidget(row2)

        self._layout.addWidget(card)

    # --- Card 3: 操作区 ---
    def _build_action_card(self):
        card, lay = self._make_card()

        btn_row = QWidget()
        bl = QHBoxLayout(btn_row)
        bl.setContentsMargins(0, 0, 0, 0)
        bl.addStretch()
        self._start_btn = QPushButton("开始转换")
        self._start_btn.setObjectName("primary")
        self._start_btn.clicked.connect(self._start_convert)
        self._start_btn.setEnabled(False)
        bl.addWidget(self._start_btn)

        self._cancel_btn = QPushButton("取消")
        self._cancel_btn.setObjectName("danger")
        self._cancel_btn.clicked.connect(self._cancel_convert)
        self._cancel_btn.hide()
        bl.addWidget(self._cancel_btn)
        bl.addStretch()
        lay.addWidget(btn_row)

        self._progress = QProgressBar()
        self._progress.hide()
        lay.addWidget(self._progress)

        self._status_label = QLabel("")
        self._status_label.setObjectName("hint")
        self._status_label.hide()
        lay.addWidget(self._status_label)

        self._layout.addWidget(card)

    # --- Card 4: 结果 ---
    def _build_result_card(self):
        self._result_card, lay = self._make_card()

        h = QLabel("转换结果")
        h.setObjectName("h2")
        lay.addWidget(h)

        self._result_summary = QLabel("")
        self._result_summary.setWordWrap(True)
        lay.addWidget(self._result_summary)

        self._result_detail = QTextEdit()
        self._result_detail.setReadOnly(True)
        self._result_detail.setMaximumHeight(200)
        self._result_detail.hide()
        lay.addWidget(self._result_detail)

        self._open_dir_btn = QPushButton("打开输出文件夹")
        self._open_dir_btn.clicked.connect(self._open_output_dir)
        lay.addWidget(self._open_dir_btn, alignment=Qt.AlignmentFlag.AlignLeft)

        self._result_card.hide()
        self._layout.addWidget(self._result_card)

    # --- Card 5: 日志 ---
    def _build_log_card(self):
        self._log_card, lay = self._make_card()

        h = QLabel("运行日志")
        h.setObjectName("h2")
        lay.addWidget(h)

        self._log_text = QTextEdit()
        self._log_text.setReadOnly(True)
        self._log_text.setMaximumHeight(180)
        lay.addWidget(self._log_text)

        self._log_card.hide()
        self._layout.addWidget(self._log_card)

    # ------------------------------------------------------------------
    # 引擎检测
    # ------------------------------------------------------------------
    def _detect_engine(self):
        self._backends = detect_backends()
        if self._backends:
            # 显示所有检测到的引擎
            names = [backend_display_name(b) for b in self._backends]
            primary = names[0]

            if len(self._backends) == 1:
                self._engine_badge.setText(f"已就绪: {primary}")
                self._engine_badge.setObjectName("badge_green")
                self._engine_hint.setText(
                    f"检测到 {primary}，将使用它来导出 PPT 幻灯片。\n"
                    f"如果转换失败，会自动尝试其他可用引擎。"
                )
            else:
                self._engine_badge.setText(f"已就绪: {primary}")
                self._engine_badge.setObjectName("badge_green")
                self._engine_hint.setText(
                    f"检测到 {len(self._backends)} 个引擎: {', '.join(names)}\n"
                    f"默认使用 {primary}，失败时自动切换备选引擎。\n"
                    f"也可以在下方手动选择优先使用的引擎。"
                )
                # 显示引擎选择下拉框
                self._engine_combo.blockSignals(True)
                self._engine_combo.clear()
                for b in self._backends:
                    self._engine_combo.addItem(
                        f"{backend_display_name(b)}（优先）", b
                    )
                self._engine_combo.setCurrentIndex(0)
                self._engine_combo.blockSignals(False)
                self._engine_row.show()

            # 设置 LibreOffice 路径
            if BACKEND_LIBREOFFICE in self._backends:
                self._soffice_path = _find_libreoffice()
        else:
            self._engine_badge.setText("未检测到转换引擎")
            self._engine_badge.setObjectName("badge_red")
            if sys.platform == "darwin":
                self._engine_hint.setText(
                    "未检测到 PowerPoint 或 LibreOffice。\n"
                    "请安装 Microsoft PowerPoint 或从 libreoffice.org 下载安装免费的 LibreOffice。\n"
                    "安装后可手动选择 LibreOffice 路径。"
                )
            else:
                self._engine_hint.setText(
                    "未检测到 WPS、PowerPoint 或 LibreOffice。\n"
                    "请安装 WPS Office、Microsoft Office 或从 libreoffice.org 下载安装免费的 LibreOffice。"
                )
            self._lo_row.show()
            saved = self._settings.value("soffice_path", "")
            if saved and os.path.isfile(saved):
                self._lo_path.setText(saved)
                self._soffice_path = saved
                self._backends = [BACKEND_LIBREOFFICE]
                self._engine_badge.setText("已就绪: LibreOffice")
                self._engine_badge.setObjectName("badge_green")
        # 刷新样式
        self._engine_badge.style().unpolish(self._engine_badge)
        self._engine_badge.style().polish(self._engine_badge)
        self._update_start_btn()

    def _on_engine_changed(self, index: int):
        """用户手动切换优先引擎时，调整后端顺序。"""
        if index < 0 or index >= len(self._backends):
            return
        # 把选中的引擎放到第一位，其余保持原序
        selected = self._backends[index]
        new_order = [selected] + [b for b in self._backends if b != selected]
        self._backends = new_order
        name = backend_display_name(selected)
        self._engine_badge.setText(f"已就绪: {name}")
        self._engine_hint.setText(
            f"已切换为优先使用 {name}，失败时自动尝试其他引擎。"
        )

    def _browse_libreoffice(self):
        if sys.platform == "darwin":
            path, _ = QFileDialog.getOpenFileName(
                self, "选择 LibreOffice soffice",
                "/Applications",
                "soffice (*)"
            )
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, "选择 soffice.exe",
                r"C:\Program Files",
                "soffice (soffice.exe)"
            )
        if path:
            self._lo_path.setText(path)
            self._soffice_path = path
            self._settings.setValue("soffice_path", path)
            if BACKEND_LIBREOFFICE not in self._backends:
                self._backends.append(BACKEND_LIBREOFFICE)
            self._engine_badge.setText("已就绪: LibreOffice")
            self._engine_badge.setObjectName("badge_green")
            self._engine_badge.style().unpolish(self._engine_badge)
            self._engine_badge.style().polish(self._engine_badge)
            self._update_start_btn()

    # ------------------------------------------------------------------
    # 文件夹选择
    # ------------------------------------------------------------------
    def _browse_folder(self):
        last = self._settings.value("last_input_dir", "")
        if sys.platform == "darwin":
            folder = self._mac_pick_folder("选择包含 PPT 文件的文件夹", last)
        else:
            folder = QFileDialog.getExistingDirectory(self, "选择 PPT 文件夹", last)
        if not folder:
            return

        self._settings.setValue("last_input_dir", folder)
        self._folder_input.setText(folder)

        # 立即扫描
        self._ppt_files = scan_ppt_files(folder)
        n = len(self._ppt_files)
        if n > 0:
            self._scan_badge.setText(f"找到 {n} 个 PPT 文件")
            self._scan_badge.setObjectName("badge_green")
        else:
            self._scan_badge.setText("未找到 PPT 文件")
            self._scan_badge.setObjectName("badge_red")
        self._scan_badge.style().unpolish(self._scan_badge)
        self._scan_badge.style().polish(self._scan_badge)
        self._scan_badge.show()

        # 设置默认输出目录
        if not self._output_input.text():
            self._output_input.setText(os.path.join(folder, "导出图片"))

        self._update_start_btn()

    def _browse_output(self):
        last = self._output_input.text() or self._folder_input.text() or ""
        if sys.platform == "darwin":
            folder = self._mac_pick_folder("选择输出文件夹", last)
        else:
            folder = QFileDialog.getExistingDirectory(self, "选择输出文件夹", last)
        if folder:
            self._output_input.setText(folder)

    def _mac_pick_folder(self, prompt: str, default: str) -> str:
        """macOS 使用 osascript 调用原生 Finder 选择器。"""
        try:
            cmd = f'choose folder with prompt "{prompt}"'
            if default and os.path.isdir(default):
                cmd += f' default location POSIX file "{default}"'
            r = subprocess.run(
                ["osascript", "-e", f'POSIX path of ({cmd})'],
                capture_output=True, text=True, timeout=120,
            )
            if r.returncode == 0 and r.stdout.strip():
                return r.stdout.strip()
        except Exception:
            pass
        return QFileDialog.getExistingDirectory(self, prompt, default)

    # ------------------------------------------------------------------
    # 按钮状态
    # ------------------------------------------------------------------
    def _update_start_btn(self):
        can_start = bool(self._backends and self._ppt_files)
        self._start_btn.setEnabled(can_start)

    # ------------------------------------------------------------------
    # 转换流程
    # ------------------------------------------------------------------
    def _start_convert(self):
        output_dir = self._output_input.text()
        if not output_dir:
            output_dir = os.path.join(self._folder_input.text(), "导出图片")
            self._output_input.setText(output_dir)

        os.makedirs(output_dir, exist_ok=True)

        # 清理旧状态
        self._result_card.hide()
        self._log_card.show()
        self._log_text.clear()
        self._progress.show()
        self._progress.setValue(0)
        self._progress.setMaximum(len(self._ppt_files))
        self._status_label.show()
        self._start_btn.setEnabled(False)
        self._cancel_btn.show()

        # 日志：显示当前引擎配置
        names = [backend_display_name(b) for b in self._backends]
        self._log_text.append(
            f"引擎优先级: {' → '.join(names)}（失败自动降级）"
        )

        self._worker = ConvertWorker(
            ppt_files=self._ppt_files,
            output_dir=output_dir,
            max_slides=self._pages_spin.value(),
            backends=self._backends,
            soffice_path=self._soffice_path,
        )
        self._worker.progress.connect(self._on_progress)
        self._worker.log.connect(self._on_log)
        self._worker.finished_all.connect(self._on_finished)
        self._worker.start()

    def _cancel_convert(self):
        if self._worker:
            self._worker.abort()
            self._cancel_btn.setEnabled(False)
            self._cancel_btn.setText("正在取消...")

    def _on_progress(self, done: int, total: int, name: str):
        self._progress.setValue(done)
        self._status_label.setText(f"正在处理: {name}")

    def _on_log(self, msg: str):
        self._log_text.append(msg)

    def _on_finished(self, results: List[ConvertResult]):
        self._progress.hide()
        self._status_label.hide()
        self._cancel_btn.hide()
        self._cancel_btn.setEnabled(True)
        self._cancel_btn.setText("取消")
        self._start_btn.setEnabled(True)
        self._worker = None

        # 统计
        total = len(results)
        success = sum(1 for r in results if r.success)
        fail = total - success
        total_pages = sum(r.pages_exported for r in results)

        summary = (
            f"共处理 <b>{total}</b> 个 PPT，"
            f"<span style='color:{_GREEN}'>成功 {success} 个</span>"
        )
        if fail > 0:
            summary += f"，<span style='color:{_RED}'>失败 {fail} 个</span>"
        summary += f"<br>共导出 {total_pages} 页图片"
        self._result_summary.setText(summary)

        # 失败详情（包含引擎信息）
        if fail > 0:
            lines = []
            for r in results:
                if not r.success:
                    lines.append(f"  {r.name}: {r.error}")
            self._result_detail.setText("\n".join(lines))
            self._result_detail.show()
        else:
            self._result_detail.hide()

        self._result_card.show()

    def _open_output_dir(self):
        output_dir = self._output_input.text()
        if not output_dir or not os.path.isdir(output_dir):
            return
        if sys.platform == "darwin":
            subprocess.Popen(["open", output_dir])
        elif sys.platform == "win32":
            os.startfile(output_dir)
        else:
            subprocess.Popen(["xdg-open", output_dir])
