"""PPT → PNG 转换引擎。

支持后端（按优先级自动选择，运行时自动降级）：
  A. PowerPoint COM (仅 Windows) / PowerPoint AppleScript (仅 macOS)
  B. LibreOffice headless (跨平台备选)
"""

import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from typing import List, Optional, Callable

from PyQt6.QtCore import QThread, pyqtSignal

from core.filename_cleaner import clean_filename


# ---------------------------------------------------------------------------
# 数据结构
# ---------------------------------------------------------------------------

@dataclass
class ConvertResult:
    """单个 PPT 的转换结果。"""
    filepath: str
    name: str
    success: bool = False
    pages_exported: int = 0
    error: str = ""
    output_dir: str = ""
    backend_used: str = ""  # 实际使用的后端


# ---------------------------------------------------------------------------
# 后端检测
# ---------------------------------------------------------------------------

def _detect_powerpoint_com() -> bool:
    """检测 Windows 上是否可用 PowerPoint COM。"""
    if sys.platform != "win32":
        return False
    try:
        import comtypes.client  # noqa: F401
        app = comtypes.client.CreateObject("PowerPoint.Application")
        app.Quit()
        return True
    except Exception:
        return False


def _detect_powerpoint_mac() -> bool:
    """检测 macOS 上是否安装了 Microsoft PowerPoint。"""
    if sys.platform != "darwin":
        return False
    return os.path.isdir("/Applications/Microsoft PowerPoint.app")


def _find_libreoffice() -> Optional[str]:
    """查找 LibreOffice soffice 可执行文件路径。"""
    candidates = []
    if sys.platform == "darwin":
        candidates = ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    elif sys.platform == "win32":
        for pf in [os.environ.get("ProgramFiles", r"C:\Program Files"),
                    os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")]:
            if pf:
                candidates.append(os.path.join(pf, "LibreOffice", "program", "soffice.exe"))
    # 兜底：PATH 上的 soffice / libreoffice
    for cmd in ("soffice", "libreoffice"):
        p = shutil.which(cmd)
        if p:
            candidates.append(p)
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


# ---------------------------------------------------------------------------
# 后端枚举
# ---------------------------------------------------------------------------

BACKEND_PPT_COM = "ppt_com"
BACKEND_PPT_MAC = "ppt_mac"
BACKEND_LIBREOFFICE = "libreoffice"


def detect_backends() -> List[str]:
    """返回当前系统可用的后端列表（按优先级排序）。"""
    available = []
    if sys.platform == "win32":
        if _detect_powerpoint_com():
            available.append(BACKEND_PPT_COM)
    elif sys.platform == "darwin":
        if _detect_powerpoint_mac():
            available.append(BACKEND_PPT_MAC)
    if _find_libreoffice():
        available.append(BACKEND_LIBREOFFICE)
    return available


def backend_display_name(backend: str) -> str:
    return {
        BACKEND_PPT_COM: "PowerPoint (COM)",
        BACKEND_PPT_MAC: "PowerPoint (AppleScript)",
        BACKEND_LIBREOFFICE: "LibreOffice",
    }.get(backend, backend)


# ---------------------------------------------------------------------------
# 转换实现
# ---------------------------------------------------------------------------

def _convert_ppt_com(filepath: str, out_dir: str, max_slides: int,
                     log: Optional[Callable] = None) -> int:
    """使用 PowerPoint COM 导出幻灯片为 PNG（PDF 中转方案）。返回导出页数。"""
    import comtypes.client

    abs_path = os.path.abspath(filepath).replace("/", "\\")
    if log:
        log(f"    [PowerPoint] 正在创建 PowerPoint 实例...")
    ppt = comtypes.client.CreateObject("PowerPoint.Application")
    try:
        if log:
            log(f"    [PowerPoint] 正在打开文件: {os.path.basename(filepath)}")
        # ReadOnly=-1, Untitled=0, WithWindow=0
        pres = ppt.Presentations.Open(abs_path, -1, 0, 0)
        try:
            pdf_dir = os.path.join(os.path.dirname(out_dir), ".ppt2img_tmp")
            os.makedirs(pdf_dir, exist_ok=True)
            pdf_path = os.path.join(pdf_dir, "output.pdf").replace("/", "\\")
            try:
                if log:
                    log(f"    [PowerPoint] 正在导出 PDF...")
                pres.ExportAsFixedFormat(pdf_path, 2)
                if log:
                    log(f"    [PowerPoint] PDF 导出成功，正在渲染 PNG...")
                pres.Close()
                return _pdf_to_png(pdf_path, out_dir, max_slides)
            finally:
                shutil.rmtree(pdf_dir, ignore_errors=True)
        except Exception:
            try:
                pres.Close()
            except Exception:
                pass
            raise
    finally:
        try:
            ppt.Quit()
        except Exception:
            pass


def _ppt_mac_batch_export_pdf(ppt_files: List[str], pdf_dir: str,
                              log: Optional[Callable] = None) -> dict:
    """用一次 AppleScript 调用批量将所有 PPT 导出为 PDF。

    PowerPoint 只启动/授权一次，大幅减少权限弹窗。
    返回 {ppt绝对路径: pdf绝对路径} 映射，失败的文件不在字典中。
    """
    os.makedirs(pdf_dir, exist_ok=True)
    file_commands = []
    path_map = {}
    for i, ppt_path in enumerate(ppt_files):
        abs_path = os.path.abspath(ppt_path)
        pdf_path = os.path.join(pdf_dir, f"{i}.pdf")
        path_map[abs_path] = pdf_path
        file_commands.append(f'''
        try
            open POSIX file "{abs_path}"
            save active presentation in POSIX file "{pdf_path}" as save as PDF
            close active presentation saving no
        end try''')

    script = f'''
tell application "Microsoft PowerPoint"
    {"".join(file_commands)}
end tell
'''
    if log:
        log(f"    [PowerPoint macOS] 正在批量导出 {len(ppt_files)} 个 PDF...")
    subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True,
        timeout=len(ppt_files) * 120 + 60,
    )
    result = {k: v for k, v in path_map.items() if os.path.isfile(v)}
    if log:
        log(f"    [PowerPoint macOS] PDF 导出完成：{len(result)}/{len(ppt_files)} 个成功")
    return result


def _pdf_to_png(pdf_path: str, out_dir: str, max_slides: int) -> int:
    """用 PyMuPDF 将 PDF 每页渲染为 PNG。返回导出页数。"""
    import fitz  # PyMuPDF

    os.makedirs(out_dir, exist_ok=True)
    doc = fitz.open(pdf_path)
    try:
        total = len(doc)
        n = min(total, max_slides)
        for i in range(n):
            page = doc[i]
            pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            pixmap.save(os.path.join(out_dir, f"{i + 1}.png"))
        return n
    finally:
        doc.close()


def _convert_libreoffice(filepath: str, out_dir: str, max_slides: int,
                         soffice_path: Optional[str] = None,
                         log: Optional[Callable] = None) -> int:
    """使用 LibreOffice headless 转换 PPT → PDF → PNG。返回导出页数。"""
    import fitz  # PyMuPDF

    if not soffice_path:
        soffice_path = _find_libreoffice()
    if not soffice_path:
        raise RuntimeError("未找到 LibreOffice，请先安装")

    with tempfile.TemporaryDirectory(prefix="ppt2img_") as tmpdir:
        profile_dir = tempfile.mkdtemp(prefix="ppt2img_profile_")
        try:
            profile_url = "file://" + profile_dir.replace("\\", "/")
            cmd = [
                soffice_path,
                "--headless",
                "--norestore",
                f"--env:UserInstallation={profile_url}",
                "--convert-to", "pdf",
                "--outdir", tmpdir,
                os.path.abspath(filepath),
            ]
            if log:
                log(f"    [LibreOffice] 正在转换为 PDF...")
            subprocess.run(cmd, capture_output=True, timeout=120, check=True)
        finally:
            shutil.rmtree(profile_dir, ignore_errors=True)

        pdfs = [f for f in os.listdir(tmpdir) if f.lower().endswith(".pdf")]
        if not pdfs:
            raise RuntimeError("LibreOffice 转换失败：未生成 PDF 文件")

        pdf_path = os.path.join(tmpdir, pdfs[0])
        if log:
            log(f"    [LibreOffice] PDF 转换成功，正在渲染 PNG...")
        doc = fitz.open(pdf_path)
        try:
            total = len(doc)
            n = min(total, max_slides)
            for i in range(n):
                page = doc[i]
                pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                pixmap.save(os.path.join(out_dir, f"{i + 1}.png"))
            return n
        finally:
            doc.close()


# ---------------------------------------------------------------------------
# 统一转换入口（支持 fallback）
# ---------------------------------------------------------------------------

def convert_one_with_fallback(filepath: str, out_dir: str, max_slides: int,
                              backends: List[str],
                              soffice_path: Optional[str] = None,
                              pdf_path: Optional[str] = None,
                              log: Optional[Callable] = None) -> tuple:
    """依次尝试所有可用后端，返回 (导出页数, 实际使用的后端名)。

    如果 pdf_path 已提供（macOS 批量模式预导出的 PDF），直接渲染 PNG。
    """
    os.makedirs(out_dir, exist_ok=True)

    # macOS 预导出 PDF 模式
    if pdf_path:
        pages = _pdf_to_png(pdf_path, out_dir, max_slides)
        return pages, BACKEND_PPT_MAC

    last_error = None
    for backend in backends:
        try:
            if backend == BACKEND_PPT_COM:
                pages = _convert_ppt_com(filepath, out_dir, max_slides, log)
            elif backend == BACKEND_LIBREOFFICE:
                pages = _convert_libreoffice(filepath, out_dir, max_slides,
                                              soffice_path, log)
            else:
                continue
            return pages, backend
        except Exception as e:
            last_error = e
            name = backend_display_name(backend)
            if log:
                log(f"    ⚠ {name} 转换失败: {e}")
                if len(backends) > 1:
                    log(f"    → 正在尝试下一个引擎...")
            # 清理可能创建的空输出目录
            if os.path.isdir(out_dir) and not os.listdir(out_dir):
                shutil.rmtree(out_dir, ignore_errors=True)
            continue

    # 所有后端都失败
    if last_error:
        raise RuntimeError(
            f"所有转换引擎均失败。最后错误: {last_error}\n"
            f"建议：尝试手动切换引擎，或安装 LibreOffice 作为备选。"
        )
    raise RuntimeError("没有可用的转换引擎")


# ---------------------------------------------------------------------------
# 输出目录命名（去重）
# ---------------------------------------------------------------------------

def _unique_dir(base: str) -> str:
    """如果目录已存在，追加 _2, _3... 返回不冲突的路径。"""
    if not os.path.exists(base):
        return base
    i = 2
    while os.path.exists(f"{base}_{i}"):
        i += 1
    return f"{base}_{i}"


# ---------------------------------------------------------------------------
# QThread Worker（支持 fallback）
# ---------------------------------------------------------------------------

class ConvertWorker(QThread):
    """后台批量转换线程，支持引擎自动降级。"""

    # 信号：(完成数, 总数, 当前文件名)
    progress = pyqtSignal(int, int, str)
    # 信号：实时日志
    log = pyqtSignal(str)
    # 信号：全部完成，传递结果列表
    finished_all = pyqtSignal(list)

    def __init__(self, ppt_files: List[str], output_dir: str,
                 max_slides: int, backends: List[str],
                 soffice_path: Optional[str] = None):
        super().__init__()
        self.ppt_files = ppt_files
        self.output_dir = output_dir
        self.max_slides = max_slides
        self.backends = backends  # 所有可用后端列表（用于 fallback）
        self.soffice_path = soffice_path
        self._abort = False

    def abort(self):
        self._abort = True

    def _log(self, msg: str):
        """线程安全的日志发射。"""
        self.log.emit(msg)

    def run(self):
        results: List[ConvertResult] = []
        total = len(self.ppt_files)

        # macOS PowerPoint 批量模式
        pdf_map: dict = {}
        tmp_pdf_dir: Optional[str] = None
        primary_backend = self.backends[0] if self.backends else None

        if primary_backend == BACKEND_PPT_MAC:
            self._log("正在通过 PowerPoint 批量导出 PDF（仅需授权一次）...")
            self.progress.emit(0, total, "批量导出 PDF 中...")
            # 在输出目录下创建临时 PDF 目录（避免授权系统临时目录）
            tmp_pdf_dir = os.path.join(self.output_dir, ".ppt2img_pdf_tmp")
            try:
                pdf_map = _ppt_mac_batch_export_pdf(
                    self.ppt_files, tmp_pdf_dir, self._log
                )
            except Exception as e:
                self._log(f"PowerPoint 批量导出失败: {e}")
            self._log(f"PDF 导出完成：{len(pdf_map)}/{total} 个成功")

        try:
            for idx, filepath in enumerate(self.ppt_files):
                if self._abort:
                    self._log("已取消转换")
                    break

                basename = os.path.splitext(os.path.basename(filepath))[0]
                cleaned = clean_filename(basename)
                self.progress.emit(idx, total, os.path.basename(filepath))
                self._log(f"[{idx + 1}/{total}] 正在处理: {os.path.basename(filepath)}")

                result = ConvertResult(
                    filepath=filepath, name=os.path.basename(filepath)
                )
                abs_path = os.path.abspath(filepath)
                try:
                    out_dir = _unique_dir(
                        os.path.join(self.output_dir, cleaned)
                    )

                    if primary_backend == BACKEND_PPT_MAC:
                        pdf_path = pdf_map.get(abs_path)
                        if not pdf_path:
                            raise RuntimeError("PowerPoint 导出 PDF 失败")
                        pages, used = convert_one_with_fallback(
                            filepath, out_dir, self.max_slides,
                            self.backends, pdf_path=pdf_path, log=self._log,
                        )
                    else:
                        pages, used = convert_one_with_fallback(
                            filepath, out_dir, self.max_slides,
                            self.backends, self.soffice_path,
                            log=self._log,
                        )

                    result.success = True
                    result.pages_exported = pages
                    result.output_dir = out_dir
                    result.backend_used = used
                    used_name = backend_display_name(used)
                    self._log(
                        f"  ✓ 成功（{used_name}），"
                        f"导出 {pages} 页到 {os.path.basename(out_dir)}/"
                    )
                except Exception as e:
                    result.error = str(e)
                    self._log(f"  ✗ 失败: {e}")

                results.append(result)
        finally:
            if tmp_pdf_dir:
                shutil.rmtree(tmp_pdf_dir, ignore_errors=True)

        self.progress.emit(total, total, "完成")
        self.finished_all.emit(results)
