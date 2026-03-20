"""PPT → PNG 转换引擎。

支持三种后端（按优先级自动选择）：
  A. WPS COM (仅 Windows，用户已安装 WPS)
  B. PowerPoint COM (仅 Windows) / PowerPoint AppleScript (仅 macOS)
  C. LibreOffice headless (跨平台备选)
"""

import os
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass, field
from typing import List, Optional

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


# ---------------------------------------------------------------------------
# 后端检测
# ---------------------------------------------------------------------------

def _detect_wps_com() -> bool:
    """检测 Windows 上是否可用 WPS COM。"""
    if sys.platform != "win32":
        return False
    try:
        import comtypes.client  # noqa: F401
        app = comtypes.client.CreateObject("KWPP.Application")
        app.Quit()
        return True
    except Exception:
        return False


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
    try:
        r = subprocess.run(
            ["osascript", "-e",
             'tell application "System Events" to (name of processes) contains "Microsoft PowerPoint"'],
            capture_output=True, text=True, timeout=5,
        )
        # 即使 PowerPoint 没运行，只要安装了就能用
        # 检查应用是否存在
        r2 = subprocess.run(
            ["mdfind", "kMDItemCFBundleIdentifier == 'com.microsoft.Powerpoint'"],
            capture_output=True, text=True, timeout=10,
        )
        return bool(r2.stdout.strip())
    except Exception:
        return False


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

BACKEND_WPS_COM = "wps_com"
BACKEND_PPT_COM = "ppt_com"
BACKEND_PPT_MAC = "ppt_mac"
BACKEND_LIBREOFFICE = "libreoffice"


def detect_backends() -> List[str]:
    """返回当前系统可用的后端列表（按优先级排序）。"""
    available = []
    if sys.platform == "win32":
        if _detect_wps_com():
            available.append(BACKEND_WPS_COM)
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
        BACKEND_WPS_COM: "WPS (COM)",
        BACKEND_PPT_COM: "PowerPoint (COM)",
        BACKEND_PPT_MAC: "PowerPoint (AppleScript)",
        BACKEND_LIBREOFFICE: "LibreOffice",
    }.get(backend, backend)


# ---------------------------------------------------------------------------
# 转换实现
# ---------------------------------------------------------------------------

def _convert_wps_com(filepath: str, out_dir: str, max_slides: int) -> int:
    """使用 WPS COM 导出幻灯片为 PNG。返回导出页数。"""
    import comtypes.client
    wpp = comtypes.client.CreateObject("KWPP.Application")
    wpp.Visible = False
    try:
        pres = wpp.Presentations.Open(os.path.abspath(filepath), WithWindow=False)
        try:
            total = pres.Slides.Count
            n = min(total, max_slides)
            for i in range(1, n + 1):
                slide = pres.Slides.Item(i)
                out_path = os.path.join(out_dir, f"{i}.png")
                slide.Export(out_path, "PNG", 1920, 1080)
            return n
        finally:
            pres.Close()
    finally:
        wpp.Quit()


def _convert_ppt_com(filepath: str, out_dir: str, max_slides: int) -> int:
    """使用 PowerPoint COM 导出幻灯片为 PNG。返回导出页数。"""
    import comtypes.client
    ppt = comtypes.client.CreateObject("PowerPoint.Application")
    try:
        pres = ppt.Presentations.Open(
            os.path.abspath(filepath), ReadOnly=True, WithWindow=False
        )
        try:
            total = pres.Slides.Count
            n = min(total, max_slides)
            for i in range(1, n + 1):
                slide = pres.Slides.Item(i)
                out_path = os.path.join(out_dir, f"{i}.png")
                slide.Export(out_path, "PNG", 1920, 1080)
            return n
        finally:
            pres.Close()
    finally:
        ppt.Quit()


def _convert_ppt_mac(filepath: str, out_dir: str, max_slides: int) -> int:
    """使用 macOS PowerPoint AppleScript 导出幻灯片为 PNG。返回导出页数。"""
    abs_path = os.path.abspath(filepath)
    # AppleScript 需要 POSIX 路径
    script = f'''
    tell application "Microsoft PowerPoint"
        open POSIX file "{abs_path}"
        set pres to active presentation
        set slideCount to count of slides of pres
        set maxN to {max_slides}
        if slideCount < maxN then
            set maxN to slideCount
        end if
        repeat with i from 1 to maxN
            set outPath to POSIX path of "{out_dir}/" & i & ".png"
            save slide i of pres in outPath as save as PNG
        end repeat
        close pres saving no
    end tell
    '''
    r = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True, text=True, timeout=300,
    )
    if r.returncode != 0:
        raise RuntimeError(f"AppleScript 错误: {r.stderr.strip()}")
    # 计算实际导出了多少页
    exported = len([f for f in os.listdir(out_dir) if f.endswith(".png")])
    return exported


def _convert_libreoffice(filepath: str, out_dir: str, max_slides: int,
                         soffice_path: Optional[str] = None) -> int:
    """使用 LibreOffice headless 转换 PPT → PDF → PNG。返回导出页数。"""
    import fitz  # PyMuPDF

    if not soffice_path:
        soffice_path = _find_libreoffice()
    if not soffice_path:
        raise RuntimeError("未找到 LibreOffice，请先安装")

    with tempfile.TemporaryDirectory(prefix="ppt2img_") as tmpdir:
        # 使用独立的用户配置避免锁冲突
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
            subprocess.run(cmd, capture_output=True, timeout=120, check=True)
        finally:
            shutil.rmtree(profile_dir, ignore_errors=True)

        # 找到生成的 PDF
        pdfs = [f for f in os.listdir(tmpdir) if f.lower().endswith(".pdf")]
        if not pdfs:
            raise RuntimeError("LibreOffice 转换失败：未生成 PDF 文件")

        pdf_path = os.path.join(tmpdir, pdfs[0])
        doc = fitz.open(pdf_path)
        try:
            total = len(doc)
            n = min(total, max_slides)
            for i in range(n):
                page = doc[i]
                # 2x 缩放 → 144 DPI，适合大多数用途
                pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                pixmap.save(os.path.join(out_dir, f"{i + 1}.png"))
            return n
        finally:
            doc.close()


# ---------------------------------------------------------------------------
# 统一转换入口
# ---------------------------------------------------------------------------

def convert_one(filepath: str, out_dir: str, max_slides: int,
                backend: str, soffice_path: Optional[str] = None) -> int:
    """使用指定后端转换一个 PPT 文件。返回导出页数。"""
    os.makedirs(out_dir, exist_ok=True)
    if backend == BACKEND_WPS_COM:
        return _convert_wps_com(filepath, out_dir, max_slides)
    elif backend == BACKEND_PPT_COM:
        return _convert_ppt_com(filepath, out_dir, max_slides)
    elif backend == BACKEND_PPT_MAC:
        return _convert_ppt_mac(filepath, out_dir, max_slides)
    elif backend == BACKEND_LIBREOFFICE:
        return _convert_libreoffice(filepath, out_dir, max_slides, soffice_path)
    else:
        raise ValueError(f"未知后端: {backend}")


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
# QThread Worker
# ---------------------------------------------------------------------------

class ConvertWorker(QThread):
    """后台批量转换线程。"""

    # 信号：(完成数, 总数, 当前文件名)
    progress = pyqtSignal(int, int, str)
    # 信号：实时日志
    log = pyqtSignal(str)
    # 信号：全部完成，传递结果列表
    finished_all = pyqtSignal(list)

    def __init__(self, ppt_files: List[str], output_dir: str,
                 max_slides: int, backend: str,
                 soffice_path: Optional[str] = None):
        super().__init__()
        self.ppt_files = ppt_files
        self.output_dir = output_dir
        self.max_slides = max_slides
        self.backend = backend
        self.soffice_path = soffice_path
        self._abort = False

    def abort(self):
        self._abort = True

    def run(self):
        results: List[ConvertResult] = []
        total = len(self.ppt_files)

        for idx, filepath in enumerate(self.ppt_files):
            if self._abort:
                self.log.emit("已取消转换")
                break

            basename = os.path.splitext(os.path.basename(filepath))[0]
            cleaned = clean_filename(basename)
            self.progress.emit(idx, total, os.path.basename(filepath))
            self.log.emit(f"[{idx + 1}/{total}] 正在转换: {os.path.basename(filepath)}")

            result = ConvertResult(filepath=filepath, name=os.path.basename(filepath))
            try:
                out_dir = _unique_dir(os.path.join(self.output_dir, cleaned))
                pages = convert_one(
                    filepath, out_dir, self.max_slides,
                    self.backend, self.soffice_path,
                )
                result.success = True
                result.pages_exported = pages
                result.output_dir = out_dir
                self.log.emit(f"  -> 成功，导出 {pages} 页到 {os.path.basename(out_dir)}/")
            except Exception as e:
                result.error = str(e)
                self.log.emit(f"  -> 失败: {e}")

            results.append(result)

        self.progress.emit(total, total, "完成")
        self.finished_all.emit(results)
