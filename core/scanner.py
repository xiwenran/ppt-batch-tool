"""递归扫描文件夹中的演示文稿与 Word 文档。"""

import os
from typing import List

PPT_EXTENSIONS = {".ppt", ".pptx", ".pps", ".ppsx"}
WORD_EXTENSIONS = {".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm"}
SUPPORTED_EXTENSIONS = PPT_EXTENSIONS | WORD_EXTENSIONS


def scan_ppt_files(folder: str) -> List[str]:
    """兼容旧接口：递归扫描 folder 下所有 PPT 文件。"""
    return scan_files_by_extensions(folder, PPT_EXTENSIONS)


def scan_supported_files(folder: str) -> List[str]:
    """递归扫描 folder 下所有受支持的 PPT / Word 文件。"""
    return scan_files_by_extensions(folder, SUPPORTED_EXTENSIONS)


def scan_files_by_extensions(folder: str, extensions: set[str]) -> List[str]:
    """递归扫描 folder 下指定扩展名文件，跳过 Office 临时文件。"""
    results: List[str] = []
    for dirpath, _dirnames, filenames in os.walk(folder):
        for fname in filenames:
            if fname.startswith("~$"):
                continue
            _, ext = os.path.splitext(fname)
            if ext.lower() in extensions:
                results.append(os.path.join(dirpath, fname))
    results.sort()
    return results
