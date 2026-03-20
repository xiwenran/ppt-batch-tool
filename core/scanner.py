"""递归扫描文件夹中的 PPT 文件。"""

import os
from typing import List

PPT_EXTENSIONS = {".ppt", ".pptx", ".pps", ".ppsx"}


def scan_ppt_files(folder: str) -> List[str]:
    """递归扫描 folder 下所有 PPT 文件，跳过临时文件，返回排序后的绝对路径列表。"""
    results: List[str] = []
    for dirpath, _dirnames, filenames in os.walk(folder):
        for fname in filenames:
            # 跳过 Office 临时文件
            if fname.startswith("~$"):
                continue
            _, ext = os.path.splitext(fname)
            if ext.lower() in PPT_EXTENSIONS:
                results.append(os.path.join(dirpath, fname))
    results.sort()
    return results
