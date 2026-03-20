"""清理 PPT 文件名：去除版权声明、平台标记、非法字符等。"""

import hashlib
import re

# 按顺序应用的清理规则（顺序很重要）
_STRIP_PATTERNS = [
    # 先去括号内容（可能包含平台/版权信息）
    r"【[^】]*】",
    r"\[[^\]]*\]",
    r"（[^）]*）",
    r"\([^)]*\)",
    # @用户名（含中文用户名）
    r"@[\w\u4e00-\u9fff]+",
    # 平台标记（公众号：XXX、小红书XXX 等）
    r"公众号[：:·]?\s*[\w\u4e00-\u9fff]*",
    r"小红书[：:·]?\s*[\w\u4e00-\u9fff]*",
    r"抖音[：:·]?\s*[\w\u4e00-\u9fff]*",
    r"微博[：:·]?\s*[\w\u4e00-\u9fff]*",
    r"[Bb]站[：:·]?\s*[\w\u4e00-\u9fff]*",
    r"知乎[：:·]?\s*[\w\u4e00-\u9fff]*",
    r"快手[：:·]?\s*[\w\u4e00-\u9fff]*",
    r"微信[：:·]?\s*[\w\u4e00-\u9fff]*",
    r"头条[：:·]?\s*[\w\u4e00-\u9fff]*",
    # 版权/法律声明
    r"如有侵权[\w\u4e00-\u9fff]*",
    r"侵权?删除?",
    r"侵删",
    r"转载[\w\u4e00-\u9fff]*",
    r"版权[\w\u4e00-\u9fff]*",
    r"免责[\w\u4e00-\u9fff]*",
    r"仅供[\w\u4e00-\u9fff]*学习[\w\u4e00-\u9fff]*",
    r"禁止[\w\u4e00-\u9fff]*商用[\w\u4e00-\u9fff]*",
    r"来源[：:][\w\u4e00-\u9fff]*",
    r"作者[：:][\w\u4e00-\u9fff]*",
    r"出处[：:][\w\u4e00-\u9fff]*",
    # "同名" 常见于 "抖音同名"
    r"同名",
]

# Windows / macOS 文件名非法字符
_ILLEGAL_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

# 连续空白/分隔符
_MULTI_SEP = re.compile(r"[\s_\-—]+")


def clean_filename(original: str) -> str:
    """
    清理文件名（不含扩展名），返回适合用作文件夹名的字符串。
    如果清理后为空，使用原始文件名的 hash 前 8 位。
    """
    name = original

    for pattern in _STRIP_PATTERNS:
        name = re.sub(pattern, " ", name)

    # 去除非法字符
    name = _ILLEGAL_CHARS.sub("", name)

    # 合并连续空白/分隔符为单个空格
    name = _MULTI_SEP.sub(" ", name).strip()

    # 去掉首尾的标点和分隔符
    name = name.strip(" _-—.·,，、;；!！~")

    # 限制长度
    if len(name) > 80:
        name = name[:80].rstrip()

    # 兜底：空字符串时用 hash
    if not name:
        name = hashlib.md5(original.encode("utf-8")).hexdigest()[:8]

    return name
