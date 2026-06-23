#!/usr/bin/env python3
"""
PPT / Word 转图片 CLI — 命令行接口，供 Claude Code Skill 调用。
不依赖 PyQt6，直接调用 core/ 的转换函数。

用法：
  python3 cli.py detect                               # 检测可用转换引擎
  python3 cli.py convert --input <文件夹> --output <输出目录> [--max-slides N]
"""

import argparse
import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(__file__))


def cmd_detect():
    from core.converter import detect_backends, backend_display_name, _find_libreoffice
    backends = detect_backends()
    if backends:
        print(f"可用引擎（按优先级）：")
        for b in backends:
            print(f"  ✅ {backend_display_name(b)}")
    else:
        print("❌ 没有找到可用引擎（需要安装 Microsoft Office 或 LibreOffice）")
    lo = _find_libreoffice()
    if lo:
        print(f"  LibreOffice 路径：{lo}")


def cmd_convert(input_folder: str, output_dir: str, max_slides: int, only_file: str | None = None):
    from core.scanner import scan_supported_files
    from core.filename_cleaner import clean_filename
    from core.converter import (
        detect_backends, convert_one_with_fallback,
        _ppt_mac_batch_export_pdf, _word_mac_batch_export_pdf,
        _find_libreoffice, BACKEND_PPT_MAC, BACKEND_WORD_MAC,
        backends_for_file, is_ppt_file, is_word_file,
    )

    input_folder = os.path.expanduser(input_folder)
    output_dir = os.path.expanduser(output_dir)

    if not os.path.isdir(input_folder):
        print(f"[错误] 输入文件夹不存在：{input_folder}", file=sys.stderr)
        sys.exit(1)

    backends = detect_backends()
    if not backends:
        print("[错误] 没有可用的转换引擎，请安装 Microsoft Office 或 LibreOffice", file=sys.stderr)
        sys.exit(1)

    source_files = scan_supported_files(input_folder)
    if only_file:
        source_files = [f for f in source_files if os.path.basename(f) == only_file]
        if not source_files:
            print(f"[错误] 在 {input_folder} 中没有找到文件：{only_file}", file=sys.stderr)
            sys.exit(1)
    if not source_files:
        print(f"[错误] 在 {input_folder} 中没有找到受支持的 PPT / Word 文件", file=sys.stderr)
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)
    total = len(source_files)
    print(f"找到 {total} 个 PPT / Word 文件，使用引擎：{backends[0]}")
    print(f"输出目录：{output_dir}\n")

    soffice_path = _find_libreoffice()
    mac_pdf_map: dict[str, tuple[str, str]] = {}
    tmp_pdf_dirs: list[str] = []
    if sys.platform == "darwin":
        for label, backend, exporter, predicate in [
            ("PowerPoint", BACKEND_PPT_MAC, _ppt_mac_batch_export_pdf, is_ppt_file),
            ("Word", BACKEND_WORD_MAC, _word_mac_batch_export_pdf, is_word_file),
        ]:
            eligible = [
                path for path in source_files
                if predicate(path) and backends_for_file(path, backends)[:1] == [backend]
            ]
            if not eligible:
                continue
            tmp_pdf_dir = os.path.join(output_dir, f".ppt2img_{backend}_pdf_tmp")
            tmp_pdf_dirs.append(tmp_pdf_dir)
            print(f"正在通过 {label} 批量导出 PDF（此步骤可能弹出授权窗口，点一次允许即可）...")
            try:
                pdf_map = exporter(eligible, tmp_pdf_dir, log=print)
            except Exception as e:
                print(f"    ⚠ {label} 批量导出失败：{e}")
                print("    → 将回退到逐文件转换，继续尝试其他引擎\n")
                continue
            for source_path, pdf_path in pdf_map.items():
                mac_pdf_map[source_path] = (pdf_path, backend)
            print(f"PDF 导出完成：{len(pdf_map)}/{len(eligible)} 个成功\n")

    success, failed = 0, []
    try:
        for idx, filepath in enumerate(source_files, 1):
            basename = os.path.splitext(os.path.basename(filepath))[0]
            cleaned = clean_filename(basename)
            out_dir = os.path.join(output_dir, cleaned)
            print(f"[{idx}/{total}] {os.path.basename(filepath)} → {cleaned}/")

            try:
                abs_path = os.path.abspath(filepath)
                pdf_entry = mac_pdf_map.get(abs_path)
                if pdf_entry:
                    pdf_path, pdf_backend = pdf_entry
                    pages, used, actual_out_dir = convert_one_with_fallback(
                        filepath, out_dir, max_slides, backends,
                        pdf_path=pdf_path, pdf_backend=pdf_backend, log=print
                    )
                else:
                    pages, used, actual_out_dir = convert_one_with_fallback(
                        filepath, out_dir, max_slides, backends, soffice_path, log=print
                    )
                print(f"       ✅ 导出 {pages} 页到 {os.path.basename(actual_out_dir)}/")
                success += 1
            except Exception as e:
                print(f"       ❌ 失败：{e}")
                failed.append((os.path.basename(filepath), str(e)))
    finally:
        for tmp_pdf_dir in tmp_pdf_dirs:
            if os.path.isdir(tmp_pdf_dir):
                shutil.rmtree(tmp_pdf_dir, ignore_errors=True)

    print(f"\n完成！成功 {success}/{total} 个")
    if failed:
        print(f"失败 {len(failed)} 个：")
        for name, err in failed:
            print(f"  ✗ {name}：{err}")
    print(f"输出目录：{output_dir}")

    # JSON summary — 供 Skill/Agent 读取，只汇报数字和失败文件，不展开所有成功路径
    skipped = total - success - len(failed)
    # 采样最多 3 个成功输出目录作为样本
    sample_outputs = []
    if os.path.isdir(output_dir):
        for entry in sorted(os.listdir(output_dir))[:3]:
            candidate = os.path.join(output_dir, entry)
            if os.path.isdir(candidate):
                sample_outputs.append(candidate)
    summary = {
        "input_dir": input_folder,
        "output_dir": output_dir,
        "success_count": success,
        "failed_count": len(failed),
        "skipped_count": skipped,
        "failed_files": [{"file": name, "error": err} for name, err in failed],
        "sample_outputs": sample_outputs,
    }
    summary_path = os.path.join(output_dir, "convert_summary.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"\n[JSON summary] {summary_path}")


def main():
    parser = argparse.ArgumentParser(description="PPT / Word 转图片命令行工具")
    sub = parser.add_subparsers(dest="cmd")

    sub.add_parser("detect", help="检测可用的转换引擎")

    p = sub.add_parser("convert", help="批量转换 PPT / Word 为 PNG 图片")
    p.add_argument("--input", required=True, help="包含 PPT / Word 文件的文件夹（递归扫描）")
    p.add_argument("--output", required=True, help="图片输出目录")
    p.add_argument("--max-slides", type=int, default=17, help="每个文件最多导出页数（默认 17）")
    p.add_argument("--only-file", default=None, help="只处理指定文件名（单文件模式，由 pipeline 传入）")

    args = parser.parse_args()

    if args.cmd == "detect":
        cmd_detect()
    elif args.cmd == "convert":
        cmd_convert(args.input, args.output, args.max_slides, args.only_file)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
