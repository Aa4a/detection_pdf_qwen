# -*- coding: utf-8 -*-
"""
仅刀版线识别, 结果保存为 JSON, 供后续排版使用 (如 /ai_layout、example_ai_layout.py)。

流程: (可选清理 Illustrator 私有数据) -> 检测刀版线 -> 写入 layout_result/*.json
不执行去底色、不排版。
"""

import json
import os
import sys
import traceback
from datetime import datetime

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from qwen_adobe_detection import visualize_dielines
from pdf_strip_illustrator import remove_illustrator_private_data

# 输入 PDF 目录
INPUT_DIR = r"C:\Users\Administrator\Desktop\detection_pdf_purecode\layout_result"

# JSON 输出目录
OUTPUT_DIR = r"C:\Users\Administrator\Desktop\detection_pdf_purecode\layout_result\json"

# 可视化 PDF 子目录 (检测过程会生成, 可选保留)
VIS_DIR = os.path.join(OUTPUT_DIR, "_vis")
CLEANED_DIR = os.path.join(OUTPUT_DIR, "_cleaned")

DEFAULT_LAYOUT_INSTRUCTION = {
    "正唛内容": [1, 3],
    "侧唛内容": [2, 4],
    "正唛上摇盖": [1, 3],
    "侧唛上摇盖": [2, 4],
    "正唛下摇盖": [1, 3],
    "侧唛下摇盖": [2, 4],
}


def _dieline_to_box_size_mm(dieline_size_mm):
    """把检测出的刀版尺寸转成排版用的 box_size_mm 格式。"""
    if not dieline_size_mm:
        return None
    return {
        "排版尺寸": [
            dieline_size_mm.get("w_main_mm", 0.0),
            dieline_size_mm.get("w_side_mm", 0.0),
            dieline_size_mm.get("h_panel_mm", 0.0),
            dieline_size_mm.get("h_flap_mm", 0.0),
        ]
    }


def build_layout_record(source_pdf, det_result, cleaned_pdf=None):
    """组装后续排版可直接使用的 JSON 结构。"""
    dieline = det_result.get("dieline_size_mm") or {}
    success = bool(det_result.get("success"))
    final_boxes = det_result.get("final_boxes") or []
    valid_count = sum(1 for d in final_boxes if d.get("box"))

    record = {
        "success": success,
        "source_pdf": os.path.abspath(source_pdf),
        "cleaned_pdf": os.path.abspath(cleaned_pdf) if cleaned_pdf else None,
        "detected_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "final_boxes": final_boxes,
        "dieline_size_mm": dieline,
        "box_size_mm": _dieline_to_box_size_mm(dieline),
        "layout_instruction": DEFAULT_LAYOUT_INSTRUCTION,
        "rows_skipped": det_result.get("rows_skipped", []),
        "valid_panel_count": valid_count,
        "error": det_result.get("error"),
        "vis_pdf": det_result.get("output_path"),
    }
    return record


def detect_one(
    input_pdf,
    output_dir=OUTPUT_DIR,
    vis_dir=VIS_DIR,
    cleaned_dir=CLEANED_DIR,
    strip_illustrator=True,
    use_qwen=True,
    qwen_requirements="",
):
    """
    对单个 PDF 做刀版线识别, 返回 record dict 并写入 JSON。
    """
    input_pdf = os.path.abspath(input_pdf)
    if not os.path.exists(input_pdf):
        raise FileNotFoundError(f"输入 PDF 不存在: {input_pdf}")

    os.makedirs(output_dir, exist_ok=True)
    os.makedirs(vis_dir, exist_ok=True)

    base = os.path.splitext(os.path.basename(input_pdf))[0]
    json_path = os.path.join(output_dir, f"{base}.json")
    vis_pdf = os.path.join(vis_dir, f"{base}_vis.pdf")

    pdf_for_detect = input_pdf
    cleaned_pdf_path = None

    if strip_illustrator:
        os.makedirs(cleaned_dir, exist_ok=True)
        cleaned_pdf_path = os.path.join(cleaned_dir, f"{base}_clean.pdf")
        try:
            remove_illustrator_private_data(input_pdf, cleaned_pdf_path)
            if os.path.exists(cleaned_pdf_path):
                pdf_for_detect = cleaned_pdf_path
        except Exception as e:
            print(f"[Detect] 警告: 规范化失败 ({e}), 使用原 PDF。")

    print("=" * 60)
    print(f"[Detect] 刀版线识别: {pdf_for_detect}")
    print("=" * 60)

    det_result = visualize_dielines(
        input_path=pdf_for_detect,
        output_path=vis_pdf,
        use_qwen=use_qwen,
        qwen_requirements=qwen_requirements,
    )

    record = build_layout_record(input_pdf, det_result, cleaned_pdf=cleaned_pdf_path)
    record["json_path"] = json_path

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(record, f, ensure_ascii=False, indent=2)

    print(f"[Detect] JSON 已保存: {json_path}")
    return record


if __name__ == "__main__":
    input_path = INPUT_DIR

    if os.path.isfile(input_path):
        if not input_path.lower().endswith(".pdf"):
            print(f"[X] 输入文件不是 PDF: {input_path}")
            sys.exit(1)
        pdf_paths = [input_path]
    elif os.path.isdir(input_path):
        pdf_paths = sorted([
            os.path.join(input_path, f)
            for f in os.listdir(input_path)
            if f.lower().endswith(".pdf")
        ])
    else:
        print(f"[X] 输入路径不存在: {input_path}")
        sys.exit(1)

    print(f"共发现 {len(pdf_paths)} 个 PDF 待识别。")
    print(f"JSON 输出目录: {OUTPUT_DIR}")

    ok, fail, skipped = [], [], []
    for i, in_path in enumerate(pdf_paths, 1):
        name = os.path.basename(in_path)
        base = os.path.splitext(name)[0]
        json_path = os.path.join(OUTPUT_DIR, f"{base}.json")
        if os.path.exists(json_path):
            print(f"\n########## [{i}/{len(pdf_paths)}] {name} (已存在 JSON, 跳过) ##########")
            skipped.append(name)
            continue
        print(f"\n########## [{i}/{len(pdf_paths)}] {name} ##########")
        try:
            record = detect_one(in_path)
            if record.get("success"):
                ok.append((name, record["json_path"]))
            else:
                fail.append((name, record.get("error") or "检测失败"))
        except Exception as e:
            traceback.print_exc()
            fail.append((name, str(e)))

    summary = {
        "total": len(pdf_paths),
        "skipped": len(skipped),
        "success": len(ok),
        "fail": len(fail),
        "ok": [{"name": n, "json": p} for n, p in ok],
        "fail_list": [{"name": n, "error": err} for n, err in fail],
        "finished_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    summary_path = os.path.join(OUTPUT_DIR, "_summary.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    print("\n" + "=" * 60)
    print(f"识别完成: 跳过 {len(skipped)} 个, 成功 {len(ok)} 个, 失败 {len(fail)} 个")
    print(f"汇总: {summary_path}")
    if ok:
        print("成功列表:")
        for n, p in ok:
            print(f"  + {n} -> {p}")
    if fail:
        print("失败列表:")
        for n, err in fail:
            print(f"  - {n}: {err}")
