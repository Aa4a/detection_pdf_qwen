# -*- coding: utf-8 -*-
"""
从 detect_to_json 产出的 JSON 批量排版，每种 PDF 生成多种随机刀版尺寸:

  type4 - 宽不变，高随机（正侧唛面板高相同）
  type5 - 高不变，宽随机
  type6 - 宽高都随机

宽/高随机基准在 TARGET_WIDTH_MM / TARGET_HEIGHT_MM ± JITTER 附近 (默认约 50mm)。
结果 PDF 命名: layout_{原名}.pdf
排版前流程: pdf_strip_illustrator -> remove_color -> 排版
可通过环境变量 LAYOUT_TYPES_FILTER 只跑指定类型。
"""

import json
import os
import random
import shutil
import sys
import traceback
from datetime import datetime

for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import fitz
import ai_toolkit_modifications
from ai_toolkit_modifications import apply_layout_pure_python, ZOOM_FACTOR
from pdf_strip_illustrator import remove_illustrator_private_data
from remove_color import re_color_pdf

JSON_DIR = r"C:\Users\Administrator\Desktop\detection_pdf_purecode\layout_result\json"
OUTPUT_ROOT = r"C:\Users\Administrator\Desktop\detection_pdf_purecode\layout_result\layout_out"
PREP_ROOT = os.path.join(OUTPUT_ROOT, "_prep")
CLEANED_DIR = os.path.join(PREP_ROOT, "_cleaned")
NOCOLOR_DIR = os.path.join(PREP_ROOT, "_nocolor")
PREVIEW_DIR = os.path.join(PREP_ROOT, "_preview")

_PREP_CACHE = {}

TARGET_WIDTH_MM = 50.0
TARGET_HEIGHT_MM = 50.0
WIDTH_JITTER = 5.0
HEIGHT_JITTER = 5.0
MIN_VALID_BOXES = 4

LAYOUT_TYPES = {
    "type4_fixed_width_random_height": "宽不变+高随机(正侧唛同高)",
    "type5_fixed_height_random_width": "高不变+宽随机",
    "type6_random_width_random_height": "宽高都随机",
}


def _rand_width(rng):
    return rng.uniform(TARGET_WIDTH_MM - WIDTH_JITTER, TARGET_WIDTH_MM + WIDTH_JITTER)


def _rand_height(rng):
    return rng.uniform(TARGET_HEIGHT_MM - HEIGHT_JITTER, TARGET_HEIGHT_MM + HEIGHT_JITTER)


def _orig_dims(record):
    d = record.get("dieline_size_mm") or {}
    box = (record.get("box_size_mm") or {}).get("排版尺寸") or [0, 0, 0, 0]
    w_main = float(d.get("w_main_mm") or box[0] or 1.0)
    w_side = float(d.get("w_side_mm") or box[1] or 1.0)
    h_panel = float(d.get("h_panel_mm") or box[2] or 1.0)
    h_flap = float(d.get("h_flap_mm") or box[3] or 0.0)
    return w_main, w_side, h_panel, h_flap


def gen_box_size_mm(layout_type, record, rng):
    """按规则生成随机排版尺寸 [正唛宽, 侧唛宽, 面板高, 摇盖高] (mm)。"""
    w_main_o, w_side_o, h_panel_o, h_flap_o = _orig_dims(record)
    w_main_o = max(w_main_o, 1.0)
    w_side_o = max(w_side_o, 1.0)
    h_panel_o = max(h_panel_o, 1.0)

    if layout_type == "type4_fixed_width_random_height":
        h_panel = _rand_height(rng)
        h_flap = _rand_height(rng) if h_flap_o > 0 else 0.0
        return {"排版尺寸": [w_main_o, w_side_o, h_panel, h_flap]}

    if layout_type == "type5_fixed_height_random_width":
        w_main = _rand_width(rng)
        w_side = _rand_width(rng)
        return {"排版尺寸": [w_main, w_side, h_panel_o, h_flap_o]}

    if layout_type == "type6_random_width_random_height":
        w_main = _rand_width(rng)
        w_side = _rand_width(rng)
        h_panel = _rand_height(rng)
        h_flap = _rand_height(rng) if h_flap_o > 0 else 0.0
        return {"排版尺寸": [w_main, w_side, h_panel, h_flap]}

    raise ValueError(f"未知排版类型: {layout_type}")


def _pick_source_pdf(record):
    for key in ("source_pdf", "cleaned_pdf"):
        path = record.get(key)
        if path and os.path.isfile(path):
            return os.path.abspath(path)
    return None


def prepare_pdf(record, strip_illustrator=True, remove_color=True):
    """规范 PDF + 去底色，同一 record 只处理一次（缓存）。"""
    base = record["_base"]
    if base in _PREP_CACHE:
        return _PREP_CACHE[base]

    source_pdf = _pick_source_pdf(record)
    if not source_pdf:
        raise FileNotFoundError("JSON 中无可用 PDF 路径")

    pdf_for_layout = source_pdf
    cleaned_pdf_path = None
    nocolor_pdf_path = None

    if strip_illustrator:
        os.makedirs(CLEANED_DIR, exist_ok=True)
        cleaned_pdf_path = os.path.join(CLEANED_DIR, f"{base}_clean.pdf")
        if os.path.exists(cleaned_pdf_path) and os.path.getsize(cleaned_pdf_path) > 0:
            pdf_for_layout = cleaned_pdf_path
            print(f"[Prep] 使用已有规范 PDF: {cleaned_pdf_path}")
        else:
            print(f"[Prep] 规范 PDF: {source_pdf} -> {cleaned_pdf_path}")
            try:
                remove_illustrator_private_data(source_pdf, cleaned_pdf_path)
                if os.path.exists(cleaned_pdf_path):
                    pdf_for_layout = cleaned_pdf_path
                else:
                    print("[Prep] 警告: 规范化未生成文件, 回退使用原 PDF。")
            except Exception as e:
                print(f"[Prep] 警告: 规范化失败 ({e}), 回退使用原 PDF。")

    if remove_color:
        detections = record.get("final_boxes") or []
        cells_pt = [[v / ZOOM_FACTOR for v in d["box"]] for d in detections if d.get("box")]
        try:
            os.makedirs(PREVIEW_DIR, exist_ok=True)
            os.makedirs(NOCOLOR_DIR, exist_ok=True)
            preview_base = os.path.splitext(os.path.basename(pdf_for_layout))[0]
            preview_png = os.path.join(PREVIEW_DIR, f"{preview_base}_preview.png")
            if not os.path.exists(preview_png):
                _doc = fitz.open(pdf_for_layout)
                _doc[0].get_pixmap(matrix=fitz.Matrix(2, 2)).save(preview_png)
                _doc.close()

            expected_nocolor = os.path.join(
                NOCOLOR_DIR,
                os.path.basename(pdf_for_layout).replace(".pdf", "_nocolor.pdf"),
            )
            if os.path.exists(expected_nocolor) and os.path.getsize(expected_nocolor) > 0:
                nocolor_pdf_path = expected_nocolor
                pdf_for_layout = nocolor_pdf_path
                print(f"[Prep] 使用已有去色 PDF: {nocolor_pdf_path}")
            else:
                print(f"[Prep] 去底色 (刀版格={len(cells_pt)}): {pdf_for_layout}")
                nocolor_pdf = re_color_pdf(pdf_for_layout, preview_png, NOCOLOR_DIR, cells=cells_pt)
                if nocolor_pdf and os.path.exists(nocolor_pdf) and nocolor_pdf != pdf_for_layout:
                    nocolor_pdf_path = nocolor_pdf
                    pdf_for_layout = nocolor_pdf
                    print(f"[Prep] 已生成去色 PDF: {nocolor_pdf_path}")
                else:
                    print("[Prep] 未执行去色或未生成新文件, 使用当前 PDF 继续。")
        except Exception as e:
            print(f"[Prep] 警告: 去色步骤失败 ({e}), 继续使用当前 PDF。")

    result = {
        "source_pdf": source_pdf,
        "cleaned_pdf": cleaned_pdf_path,
        "nocolor_pdf": nocolor_pdf_path,
        "pdf_for_layout": pdf_for_layout,
    }
    _PREP_CACHE[base] = result
    return result


def _load_json_records(json_dir):
    records = []
    for name in sorted(os.listdir(json_dir)):
        if not name.endswith(".json") or name.startswith("_"):
            continue
        path = os.path.join(json_dir, name)
        try:
            with open(path, "r", encoding="utf-8") as f:
                record = json.load(f)
        except Exception:
            continue
        if not record.get("success"):
            continue
        boxes = [d for d in (record.get("final_boxes") or []) if d.get("box")]
        if len(boxes) < MIN_VALID_BOXES:
            continue
        record["_json_path"] = path
        record["_base"] = os.path.splitext(name)[0]
        records.append(record)
    return records


def layout_one(record, layout_type, out_dir, rng, margin_mm=0.0):
    prep = prepare_pdf(record)
    pdf_path = prep["pdf_for_layout"]

    base = record["_base"]
    out_pdf = os.path.join(out_dir, f"layout_{base}.pdf")
    out_preview = os.path.join(out_dir, f"layout_{base}_preview.png")
    meta_path = os.path.join(out_dir, f"layout_{base}.meta.json")

    if os.path.exists(out_pdf) and os.path.getsize(out_pdf) > 0:
        return out_pdf, out_preview, None, True

    box_size_mm = gen_box_size_mm(layout_type, record, rng)
    layout_instruction = record.get("layout_instruction") or ai_toolkit_modifications.__dict__.get("DEFAULT_LAYOUT_INSTRUCTION")
    if not layout_instruction:
        layout_instruction = {
            "正唛内容": [1, 3], "侧唛内容": [2, 4],
            "正唛上摇盖": [1, 3], "侧唛上摇盖": [2, 4],
            "正唛下摇盖": [1, 3], "侧唛下摇盖": [2, 4],
        }

    os.makedirs(out_dir, exist_ok=True)
    ai_toolkit_modifications.output_dir = out_dir

    layout_pdf, preview = apply_layout_pure_python(
        pdf_path,
        layout_instruction,
        box_size_mm,
        record.get("final_boxes") or [],
        zoom_factor=ZOOM_FACTOR,
        resize=True,
        margin_mm=margin_mm,
    )

    if not layout_pdf or not os.path.exists(layout_pdf):
        raise RuntimeError("排版未生成 PDF")

    if os.path.abspath(layout_pdf) != os.path.abspath(out_pdf):
        shutil.move(layout_pdf, out_pdf)
    if preview and os.path.exists(preview):
        if os.path.abspath(preview) != os.path.abspath(out_preview):
            shutil.move(preview, out_preview)

    meta = {
        "layout_type": layout_type,
        "layout_type_desc": LAYOUT_TYPES[layout_type],
        "source_json": record["_json_path"],
        "source_pdf": prep["source_pdf"],
        "cleaned_pdf": prep["cleaned_pdf"],
        "nocolor_pdf": prep["nocolor_pdf"],
        "pdf_for_layout": pdf_path,
        "box_size_mm": box_size_mm,
        "layout_pdf": out_pdf,
        "layout_preview": out_preview if os.path.exists(out_preview) else None,
        "finished_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

    return out_pdf, out_preview, box_size_mm, False


def _active_layout_types():
    filt = os.environ.get("LAYOUT_TYPES_FILTER", "").strip()
    if not filt:
        return LAYOUT_TYPES
    keys = [k.strip() for k in filt.split(",") if k.strip()]
    return {k: LAYOUT_TYPES[k] for k in keys if k in LAYOUT_TYPES}


def main():
    seed = int(os.environ.get("LAYOUT_SEED", "20260708"))
    rng = random.Random(seed)
    active_types = _active_layout_types()
    if not active_types:
        print("[X] LAYOUT_TYPES_FILTER 未匹配到任何类型")
        sys.exit(1)

    records = _load_json_records(JSON_DIR)
    print(f"共 {len(records)} 个有效 JSON 待排版 (本次 {len(active_types)} 类尺寸)。")
    print(f"输出目录: {OUTPUT_ROOT}")
    print(f"随机宽度: {TARGET_WIDTH_MM} ± {WIDTH_JITTER} mm")
    print(f"随机高度: {TARGET_HEIGHT_MM} ± {HEIGHT_JITTER} mm")
    print(f"本次类型: {', '.join(active_types.keys())}")

    summary = {
        "total_json": len(records),
        "types": list(active_types.keys()),
        "ok": [],
        "skipped": [],
        "fail": [],
        "started_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }

    total_jobs = len(records) * len(active_types)
    job_idx = 0

    for layout_type, desc in active_types.items():
        type_dir = os.path.join(OUTPUT_ROOT, layout_type)
        print(f"\n{'=' * 60}\n[{layout_type}] {desc}\n{'=' * 60}")

        for i, record in enumerate(records, 1):
            job_idx += 1
            base = record["_base"]
            print(f"\n--- [{job_idx}/{total_jobs}] {layout_type} | {base} ({i}/{len(records)}) ---")
            try:
                out_pdf, _, box_size_mm, was_skipped = layout_one(record, layout_type, type_dir, rng)
                entry = {
                    "name": base,
                    "layout_type": layout_type,
                    "pdf": out_pdf,
                    "box_size_mm": box_size_mm,
                }
                if was_skipped:
                    summary["skipped"].append(entry)
                    print(f"[跳过] 已存在: {out_pdf}")
                else:
                    summary["ok"].append(entry)
                    dims = (box_size_mm or {}).get("排版尺寸", [])
                    print(f"[完成] {out_pdf}")
                    if dims:
                        print(f"       尺寸(mm): 正={dims[0]:.1f} 侧={dims[1]:.1f} 高={dims[2]:.1f} 盖={dims[3]:.1f}")
            except Exception as e:
                traceback.print_exc()
                summary["fail"].append({
                    "name": base,
                    "layout_type": layout_type,
                    "error": str(e),
                })

    summary["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    summary_path = os.path.join(OUTPUT_ROOT, "_summary.json")
    os.makedirs(OUTPUT_ROOT, exist_ok=True)
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    print("\n" + "=" * 60)
    print(f"排版结束: 成功 {len(summary['ok'])}, 跳过 {len(summary['skipped'])}, 失败 {len(summary['fail'])}")
    print(f"汇总: {summary_path}")


if __name__ == "__main__":
    main()
