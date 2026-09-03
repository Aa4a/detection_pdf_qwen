# -*- coding: utf-8 -*-
"""
Pipeline (自动尺寸 + 放大版):
不再手写 box_size_mm, 而是在检测刀版格后, 用检测框的实际尺寸自动推算排版尺寸,
再整体放大 SCALE_FACTOR(=1.3) 倍进行排版。

流程: (清理 Illustrator 私有数据) -> 检测 -> 自动推尺寸并×1.3 -> 去底色 -> 排版
"""

import os
import sys
from statistics import median

# 底层脚本 (remove_color / resize_adobe 等) 有 emoji 或中文打印, Windows 控制台默认 GBK
# 会抛 UnicodeEncodeError, 这里把标准输出/错误统一改成 UTF-8 (无法编码时替换而非崩溃)。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import fitz

import resize_adobe
from qwen_adobe_detection import visualize_dielines
from resize_adobe import apply_layout_pure_python, ZOOM_FACTOR
from pdf_strip_illustrator import remove_illustrator_private_data
from remove_color import re_color_pdf

# 排版整体放大倍数
SCALE_FACTOR = 1.3

# 结果输出目录 (所有中间/最终文件都放这里)
OUTPUT_DIR = r"C:\Users\Administrator\Desktop\detection_pdf_purecode\cleaned\result"

# resize_adobe 里最终排版 PDF/预览图用的是它的模块级全局 output_dir, 这里一并覆盖,
# 保证最终 Layout_*.pdf 也保存到 OUTPUT_DIR。
os.makedirs(OUTPUT_DIR, exist_ok=True)
resize_adobe.output_dir = OUTPUT_DIR


def _pt_to_mm(pt):
    return pt * 25.4 / 72.0


def compute_box_size_mm(detections, zoom_factor=ZOOM_FACTOR, scale=SCALE_FACTOR):
    """
    从检测框自动推算排版尺寸(mm), 并整体放大 scale 倍。
    排版尺寸 = [正唛宽, 侧唛宽, 面板高, 摇盖高]
      - 正唛宽 : 所有"正唛"框宽度的中位数
      - 侧唛宽 : 所有"侧唛"框宽度的中位数
      - 面板高 : 所有"内容"框高度的中位数
      - 摇盖高 : 所有"摇盖"框高度的中位数 (无摇盖则为 0)
    检测框为 zoom_factor 倍图像坐标, 需先 /zoom_factor 转 pt 再转 mm。
    """
    front_w, side_w, panel_h, flap_h = [], [], [], []
    for d in detections:
        box = d.get("box")
        if not box:
            continue
        cls = d.get("class_name", "") or ""
        w_mm = _pt_to_mm(abs(box[2] - box[0]) / zoom_factor)
        h_mm = _pt_to_mm(abs(box[3] - box[1]) / zoom_factor)

        if "正唛" in cls:
            front_w.append(w_mm)
        elif "侧唛" in cls:
            side_w.append(w_mm)

        if "摇盖" in cls:
            flap_h.append(h_mm)
        elif "内容" in cls:
            panel_h.append(h_mm)

    def _med(vals):
        return median(vals) if vals else 0.0

    dims = [
        _med(front_w) * scale,
        _med(side_w) * scale,
        _med(panel_h) * scale,
        _med(flap_h) * scale,
    ]
    return {"排版尺寸": dims}


def run_pipeline_auto(
    input_pdf,
    layout_instruction,
    output_dir="result",
    vis_suffix="_vis",
    margin_mm=0.0,
    use_qwen=True,
    qwen_requirements="",
    strip_illustrator=True,
    cleaned_dir="result/_cleaned",
    scale=SCALE_FACTOR,
):
    """
    端到端: (清理) + 检测 + 自动推尺寸(×scale) + 去底色 + 排版
    返回: dict { detection, box_size_mm, layout_pdf, layout_preview, cleaned_pdf }
    """
    input_pdf = os.path.abspath(input_pdf)
    if not os.path.exists(input_pdf):
        raise FileNotFoundError(f"输入 PDF 不存在: {input_pdf}")

    os.makedirs(output_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(input_pdf))[0]
    vis_pdf = os.path.join(output_dir, f"{base}{vis_suffix}.pdf")

    # ===== 步骤 0: 规范 PDF (剥离 Illustrator 私有数据) =====
    pdf_for_pipeline = input_pdf
    cleaned_pdf_path = None
    if strip_illustrator:
        os.makedirs(cleaned_dir, exist_ok=True)
        cleaned_pdf_path = os.path.abspath(os.path.join(cleaned_dir, f"{base}_clean.pdf"))
        print("=" * 60)
        print(f"[Pipeline] 步骤 0 规范 PDF: {input_pdf} -> {cleaned_pdf_path}")
        print("=" * 60)
        try:
            remove_illustrator_private_data(input_pdf, cleaned_pdf_path)
            if os.path.exists(cleaned_pdf_path):
                pdf_for_pipeline = cleaned_pdf_path
            else:
                print("[Pipeline] 警告: 规范化未生成文件, 回退使用原始 PDF。")
        except Exception as e:
            print(f"[Pipeline] 警告: 规范化失败 ({e}), 回退使用原始 PDF。")

    # ===== 步骤 1: 检测刀版线 =====
    print("=" * 60)
    print(f"[Pipeline] 步骤 1 检测刀版线: {pdf_for_pipeline}")
    print("=" * 60)
    det_result = visualize_dielines(
        input_path=pdf_for_pipeline,
        output_path=vis_pdf,
        use_qwen=use_qwen,
        qwen_requirements=qwen_requirements,
    )

    if not (isinstance(det_result, dict) and det_result.get("success")):
        print("[Pipeline] 检测失败, 终止。")
        return {
            "detection": det_result,
            "box_size_mm": None,
            "layout_pdf": None,
            "layout_preview": None,
            "cleaned_pdf": cleaned_pdf_path,
        }

    detections = det_result.get("final_boxes", [])
    print(f"[Pipeline] 检测成功, 共 {len(detections)} 个 box。可视化: {vis_pdf}")

    # ===== 步骤 1.1: 板块数量校验 (刀版线识别错误或不足 4 个板块 -> 跳过排版) =====
    valid_boxes = [d for d in detections if d.get("box")]
    if len(valid_boxes) < 4:
        print(f"[Pipeline] 识别到有效板块仅 {len(valid_boxes)} 个 (<4), 判定刀版线识别错误/不足, 跳过排版。")
        return {
            "detection": det_result,
            "box_size_mm": None,
            "layout_pdf": None,
            "layout_preview": None,
            "cleaned_pdf": cleaned_pdf_path,
            "skipped": True,
            "skip_reason": f"有效板块 {len(valid_boxes)} 个 (<4)",
        }

    # ===== 步骤 1.2: 从检测框自动推算排版尺寸, 并放大 scale 倍 =====
    box_size_mm = compute_box_size_mm(detections, zoom_factor=ZOOM_FACTOR, scale=scale)
    _d = box_size_mm["排版尺寸"]
    print("=" * 60)
    print(f"[Pipeline] 步骤 1.2 自动排版尺寸(×{scale}, mm): "
          f"正唛宽={_d[0]:.2f} 侧唛宽={_d[1]:.2f} 面板高={_d[2]:.2f} 摇盖高={_d[3]:.2f}")
    print("=" * 60)

    # ===== 步骤 1.5: 去除"铺满刀版格的底色" =====
    try:
        cells_pt = [[v / 2.0 for v in d["box"]] for d in detections if "box" in d]

        preview_dir = os.path.join(output_dir, "_preview")
        os.makedirs(preview_dir, exist_ok=True)
        preview_base = os.path.splitext(os.path.basename(pdf_for_pipeline))[0]
        preview_png = os.path.join(preview_dir, f"{preview_base}_preview.png")
        _doc = fitz.open(pdf_for_pipeline)
        _doc[0].get_pixmap(matrix=fitz.Matrix(2, 2)).save(preview_png)
        _doc.close()

        nocolor_dir = os.path.join(output_dir, "_nocolor")
        os.makedirs(nocolor_dir, exist_ok=True)
        print("=" * 60)
        print(f"[Pipeline] 步骤 1.5 去除底色 (刀版格={len(cells_pt)}): {pdf_for_pipeline}")
        print("=" * 60)
        nocolor_pdf = re_color_pdf(pdf_for_pipeline, preview_png, nocolor_dir, cells=cells_pt)
        if nocolor_pdf and os.path.exists(nocolor_pdf) and nocolor_pdf != pdf_for_pipeline:
            print(f"[Pipeline] 已生成去色 PDF: {nocolor_pdf}")
            pdf_for_pipeline = nocolor_pdf
        else:
            print("[Pipeline] 未执行去色或未生成新文件, 使用原 PDF 继续。")
    except Exception as e:
        print(f"[Pipeline] 警告: 去色步骤失败 ({e}), 继续使用当前 PDF。")

    # ===== 步骤 2: 排版 =====
    print("=" * 60)
    print("[Pipeline] 步骤 2 排版")
    print("=" * 60)
    layout_pdf, layout_preview = apply_layout_pure_python(
        pdf_for_pipeline,
        layout_instruction,
        box_size_mm,
        detections,
        zoom_factor=ZOOM_FACTOR,
        resize=True,
        margin_mm=margin_mm,
    )

    if layout_pdf:
        print(f"[Pipeline] 排版完成: {layout_pdf}")
    else:
        print("[Pipeline] 排版失败或被中断。")

    return {
        "detection": det_result,
        "box_size_mm": box_size_mm,
        "layout_pdf": layout_pdf,
        "layout_preview": layout_preview,
        "cleaned_pdf": cleaned_pdf_path,
    }


if __name__ == "__main__":
    # 支持单个 PDF 文件或整个文件夹
    input_path = r"C:\Users\Administrator\Desktop\detection_pdf_purecode\cleaned"

    layout_instruction = {
      "正唛内容": [
        1,
        3
      ],
      "侧唛内容": [
        2,
        4
      ],
      "正唛上摇盖": [
        1,
        3
      ],
      "侧唛上摇盖": [
        2,
        4
      ],
      "正唛下摇盖": [
        1,
        3
      ],
      "侧唛下摇盖": [
        2,
        4
      ]
    }

    # 根据输入路径判断是单文件还是文件夹
    if os.path.isfile(input_path):
        if not input_path.lower().endswith(".pdf"):
            print(f"❌ 输入文件不是 PDF: {input_path}")
            sys.exit(1)
        pdf_paths = [input_path]
    elif os.path.isdir(input_path):
        pdf_paths = sorted([
            os.path.join(input_path, f)
            for f in os.listdir(input_path)
            if f.lower().endswith(".pdf")
        ])
    else:
        print(f"❌ 输入路径不存在: {input_path}")
        sys.exit(1)

    print(f"共发现 {len(pdf_paths)} 个 PDF 待处理。放大倍数={SCALE_FACTOR}")

    ok, fail, skipped = [], [], []
    for i, in_path in enumerate(pdf_paths, 1):
        name = os.path.basename(in_path)
        print(f"\n########## [{i}/{len(pdf_paths)}] {name} ##########")
        try:
            result = run_pipeline_auto(
                input_pdf=in_path,
                layout_instruction=layout_instruction,
                output_dir=OUTPUT_DIR,
                cleaned_dir=os.path.join(OUTPUT_DIR, "_cleaned"),
                margin_mm=0.0,
                scale=SCALE_FACTOR,
            )
            if result.get("layout_pdf"):
                ok.append((name, result["layout_pdf"]))
            elif result.get("skipped"):
                skipped.append((name, result.get("skip_reason", "识别不足/错误")))
            else:
                fail.append((name, "未生成排版 PDF"))
        except Exception as e:
            import traceback
            traceback.print_exc()
            fail.append((name, str(e)))

    print("\n" + "=" * 60)
    print(f"处理完成: 成功 {len(ok)} 个, 跳过 {len(skipped)} 个, 失败 {len(fail)} 个")
    if ok:
        print("成功列表:")
        for n, p in ok:
            print(f"  + {n} -> {p}")
    if skipped:
        print("跳过列表 (刀版线识别错误/不足4个板块):")
        for n, reason in skipped:
            print(f"  ~ {n}: {reason}")
    if fail:
        print("失败列表:")
        for n, err in fail:
            print(f"  - {n}: {err}")
