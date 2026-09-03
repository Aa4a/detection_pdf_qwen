# -*- coding: utf-8 -*-
"""
Pipeline: qwen_adobe_detection -> ai_toolkit_modifications
1) 调用 visualize_dielines 检测刀版线 12 格
2) 将 final_boxes 作为 detections 传给 apply_layout_pure_python 完成排版
"""

import os
import sys

import fitz

from qwen_adobe_detection import visualize_dielines
from ai_toolkit_modifications import apply_layout_pure_python, ZOOM_FACTOR
from pdf_strip_illustrator import remove_illustrator_private_data
from remove_color import re_color_pdf


def run_pipeline(
    input_pdf,
    layout_instruction,
    box_size_mm,
    output_dir="result",
    vis_suffix="_vis",
    margin_mm=0.0,
    use_qwen=True,
    qwen_requirements="",
    strip_illustrator=True,
    cleaned_dir="result/_cleaned",
):
    """
    端到端: (清理 Illustrator 私有数据) + 检测 + 排版
    返回: dict { detection: ..., layout_pdf: str|None, layout_preview: str|None, cleaned_pdf: str }
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
        print(f"[Pipeline] 步骤 0/2 规范 PDF: {input_pdf} -> {cleaned_pdf_path}")
        print("=" * 60)
        try:
            remove_illustrator_private_data(input_pdf, cleaned_pdf_path)
            if os.path.exists(cleaned_pdf_path):
                pdf_for_pipeline = cleaned_pdf_path
            else:
                print("[Pipeline] 警告: 规范化未生成文件, 回退使用原始 PDF。")
        except Exception as e:
            print(f"[Pipeline] 警告: 规范化失败 ({e}), 回退使用原始 PDF。")

    # ===== 步骤 1: 检测刀版线 (先识别刀版线, 后续才能据此删铺满刀版格的底色) =====
    print("=" * 60)
    print(f"[Pipeline] 步骤 1/2 检测刀版线: {pdf_for_pipeline}")
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
            "layout_pdf": None,
            "layout_preview": None,
        }

    detections = det_result.get("final_boxes", [])
    print(f"[Pipeline] 检测成功, 共 {len(detections)} 个 box。可视化: {vis_pdf}")

    # ===== 步骤 1.5: 去除"铺满刀版格的底色" (用检测到的刀版格做依据) =====
    try:
        # final_boxes 为 zoom_factor=2 图像坐标, 还原为 PDF 点坐标 (除以 2)
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
    print("[Pipeline] 步骤 2/2 排版")
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
        "layout_pdf": layout_pdf,
        "layout_preview": layout_preview,
        "cleaned_pdf": cleaned_pdf_path,
    }


if __name__ == "__main__":
    # 支持单个 PDF 文件或整个文件夹
    input_path = r"C:\Users\Administrator\Desktop\detection_pdf_purecode\layout_result\AyxPytzjGtXNVoMxAwOE_clean.pdf"

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

    box_size_mm = {
      "排版尺寸": [
        79.87,
        55.87,
        55.87,
        20
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

    print(f"共发现 {len(pdf_paths)} 个 PDF 待处理。")

    ok, fail = [], []
    for i, in_path in enumerate(pdf_paths, 1):
        name = os.path.basename(in_path)
        print(f"\n########## [{i}/{len(pdf_paths)}] {name} ##########")
        try:
            result = run_pipeline(
                input_pdf=in_path,
                layout_instruction=layout_instruction,
                box_size_mm=box_size_mm,
                margin_mm=0.0,
            )
            if result.get("layout_pdf"):
                ok.append((name, result["layout_pdf"]))
            else:
                fail.append((name, "未生成排版 PDF"))
        except Exception as e:
            import traceback
            traceback.print_exc()
            fail.append((name, str(e)))

    print("\n" + "=" * 60)
    print(f"处理完成: 成功 {len(ok)} 个, 失败 {len(fail)} 个")
    if ok:
        print("成功列表:")
        for n, p in ok:
            print(f"  + {n} -> {p}")
    if fail:
        print("失败列表:")
        for n, err in fail:
            print(f"  - {n}: {err}")
