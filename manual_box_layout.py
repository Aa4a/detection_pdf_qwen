"""
手动框排版：完全按传入的检测框提取/编组/复制，不外扩、不贴刀版。

- 不做 find_best_candidate
- 不外扩到黑框
- 只选中「完全落在检测框内」的对象（越界/更大的黑框一律不带）
- 自带 Illustrator 渲染，不改 ai_toolkit_modifications.py
"""

import os
import traceback
from datetime import datetime

import fitz
import win32com.client

ZOOM_FACTOR = 2.0
output_dir = "result"
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

LABEL_MAP = {
    "正唛内容": "front_content",
    "侧唛内容": "side_content",
    "正唛上摇盖": "front_flap_up",
    "侧唛上摇盖": "side_flap_up",
    "正唛下摇盖": "front_flap_down",
    "侧唛下摇盖": "side_flap_down",
}


def _match_label(class_name):
    for k, v in LABEL_MAP.items():
        if k in (class_name or ""):
            return k, v
    return None, None


def apply_layout_illustrator_strict(input_pdf, output_pdf, refined_detections, box_size_mm, margin_mm=0.0):
    """严格按 detection box 提取：对象须完全落在框内，不带框外黑边。"""
    input_pdf_abs = os.path.abspath(input_pdf).replace("\\", "/")
    output_pdf_abs = os.path.abspath(output_pdf).replace("\\", "/")

    mm2pt = 2.834645
    front_w_mm, side_w_mm, panel_h_mm, flap_h_mm = box_size_mm["排版尺寸"]
    frontW = front_w_mm * mm2pt
    sideW = side_w_mm * mm2pt
    panelH = panel_h_mm * mm2pt
    flapH = flap_h_mm * mm2pt
    margin_pt = margin_mm * mm2pt

    colWidths = [frontW, sideW, frontW, sideW]
    rowHeights = [flapH, panelH, flapH]
    gridW = sum(colWidths)
    gridH = sum(rowHeights)

    jsx_dynamic_parts = ""

    for i, det in enumerate(refined_detections):
        cls_name = det["class_name"]
        targets = det.get("assigned_targets", [])
        if not targets:
            continue

        xmin, ymin, xmax, ymax = det["box"]
        src_w = xmax - xmin
        src_h = ymax - ymin
        if src_w <= 0 or src_h <= 0:
            print(f"[WARN] 跳过退化框 [{i}] {cls_name}: {det['box']}", flush=True)
            continue

        print(
            f"[排版] det[{i}] {cls_name}: box=[{xmin:.1f},{ymin:.1f},{xmax:.1f},{ymax:.1f}] "
            f"src_w={src_w:.1f} src_h={src_h:.1f} targets={targets}",
            flush=True,
        )

        if "上摇盖" in cls_name:
            row_idx = 0
        elif "下摇盖" in cls_name:
            row_idx = 2
        else:
            row_idx = 1

        jsx_dynamic_parts += f"""
        __step = "extract det {i} ({cls_name})";
        app.activeDocument = sourceDoc;
        sourceDoc.selection = null;

        var bl = abRect[0] + {xmin};
        var bt = abRect[1] - {ymin};
        var br = abRect[0] + {xmax};
        var bb = abRect[1] - {ymax};
        var bw = Math.abs(br - bl);
        var bh = Math.abs(bt - bb);

        extractItemsStrict(sourceDoc.pageItems, bl, br, bt, bb, bw, bh);

        var sel = sourceDoc.selection;
        if (sel.length > 0) {{
            var tempSourceGroup = sourceDoc.groupItems.add();
            for (var k = sel.length - 1; k >= 0; k--) {{
                sel[k].duplicate(tempSourceGroup, ElementPlacement.PLACEATBEGINNING);
            }}

            var pastedGroup = tempSourceGroup.duplicate(newDoc.layers[0], ElementPlacement.PLACEATEND);
            tempSourceGroup.remove();

            app.activeDocument = newDoc;
            newDoc.selection = null;

            if (pastedGroup) {{
        """

        for fid in targets:
            col_idx = (fid % 100) - 1
            cell_w = colWidths[col_idx]
            cell_h = rowHeights[row_idx]
            cell_left = sum(colWidths[:col_idx])
            cell_top = gridH - sum(rowHeights[:row_idx])

            jsx_dynamic_parts += f"""
                __step = "place det {i} -> col {col_idx} (fid {fid})";
                var dup = pastedGroup.duplicate();

                // 再清一次：越界大框 / 刀版色描边
                stripOutsideFrames(dup, bw, bh);

                var target_w = {cell_w} - 2 * {margin_pt};
                var target_h = {cell_h} - 2 * {margin_pt};
                if (target_w <= 0) target_w = 1;
                if (target_h <= 0) target_h = 1;

                // 严格按检测框尺寸等比缩放进格子
                var scale_percent = Math.min(target_w / {src_w}, target_h / {src_h}) * 100;
                if (!isFinite(scale_percent) || scale_percent <= 0) scale_percent = 100;
                dup.resize(scale_percent, scale_percent, true, true, true, true, true);

                var gBounds = dup.visibleBounds;
                var gCenter_x = (gBounds[0] + gBounds[2]) / 2.0;
                var gCenter_y = (gBounds[1] + gBounds[3]) / 2.0;
                var cell_cx = {cell_left} + {cell_w} / 2.0;
                var cell_cy = {cell_top} - {cell_h} / 2.0;
                dup.translate(cell_cx - gCenter_x, cell_cy - gCenter_y);
            """

        jsx_dynamic_parts += """
                pastedGroup.remove();
            }
        }
        """

    try:
        try:
            app = win32com.client.GetActiveObject("Illustrator.Application")
        except Exception:
            app = win32com.client.Dispatch("Illustrator.Application")
    except Exception as e:
        print(f"❌ 无法启动 Illustrator: {e}")
        return None

    jsx_code = f"""
    var __step = "init";
    try {{
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

        var mm2pt = 2.834645;
        var colWidths = [{frontW}, {sideW}, {frontW}, {sideW}];
        var rowHeights = [{flapH}, {panelH}, {flapH}];
        var gridW = {gridW};
        var gridH = {gridH};

        __step = "open source PDF";
        var fileToOpen = new File("{input_pdf_abs}");
        var sourceDoc = app.open(fileToOpen);
        var abRect = sourceDoc.artboards[0].artboardRect;

        __step = "create newDoc";
        var docMargin = 50 * mm2pt;
        var newDoc = app.documents.add(sourceDoc.documentColorSpace, gridW + 2 * docMargin, gridH + 2 * docMargin);
        newDoc.artboards[0].artboardRect = [0, gridH, gridW, 0];
        app.activeDocument = newDoc;

        var dieLayer = newDoc.layers.add();
        dieLayer.name = "刀版线";
        var dieColor = (newDoc.documentColorSpace == DocumentColorSpace.CMYK) ? new CMYKColor() : new RGBColor();
        if (newDoc.documentColorSpace == DocumentColorSpace.CMYK) {{ dieColor.magenta = 100; }}
        else {{ dieColor.red = 255; dieColor.blue = 255; }}

        var currentY = gridH;
        for (var r = 0; r < 3; r++) {{
            var currentX = 0;
            for (var c = 0; c < 4; c++) {{
                var w = colWidths[c];
                var h = rowHeights[r];
                var panel = dieLayer.pathItems.rectangle(currentY, currentX, w, h);
                panel.filled = false;
                panel.stroked = true;
                panel.strokeColor = dieColor;
                panel.strokeWidth = 1.0;
                currentX += w;
            }}
            currentY -= rowHeights[r];
        }}

        function colorIsDieColor(col) {{
            if (!col) return false;
            try {{
                var tn = col.typename;
                if (tn === "RGBColor") {{
                    var isCyan = (col.blue > 100 && col.red < 130 && (col.blue - col.red) > 40);
                    var isMagenta = (col.red > 150 && col.green < 90 && col.blue > 90);
                    return isCyan || isMagenta;
                }} else if (tn === "CMYKColor") {{
                    var isCyanC = (col.cyan > 45 && col.magenta < 45 && col.yellow < 45);
                    var isMagentaC = (col.magenta > 50 && col.cyan < 45 && col.yellow < 45);
                    return isCyanC || isMagentaC;
                }} else if (tn === "SpotColor") {{
                    return colorIsDieColor(col.spot.color);
                }}
            }} catch(e) {{}}
            return false;
        }}

        function strokeIsDieline(item) {{
            try {{
                if (item.typename === "PathItem") {{
                    return (!item.filled && item.stroked && colorIsDieColor(item.strokeColor));
                }} else if (item.typename === "CompoundPathItem" && item.pathItems && item.pathItems.length > 0) {{
                    var noFill = true, dieStroke = false;
                    for (var cp = 0; cp < item.pathItems.length; cp++) {{
                        if (item.pathItems[cp].filled) noFill = false;
                        if (item.pathItems[cp].stroked && colorIsDieColor(item.pathItems[cp].strokeColor)) dieStroke = true;
                    }}
                    return (noFill && dieStroke);
                }}
            }} catch(e) {{}}
            return false;
        }}

        // 删除：刀版色描边，或尺寸 >= 检测框 90% 的框线/填充环（黑框）
        function stripOutsideFrames(item, refW, refH) {{
            if (!item) return;
            if (item.typename === "GroupItem") {{
                for (var i = item.pageItems.length - 1; i >= 0; i--) {{
                    stripOutsideFrames(item.pageItems[i], refW, refH);
                }}
                return;
            }}
            try {{
                if (strokeIsDieline(item)) {{ item.remove(); return; }}
                var bounds = item.visibleBounds;
                var cw = Math.abs(bounds[2] - bounds[0]);
                var ch = Math.abs(bounds[1] - bounds[3]);
                var spans = (cw >= refW * 0.9) || (ch >= refH * 0.9);
                if (!spans) return;
                if (item.typename === "PathItem") {{
                    // 无填充描边大框，或近似空心黑框（填充环）
                    if ((!item.filled && item.stroked) || (item.filled && !item.stroked)) {{
                        item.remove();
                    }}
                }} else if (item.typename === "CompoundPathItem") {{
                    item.remove();
                }}
            }} catch(e) {{}}
        }}

        // 严格提取：中心在框内，且整体不超出检测框（黑框更大则直接跳过）
        function extractItemsStrict(items, bl, br, bt, bb, bw, bh) {{
            var eps = 0.5;
            var pick_l = bl - eps, pick_r = br + eps;
            var pick_t = bt + eps, pick_b = bb - eps;

            for (var i = 0; i < items.length; i++) {{
                var item = items[i];
                if (item.hidden || item.locked || item.guides) continue;
                try {{
                    var bounds = item.visibleBounds;
                    var l = bounds[0], t = bounds[1], r = bounds[2], b = bounds[3];
                    var itemW = Math.abs(r - l);
                    var itemH = Math.abs(t - b);

                    if (r < pick_l || l > pick_r || b > pick_t || t < pick_b) continue;

                    // 比检测框更大（含外侧黑框）→ 组则下钻，否则跳过
                    if ((itemW > bw + eps) || (itemH > bh + eps)) {{
                        if (item.typename === "GroupItem") {{
                            extractItemsStrict(item.pageItems, bl, br, bt, bb, bw, bh);
                        }}
                        continue;
                    }}

                    // 必须完全落在检测框内
                    if (l < pick_l || r > pick_r || b < pick_b || t > pick_t) {{
                        if (item.typename === "GroupItem") {{
                            extractItemsStrict(item.pageItems, bl, br, bt, bb, bw, bh);
                        }}
                        continue;
                    }}

                    var icx = (l + r) / 2.0, icy = (t + b) / 2.0;
                    if (icx < pick_l || icx > pick_r || icy < pick_b || icy > pick_t) continue;
                    item.selected = true;
                }} catch(e) {{}}
            }}
        }}

        try {{ app.activeDocument = sourceDoc; app.executeMenuCommand('unlockAll'); }} catch(e) {{}}
        try {{ app.activeDocument = sourceDoc; app.executeMenuCommand('showAll'); }} catch(e) {{}}

        {jsx_dynamic_parts}

        __step = "resize final artboard";
        var margin = 50 * mm2pt;
        newDoc.artboards[0].artboardRect = [-margin, gridH + margin, gridW + margin, -margin];

        __step = "saveAs PDF";
        var destFile = new File("{output_pdf_abs}");
        var saveOpts = new PDFSaveOptions();
        saveOpts.preserveEditability = true;
        newDoc.saveAs(destFile, saveOpts);
        newDoc.close(SaveOptions.DONOTSAVECHANGES);
        sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
        app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;

        "Success";
    }} catch(e) {{
        "JSX Error @[" + __step + "]: " + e.message;
    }}
    """

    try:
        result = app.DoJavaScript(jsx_code)
        result_str = "" if result is None else str(result)
        file_ok = os.path.exists(output_pdf)
        if result_str.startswith("JSX Error") or not file_ok:
            print(
                f"[ERR] Illustrator 渲染失败: JSX返回={result_str!r}, 输出文件存在={file_ok}, "
                f"目标={output_pdf_abs}",
                flush=True,
            )
            return None
        print(f"[OK] Illustrator 渲染成功: {output_pdf} (JSX={result_str!r})", flush=True)
        return output_pdf
    except Exception as e:
        print(f"[ERR] Python 侧调用 Illustrator 出错：{e}", flush=True)
        traceback.print_exc()
        return None


def apply_layout_manual_boxes(
    pdf_path,
    layout_instruction,
    box_size_mm,
    detections,
    zoom_factor=2.0,
    margin_mm=0.0,
    out_dir=None,
):
    """
    完全按 detections 的 box 排版，不做贴刀版/外扩。
    box 为 zoom 图像坐标。
    out_dir: 输出目录, 默认用模块级 output_dir。
    """
    pdf_path_abs = os.path.abspath(pdf_path)
    save_dir = out_dir or output_dir
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    try:
        dims = box_size_mm["排版尺寸"] if isinstance(box_size_mm, dict) else box_size_mm
        if not os.path.exists(pdf_path_abs):
            raise FileNotFoundError(pdf_path_abs)
    except Exception as e:
        print(f"❌ 无法读取PDF或解析尺寸: {e}")
        return None, None

    refined_detections = []
    for i, det in enumerate(detections):
        label = det.get("class_name", "")
        x1, y1, x2, y2 = det["box"]
        rect = fitz.Rect(x1 / zoom_factor, y1 / zoom_factor, x2 / zoom_factor, y2 / zoom_factor)

        if rect.width <= 0 or rect.height <= 0:
            print(f"[WARN] 跳过无效手动框 [{i}] {label}: {det['box']}", flush=True)
            continue

        matched_k, key = _match_label(label)
        if not key:
            print(f"[WARN] 未识别类别，跳过 [{i}] {label}", flush=True)
            continue

        new_det = det.copy()
        new_det["box"] = [rect.x0, rect.y0, rect.x1, rect.y1]
        new_det["target_key"] = key
        new_det["original_index"] = i
        new_det["assigned_targets"] = []
        new_det["normalized_class"] = matched_k
        refined_detections.append(new_det)
        print(
            f"[手动框] det[{i}] {matched_k}: "
            f"pt=[{rect.x0:.1f},{rect.y0:.1f},{rect.x1:.1f},{rect.y1:.1f}] "
            f"({rect.width:.1f}x{rect.height:.1f})",
            flush=True,
        )

    class_to_dets = {}
    for det in refined_detections:
        cls = det.get("normalized_class") or det.get("class_name")
        class_to_dets.setdefault(cls, []).append(det)

    for instr_label, target_ids in layout_instruction.items():
        if not target_ids:
            continue
        targets = [target_ids] if not isinstance(target_ids, list) else target_ids
        available_dets = class_to_dets.get(instr_label, [])
        if not available_dets:
            continue

        base_offset = 0
        if "上摇盖" in instr_label:
            base_offset = 100
        elif "下摇盖" in instr_label:
            base_offset = 200

        for i, fid in enumerate(targets):
            available_dets[i % len(available_dets)]["assigned_targets"].append(fid + base_offset)

    if not any(d.get("assigned_targets") for d in refined_detections):
        print("[ERR] 没有任何检测框被分配到目标格子", flush=True)
        return None, None

    base_name = os.path.splitext(os.path.basename(pdf_path_abs))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    final_pdf_path = os.path.join(save_dir, f"ManualLayout_{base_name}_{timestamp}.pdf")

    result_pdf = apply_layout_illustrator_strict(
        input_pdf=pdf_path_abs,
        output_pdf=final_pdf_path,
        refined_detections=refined_detections,
        box_size_mm=box_size_mm,
        margin_mm=margin_mm,
    )

    if not result_pdf or not os.path.exists(final_pdf_path):
        print(f"[ERR] 手动框排版失败: {final_pdf_path}, 尺寸mm={dims}", flush=True)
        return None, None

    preview_path = os.path.join(save_dir, f"{base_name}_{timestamp}_manual_preview.png")
    preview_doc = fitz.open(final_pdf_path)
    preview_doc[0].get_pixmap(matrix=fitz.Matrix(2, 2)).save(preview_path)
    preview_doc.close()

    print(f"[OK] 手动框排版完成: {final_pdf_path}", flush=True)
    print(f"[OK] 预览图: {preview_path}", flush=True)
    return final_pdf_path, preview_path


if __name__ == "__main__":
    input_pdf = os.path.abspath(
        r"C:\Users\Administrator\Desktop\detection_pdf_purecode"
        r"\output\5-662Z 923759内外箱唛头 -_20260812_111007.pdf"
    )

    detections = [
        {"box": [39, 495, 729, 1059], "class_name": "正唛内容"},
        {"box": [804, 496, 1491, 1051], "class_name": "侧唛内容"},
    ]

    layout_instruction = {
        "正唛内容": [1, 3],
        "侧唛内容": [2, 4],
        "正唛上摇盖": [1, 3],
        "侧唛上摇盖": [2, 4],
        "正唛下摇盖": [1, 3],
        "侧唛下摇盖": [2, 4],
    }

    box_size_mm = {"排版尺寸": [650, 410, 260, 205]}

    if not os.path.exists(input_pdf):
        print(f"❌ 找不到 PDF: {input_pdf}")
    else:
        pdf_path, preview_path = apply_layout_manual_boxes(
            input_pdf,
            layout_instruction,
            box_size_mm,
            detections,
            zoom_factor=ZOOM_FACTOR,
            margin_mm=0.0,
        )
        print("执行完成！")
        if pdf_path:
            print(f"  PDF : {pdf_path}")
            print(f"  预览: {preview_path}")
