import os
import sys
import math
import traceback
from datetime import datetime
import fitz  # PyMuPDF
import win32com.client

# ==========================================
# 0. 基础配置与工具
# ==========================================

output_dir = "result"
if not os.path.exists(output_dir): os.makedirs(output_dir)

ZOOM_FACTOR = 2.0  # 用于 PDF 坐标转换


def mm_to_pt(mm):
    """毫米转点 (1 pt = 1/72 inch, 1 inch = 25.4 mm)"""
    return mm * (72 / 25.4)


def calculate_iou(rect1, rect2):
    """计算两个矩形的交并比"""
    intersect = rect1 & rect2
    if intersect.is_empty:
        return 0.0
    intersect_area = intersect.width * intersect.height
    area1 = rect1.width * rect1.height
    area2 = rect2.width * rect2.height
    union_area = area1 + area2 - intersect_area
    if union_area <= 0: return 0.0
    return intersect_area / union_area


# ==========================================
# 1. 核心算法 (纯矢量分析)
# ==========================================

def find_best_candidate(page, current_ai_box, all_detections, zoom_factor=2.0):
    target_rect = fitz.Rect(current_ai_box)
    target_area = target_rect.width * target_rect.height
    target_center = fitz.Point((target_rect.x0 + target_rect.x1) / 2, (target_rect.y0 + target_rect.y1) / 2)

    neighbors = []
    for det in all_detections:
        d_box = det["box"]
        r = fitz.Rect(d_box[0] / zoom_factor, d_box[1] / zoom_factor, d_box[2] / zoom_factor, d_box[3] / zoom_factor)
        d_center = fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)
        dist = math.sqrt((d_center.x - target_center.x) ** 2 + (d_center.y - target_center.y) ** 2)
        if dist > 10:
            neighbors.append(r)

    paths = page.get_drawings()
    candidates = []
    all_considered_paths = []

    for path in paths:
        vec_rect = path['rect']
        vec_area = vec_rect.width * vec_rect.height
        if vec_area < 10: continue

        intersect = vec_rect & target_rect
        if intersect.is_empty: continue

        coverage = (intersect.width * intersect.height) / target_area
        if coverage < 0.8: continue

        contains_neighbor = False
        for neighbor in neighbors:
            neighbor_center = fitz.Point((neighbor.x0 + neighbor.x1) / 2, (neighbor.y0 + neighbor.y1) / 2)
            if vec_rect.contains(neighbor_center):
                contains_neighbor = True
                break
        if contains_neighbor: continue

        candidates.append((vec_rect, path))
        all_considered_paths.append(vec_rect)

    if candidates:
        sorted_candidates = sorted(candidates, key=lambda x: x[0].width * x[0].height)
        best_candidate = sorted_candidates[0]
        final_rect = best_candidate[0]

        if len(sorted_candidates) > 1:
            first_rect = sorted_candidates[0][0]
            iou = calculate_iou(first_rect, target_rect)
            width_diff = abs(first_rect.width - target_rect.width)
            height_diff = abs(first_rect.height - target_rect.height)
            is_size_too_close = (width_diff < 10) and (height_diff < 10)

            if iou > 0.85 or is_size_too_close:
                best_candidate = sorted_candidates[1]
                final_rect = best_candidate[0]

        return final_rect, True, {
            "ai_rect": target_rect,
            "candidates": all_considered_paths,
            "best_path": best_candidate[1]
        }
    else:
        safe_rect = target_rect
        if safe_rect.width <= 1: safe_rect = fitz.Rect(0, 0, 10, 10)
        return safe_rect, False, {
            "ai_rect": target_rect,
            "candidates": [],
            "best_path": None
        }


def smart_crop_box_action(page, current_ai_box, all_detections, zoom_factor=2.0):
    final_rect, is_precise, info = find_best_candidate(page, current_ai_box, all_detections, zoom_factor)
    return final_rect, is_precise


# ==========================================
# 2. Illustrator 渲染引擎 (无蒙版纯编组版)
# ==========================================

def apply_layout_illustrator(input_pdf, output_pdf, layout_instruction, refined_detections, box_size_mm, margin_mm=0.0):
    input_pdf_abs = os.path.abspath(input_pdf).replace("\\", "/")
    output_pdf_abs = os.path.abspath(output_pdf).replace("\\", "/")

    mm2pt = 2.834645
    front_w_mm, side_w_mm, panel_h_mm, flap_h_mm = box_size_mm['排版尺寸']
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
        cls_name = det['class_name']
        targets = det.get('assigned_targets', [])
        if not targets: continue

        box = det['box']
        xmin, ymin, xmax, ymax = box
        src_w = xmax - xmin
        src_h = ymax - ymin

        # 退化检测框 (宽或高<=0) 会让缩放算出 Infinity, 触发 Illustrator 'AOoC' 错误 -> 跳过
        if src_w <= 0 or src_h <= 0:
            print(f"[WARN] 跳过退化检测框 [{i}] {cls_name}: box={box}, src_w={src_w:.3f}, src_h={src_h:.3f}", flush=True)
            continue

        print(f"[排版] det[{i}] {cls_name}: box=[{xmin:.1f},{ymin:.1f},{xmax:.1f},{ymax:.1f}] "
              f"src_w={src_w:.1f} src_h={src_h:.1f} targets={targets}", flush=True)

        if '上摇盖' in cls_name:
            row_idx = 0
        elif '下摇盖' in cls_name:
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

        extractItems(sourceDoc.pageItems, bl, br, bt, bb, bw, bh);

        var sel = sourceDoc.selection;
        if (sel.length > 0) {{
            var tempSourceGroup = sourceDoc.groupItems.add();
            // 倒序遍历时必须放置在顶层，以保持原有图层遮挡关系
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

            # 🌟 核心功能：在此处对内容进行“洗刀版”处理并进行对齐
            jsx_dynamic_parts += f"""
                __step = "place det {i} -> col {col_idx} (fid {fid})";
                var dup = pastedGroup.duplicate();

                // 递归清洗附带的原有外侧刀版线 (依据无填充、有描边且尺寸大于等于提取框)
                cleanOriginalDieLine(dup, bw, bh);

                var target_w = {cell_w} - 2 * {margin_pt};
                var target_h = {cell_h} - 2 * {margin_pt};

                if (target_w <= 0) target_w = 1;
                if (target_h <= 0) target_h = 1;

                // 严格的等比缩放 (src 尺寸已在 Python 侧保证 > 0)
                var scale_percent = Math.min(target_w / {src_w}, target_h / {src_h}) * 100;
                if (!isFinite(scale_percent) || scale_percent <= 0) scale_percent = 100;
                dup.resize(scale_percent, scale_percent, true, true, true, true, true);

                // 获取真实物理边界并居中对齐
                var gBounds = dup.visibleBounds; 
                var gCenter_x = (gBounds[0] + gBounds[2]) / 2.0;
                var gCenter_y = (gBounds[1] + gBounds[3]) / 2.0;

                var cell_cx = {cell_left} + {cell_w} / 2.0;
                var cell_cy = {cell_top} - {cell_h} / 2.0;

                var align_dx = cell_cx - gCenter_x;
                var align_dy = cell_cy - gCenter_y;

                dup.translate(align_dx, align_dy);
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

    # 🌟 刀版线清理与提取框架
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

        // 新文档尺寸按排版网格(含边距)创建, 使 Illustrator 画布以网格为中心。
        // 若沿用源 PDF 尺寸, 源稿较小时画布中心靠近原点, 大网格画板会超出 ±16383 画布边界 -> 'AOoC'。
        __step = "create newDoc (gridW=" + gridW + ", gridH=" + gridH + ")";
        var docMargin = 50 * mm2pt;
        var newDoc = app.documents.add(sourceDoc.documentColorSpace, gridW + 2 * docMargin, gridH + 2 * docMargin);
        newDoc.artboards[0].artboardRect = [0, gridH, gridW, 0];

        app.activeDocument = newDoc;

        // --- 绘制新刀版线 ---
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
        // -------------------

        // 清理原刀版线函数：递归查找并删除对应特征路径
        // 判定 (满足其一即删)：
        //   0. 颜色: 描边为蓝/青 或 品红/洋红 (源稿刀版线常用色) -> 直接删, 不依赖尺寸 (含残段)
        //   A. 框线: 无填充 + 有描边，且沿宽或高任一方向跨越面板 >= 90% (整框/长折线)
        //   B. 细长线: 最短边 <= 4pt 且沿宽或高任一方向跨越面板 >= 60% (折线/压痕/出血, 含细长填充矩形)
        //   小尺寸的表格边框/装饰线因跨度小、且非蓝青色, 会被保留
        var DIE_SPAN = 0.9;
        var THIN_PT = 4.0;
        var THIN_SPAN = 0.6;
        // 刀版线描边颜色判定 (蓝/青色 或 品红/洋红色, 均为源稿刀版线常用色)
        // 品红刀版线: RGB(236,0,140) 之类; 需与红色品牌文字 RGB(237,28,36) 区分 (刀线 blue 高, 红字 blue 低)
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
        // 仅描边且描边为刀版线专色 -> 源稿刀版线 (排除有填充的同色图形)
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
        function isFrameDieLine(item, cw, ch, refW, refH) {{
            if (item.filled || !item.stroked) return false;
            return (cw >= refW * DIE_SPAN) || (ch >= refH * DIE_SPAN);
        }}
        function isThinDieLine(cw, ch, refW, refH) {{
            var thin = Math.min(cw, ch) <= THIN_PT;
            var spanLong = (cw >= refW * THIN_SPAN) || (ch >= refH * THIN_SPAN);
            return thin && spanLong;
        }}
        function cleanOriginalDieLine(item, refW, refH) {{
            if (!item) return;
            if (item.typename === "PathItem") {{
                try {{
                    if (strokeIsDieline(item)) {{ item.remove(); return; }}
                    var bounds = item.visibleBounds;
                    var cw = Math.abs(bounds[2] - bounds[0]);
                    var ch = Math.abs(bounds[1] - bounds[3]);
                    if (isFrameDieLine(item, cw, ch, refW, refH) || isThinDieLine(cw, ch, refW, refH)) {{
                        item.remove();
                    }}
                }} catch(e) {{}}
            }} else if (item.typename === "GroupItem") {{
                for (var i = item.pageItems.length - 1; i >= 0; i--) {{
                    cleanOriginalDieLine(item.pageItems[i], refW, refH);
                }}
            }} else if (item.typename === "CompoundPathItem") {{
                try {{
                    if (strokeIsDieline(item)) {{ item.remove(); return; }}
                    var bounds = item.visibleBounds;
                    var cw = Math.abs(bounds[2] - bounds[0]);
                    var ch = Math.abs(bounds[1] - bounds[3]);
                    var p0 = (item.pathItems && item.pathItems.length > 0) ? item.pathItems[0] : null;
                    var frameLike = p0 && !p0.filled && p0.stroked && ((cw >= refW * DIE_SPAN) || (ch >= refH * DIE_SPAN));
                    if (frameLike || isThinDieLine(cw, ch, refW, refH)) {{
                        item.remove();
                    }}
                }} catch(e) {{}}
            }}
        }}

        function extractItems(items, bl, br, bt, bb, bw, bh) {{
            var tol = 0;
            var pick_l = bl - tol, pick_r = br + tol;
            var pick_t = bt + tol, pick_b = bb - tol;

            for (var i = 0; i < items.length; i++) {{
                var item = items[i];
                if (item.hidden || item.locked || item.guides) continue;
                try {{
                    var bounds = item.visibleBounds; 
                    var l = bounds[0], t = bounds[1], r = bounds[2], b = bounds[3];
                    var itemW = Math.abs(r - l);
                    var itemH = Math.abs(t - b);

                    if (r < pick_l || l > pick_r || b > pick_t || t < pick_b) continue;

                    var isHuge = (itemW > bw * 1.2) || (itemH > bh * 1.2);

                    if (isHuge) {{
                        if (item.typename === "GroupItem") {{
                            extractItems(item.pageItems, bl, br, bt, bb, bw, bh);
                        }}
                        continue;
                    }}
                    // 仅当对象中心落在提取框内才选中, 避免抓到边界处属于邻块的内容 (防止串料/撑大包围盒)
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
            print(f"[ERR] Illustrator 渲染失败: JSX返回={result_str!r}, 输出文件存在={file_ok}, "
                  f"目标={output_pdf_abs}", flush=True)
            return None
        print(f"[OK] Illustrator 渲染成功: {output_pdf} (JSX={result_str!r})", flush=True)
        return output_pdf
    except Exception as e:
        print(f"[ERR] Python 侧调用 Illustrator 出错：{e}", flush=True)
        traceback.print_exc()
        return None


# ==========================================
# 3. 业务调度层
# ==========================================

def apply_layout_pure_python(pdf_path, layout_instruction, box_size_mm, detections, zoom_factor=2.0, resize=False,
                             margin_mm=0.0):
    pdf_path_abs = os.path.abspath(pdf_path)

    try:
        dims = box_size_mm["排版尺寸"] if isinstance(box_size_mm, dict) else box_size_mm
        L_pt, W_pt, H_pt, hFlap_pt = [mm_to_pt(x) for x in dims]
        src_doc = fitz.open(pdf_path_abs)
        src_page = src_doc[0]
    except Exception as e:
        print(f"❌ 无法读取PDF或解析尺寸: {e}")
        return None, None

    label_map = {
        "正唛内容": "front_content", "侧唛内容": "side_content",
        "正唛上摇盖": "front_flap_up", "侧唛上摇盖": "side_flap_up",
        "正唛下摇盖": "front_flap_down", "侧唛下摇盖": "side_flap_down",
    }

    refined_detections = []

    for i, det in enumerate(detections):
        label = det.get('class_name', '')
        x1, y1, x2, y2 = det["box"]
        ai_raw_rect = fitz.Rect(x1 / zoom_factor, y1 / zoom_factor, x2 / zoom_factor, y2 / zoom_factor)

        final_rect, is_precise = smart_crop_box_action(src_page, ai_raw_rect, detections, zoom_factor)

        ai_area = ai_raw_rect.width * ai_raw_rect.height
        final_area = final_rect.width * final_rect.height
        should_shrink = True
        if is_precise and (final_area < ai_area):
            should_shrink = False

        if should_shrink:
            w = final_rect.width
            h = final_rect.height
            min_side = min(w, h)
            ratio_margin = min_side * 0.04
            dynamic_margin = min(3.0, ratio_margin)
            if (w > dynamic_margin * 3) and (h > dynamic_margin * 3) and (dynamic_margin > 0.1):
                final_rect = fitz.Rect(
                    final_rect.x0 + dynamic_margin,
                    final_rect.y0 + dynamic_margin,
                    final_rect.x1 - dynamic_margin,
                    final_rect.y1 - dynamic_margin
                )

        key = None
        matched_k = None
        for k, v in label_map.items():
            if k in label:
                key = v
                matched_k = k
                break

        if key:
            new_det = det.copy()
            new_det['box'] = [final_rect.x0, final_rect.y0, final_rect.x1, final_rect.y1]
            new_det['target_key'] = key
            new_det['original_index'] = i
            new_det['assigned_targets'] = []
            new_det['normalized_class'] = matched_k
            refined_detections.append(new_det)

    class_to_dets = {}
    for det in refined_detections:
        cls = det.get('normalized_class') or det.get('class_name')
        if cls not in class_to_dets: class_to_dets[cls] = []
        class_to_dets[cls].append(det)

    for instr_label, target_ids in layout_instruction.items():
        if not target_ids: continue
        targets = [target_ids] if not isinstance(target_ids, list) else target_ids
        available_dets = class_to_dets.get(instr_label, [])
        if not available_dets: continue

        base_offset = 0
        if '上摇盖' in instr_label:
            base_offset = 100
        elif '下摇盖' in instr_label:
            base_offset = 200

        for i, fid in enumerate(targets):
            assigned_det = available_dets[i % len(available_dets)]
            actual_fid = fid + base_offset
            assigned_det['assigned_targets'].append(actual_fid)

    base_name_pre = os.path.splitext(os.path.basename(pdf_path_abs))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    final_pdf_path = os.path.join(output_dir, f"Layout_{base_name_pre}_{timestamp}.pdf")

    # 进入 Illustrator 引擎进行渲染
    if resize:
        result_pdf = apply_layout_illustrator(
            input_pdf=pdf_path_abs,
            output_pdf=final_pdf_path,
            layout_instruction=layout_instruction,
            refined_detections=refined_detections,
            box_size_mm=box_size_mm,
            margin_mm=margin_mm
        )

        if not result_pdf or not os.path.exists(final_pdf_path):
            print(f"[ERR] 排版未生成 PDF (Illustrator 渲染失败): 目标={final_pdf_path}, "
                  f"排版尺寸mm={dims}, 检测框数={len(refined_detections)}", flush=True)
            src_doc.close()
            return None, None

        final_preview_path = os.path.join(output_dir, f"{base_name_pre}_{timestamp}_final_preview.png")
        preview_doc = fitz.open(final_pdf_path)
        preview_page = preview_doc[0]
        preview_page.get_pixmap(matrix=fitz.Matrix(2, 2)).save(final_preview_path)
        preview_doc.close()

        src_doc.close()
        return final_pdf_path, final_preview_path
    else:
        src_doc.close()
        return None, None


if __name__ == "__main__":
    input_pdf = os.path.abspath(r"C:\Users\Administrator\Desktop\detection_pdf_purecode\e58d837d_1256b135_5-662Z_923759内外箱唛头_-_u14-0a70222e_75c9501e_lay_u14-0a70222e_313b2775_layout.pdf")

    # 检测框 (来自 /detect 的 result.final_boxes, zoom=2 图像坐标)
    # class_name 带 "第一/第二" 前缀不影响排版 (按子串匹配 正唛/侧唛 + 内容/上摇盖/下摇盖)
    # detections = [{'box': [75.96, 120.33, 345.75, 203.16], 'class_name': '第一正唛上摇盖'}, {'box': [1198.75, 112.16, 1655.04, 198.23], 'class_name': '第二正唛上摇盖'}, {'box': [35.36, 210.86, 679.72, 598.52], 'class_name': '正唛内容'}, {'box': [709.85, 250.15, 1147.28, 548.75], 'class_name': '侧唛内容'}]

    detections = [{'box': [39,495,729,1059], 'class_name': '正唛内容'}, {'box': [804,496,1491,1051], 'class_name': '侧唛内容'}]

    layout_instruction = {'正唛内容': [1, 3], '侧唛内容': [2, 4],
                          '正唛上摇盖': [1, 3], '侧唛上摇盖': [2, 4],
                          '正唛下摇盖': [1, 3], '侧唛下摇盖': [2, 4]}

    # 排版尺寸 [正唛宽, 侧唛宽, 主面高, 摇盖高] (mm), 取检测出的刀版尺寸保持原比例
    box_size_mm = {'排版尺寸': [650,410,260,205]}

    if os.path.exists(input_pdf):
        pdf_path, preview_path = apply_layout_pure_python(
            input_pdf,
            layout_instruction,
            box_size_mm,
            detections,
            zoom_factor=ZOOM_FACTOR,
            resize=True,
            margin_mm=0.0  # 默认距边为0（直接按原比例自适应）
        )
        print("执行完成！")