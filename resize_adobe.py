import os
import sys
import math
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
# 2. Illustrator 渲染引擎 (最终版：修复刀版线丢失 + 顶层图层保护)
# ==========================================
import os
import win32com.client

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

    def _row_of(cls_name):
        if '上摇盖' in cls_name: return 0
        if '下摇盖' in cls_name: return 2
        return 1

    # 🌟 第零步：消除竖直抖动。同一物理行的各面板本应共用同一上下边界，但 smart_crop
    #            会把每个面板吸附到略有差异的矢量边，导致相同面板高低不一（向上/向下偏移）。
    #            这里把同一行所有面板的检测框 y0/y1 统一为该行中位数，保证排版后竖直对齐一致。
    from statistics import median
    _row_y = {0: [], 1: [], 2: []}
    for det in refined_detections:
        if not det.get('assigned_targets'): continue
        ri = _row_of(det['class_name'])
        box = det['box']
        _row_y[ri].append((min(box[1], box[3]), max(box[1], box[3])))
    _row_unified = {}
    for ri, vals in _row_y.items():
        if vals:
            _row_unified[ri] = (median(sorted(v[0] for v in vals)),
                                median(sorted(v[1] for v in vals)))
    for det in refined_detections:
        if not det.get('assigned_targets'): continue
        ri = _row_of(det['class_name'])
        if ri in _row_unified:
            box = det['box']
            y0, y1 = _row_unified[ri]
            # 保持原有 y 方向次序（box 通常为 y0<y1）
            if box[1] <= box[3]:
                box[1], box[3] = y0, y1
            else:
                box[1], box[3] = y1, y0

    # 🌟 第一步：在 Python 层预计算“行级统一缩放率”，确保同一行高度绝对一致，防止断层
    row_scales = {0: float('inf'), 1: float('inf'), 2: float('inf')}
    for det in refined_detections:
        cls_name = det['class_name']
        targets = det.get('assigned_targets', [])
        if not targets: continue
        
        box = det['box']
        bw_pt = abs(box[2] - box[0])
        bh_pt = abs(box[3] - box[1])
        
        if '上摇盖' in cls_name: r_idx = 0
        elif '下摇盖' in cls_name: r_idx = 2
        else: r_idx = 1
        
        for fid in targets:
            c_idx = (fid % 100) - 1
            target_w = colWidths[c_idx] - 2 * margin_pt
            target_h = rowHeights[r_idx] - 2 * margin_pt
            if target_w <= 0: target_w = 1
            if target_h <= 0: target_h = 1
            
            if bw_pt > 0 and bh_pt > 0:
                scale = min(target_w / bw_pt, target_h / bh_pt)
                if scale < row_scales[r_idx]:
                    row_scales[r_idx] = scale
                    
    for r in row_scales:
        if row_scales[r] == float('inf'): row_scales[r] = 1.0

    # 🔍 诊断：打印每个面板的检测框尺寸 vs 目标格子尺寸，以及最终行缩放率
    print("=" * 60)
    print("[Layout 诊断] 格子尺寸(pt):", [round(w, 1) for w in colWidths], "x", [round(h, 1) for h in rowHeights])
    for det in refined_detections:
        targets = det.get('assigned_targets', [])
        if not targets:
            print(f"  [无目标] {det.get('class_name')} -> assigned_targets 为空，不会被缩放/排版")
            continue
        box = det['box']
        bw_pt = abs(box[2] - box[0]); bh_pt = abs(box[3] - box[1])
        if '上摇盖' in det['class_name']: r_idx = 0
        elif '下摇盖' in det['class_name']: r_idx = 2
        else: r_idx = 1
        print(f"  {det['class_name']} FID={targets} 检测框={bw_pt:.1f}x{bh_pt:.1f}pt -> 行{r_idx}")
    print("[Layout 诊断] 最终行缩放率 row_scales:", {k: round(v, 4) for k, v in row_scales.items()})
    print("=" * 60)

    layout_log_path = os.path.abspath(os.path.join(output_dir, "layout_debug_log.txt")).replace("\\", "/")

    jsx_copy_logic = """
        debugStep = "1. 回源文档创建全局锚点并执行唯一一次复制";
        app.activeDocument = sourceDoc;
        try { app.executeMenuCommand('unlockAll'); } catch(e){}
        try { app.executeMenuCommand('showAll'); } catch(e){}
        
        var tempLayer = sourceDoc.layers.add();
        var globalAnchor = tempLayer.pathItems.rectangle(abRect[1], abRect[0], Math.abs(abRect[2]-abRect[0]), Math.abs(abRect[1]-abRect[3]));
        globalAnchor.name = "GLOBAL_ANCHOR";
        globalAnchor.filled = false; globalAnchor.stroked = false;
        
        sourceDoc.selection = null;
        app.executeMenuCommand('selectall');
        app.redraw(); 
        
        var copySuccess = false;
        for (var retry = 0; retry < 3; retry++) {
            try { app.copy(); copySuccess = true; break; } catch (e) { $.sleep(500); }
        }
        if (!copySuccess) throw new Error("剪贴板被系统占用，全局复制失败！");
    """

    jsx_dynamic_parts = ""

    for i, det in enumerate(refined_detections):
        cls_name = det['class_name']
        targets = det.get('assigned_targets', [])
        if not targets: continue

        box = det['box']
        xmin, ymin, xmax, ymax = box

        if '上摇盖' in cls_name: row_idx = 0
        elif '下摇盖' in cls_name: row_idx = 2
        else: row_idx = 1
        
        unified_scale_percent = row_scales[row_idx] * 100.0

        for fid in targets:
            col_idx = (fid % 100) - 1
            cell_w = colWidths[col_idx]
            cell_h = rowHeights[row_idx]
            cell_left = sum(colWidths[:col_idx])
            cell_top = gridH - sum(rowHeights[:row_idx])

            jsx_dynamic_parts += f"""
                globalDebugLog.push("\\n---------------------------------------");
                globalDebugLog.push("👉 开始处理面板: {cls_name} (目标格子 FID: {fid})");
                debugStep = "2. 新文档粘贴 (分类: {cls_name}, FID: {fid})";
                app.activeDocument = newDoc;
                newDoc.selection = null;
                
                var pasteSuccess = false;
                for (var retry = 0; retry < 3; retry++) {{
                    try {{ app.paste(); pasteSuccess = true; break; }} catch (e) {{ $.sleep(300); }}
                }}
                if (!pasteSuccess) throw new Error("无法从剪贴板粘贴数据");
                app.redraw();
                
                debugStep = "3. 底层 DOM 安全编组";
                var targetWrapper = newDoc.activeLayer.groupItems.add();
                targetWrapper.name = "CONTENT_WRAPPER";
                var sel = newDoc.selection;
                if (sel && sel.length > 0) {{
                    for (var si = sel.length - 1; si >= 0; si--) {{
                        try {{ sel[si].moveToBeginning(targetWrapper); }} catch(e){{}}
                    }}
                }}
                
                debugStep = "4. 对齐全局锚点";
                var gAnchor = null;
                for (var pi=0; pi<targetWrapper.pageItems.length; pi++) {{
                    if (targetWrapper.pageItems[pi].name === "GLOBAL_ANCHOR") {{ gAnchor = targetWrapper.pageItems[pi]; break; }}
                }}
                if (gAnchor) {{
                    targetWrapper.translate(abRect[0] - gAnchor.position[0], abRect[1] - gAnchor.position[1]);
                    gAnchor.remove();
                }}
                
                debugStep = "5. 预处理 (解锁/文本转曲)";
                var bl = abRect[0] + {xmin}, bt = abRect[1] - {ymin};
                var br = abRect[0] + {xmax}, bb = abRect[1] - {ymax};
                var bw = Math.abs(br - bl), bh = Math.abs(bt - bb);
                var box_l = Math.min(bl, br), box_r = Math.max(bl, br);
                var box_t = Math.max(bt, bb), box_b = Math.min(bt, bb);

                unlockAndUnhide(targetWrapper);
                convertTextToOutlines(targetWrapper);

                debugStep = "6. 删除完全越界的其它面板内容";
                deleteFullyOutside(targetWrapper, box_l, box_r, box_t, box_b);

                debugStep = "6.5 删除源稿自带边框/刀版线 (防止压线越界)";
                removeFramesAndDieLines(targetWrapper, bw, bh);

                debugStep = "7. 判定区域是否有有效内容";
                var hasContent = hasInkInBox(targetWrapper, box_l, box_r, box_t, box_b);

                if (!hasContent) {{
                    try {{ targetWrapper.remove(); }} catch(e) {{}}
                    globalDebugLog.push("  [剔除] 该区域无有效像素内容");
                }} else {{
                    debugStep = "8. 裁剪到检测框 (整体保形, 不修改内容)";
                    var clipRect = clipToBox(targetWrapper, box_l, box_t, bw, bh);

                    var scale_percent = {unified_scale_percent};
                    globalDebugLog.push("  [缩放数据] 整体等比缩放: " + scale_percent.toFixed(2) + "% (内容保持原样, 仅整体缩放+定位)");

                    var cell_cx = {cell_left} + {cell_w} / 2.0, cell_cy = {cell_top} - {cell_h} / 2.0;

                    // 整组等比缩放(关于自身中心)，再按裁剪框中心对齐到目标格子中心
                    targetWrapper.resize(scale_percent, scale_percent, true, true, true, true, true, Transformation.CENTER);
                    var rb = clipRect.geometricBounds;
                    targetWrapper.translate(cell_cx - ((rb[0] + rb[2]) / 2.0), cell_cy - ((rb[1] + rb[3]) / 2.0));

                    // 烘焙裁剪: 去掉剪切蒙版对象, 同时按目标格子边界删除越界残余(防溢出)
                    var bake_cl = {cell_left}, bake_cr = {cell_left} + {cell_w};
                    var bake_ct = {cell_top}, bake_cb = {cell_top} - {cell_h};
                    bakeClip(targetWrapper, bake_cl, bake_cr, bake_ct, bake_cb);
                }}
                newDoc.selection = null;
            """

    try:
        try:
            app = win32com.client.GetActiveObject("Illustrator.Application")
        except Exception:
            app = win32com.client.Dispatch("Illustrator.Application")
    except Exception as e:
        print(f"❌ 无法启动 Illustrator: {e}")
        return None, None

    jsx_code = f"""
    var debugStep = "0. 初始化执行环境";
    var globalDebugLog = [];
    try {{
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
        var oldPasteLayers = false;
        try {{ oldPasteLayers = app.preferences.getBooleanPreference("pasteRemembersLayers"); }} catch(e){{}}
        app.preferences.setBooleanPreference("pasteRemembersLayers", false);

        var mm2pt = 2.834645;
        var colWidths = [{frontW}, {sideW}, {frontW}, {sideW}];
        var rowHeights = [{flapH}, {panelH}, {flapH}];
        var gridW = {gridW};
        var gridH = {gridH};

        var fileToOpen = new File("{input_pdf_abs}");
        var sourceDoc = app.open(fileToOpen);
        var abRect = sourceDoc.artboards[0].artboardRect; 

        var newDoc = app.documents.add(sourceDoc.documentColorSpace, sourceDoc.width, sourceDoc.height);
        newDoc.artboards[0].artboardRect = [0, gridH, gridW, 0];
        app.activeDocument = newDoc;
        
        // 创建主内容图层，确保粘贴的内容都在这里
        var contentLayer = newDoc.layers.add();
        contentLayer.name = "排版内容";

        function unlockAndUnhide(container) {{
            for (var i = 0; i < container.pageItems.length; i++) {{
                var item = container.pageItems[i];
                try {{ item.locked = false; }} catch(e){{}}
                try {{ item.hidden = false; }} catch(e){{}}
                if (item.typename === "GroupItem") unlockAndUnhide(item);
            }}
        }}

        function convertTextToOutlines(container) {{
            for (var i = container.pageItems.length - 1; i >= 0; i--) {{
                var item = container.pageItems[i];
                if (item.typename === "GroupItem") {{
                    convertTextToOutlines(item);
                }} else if (item.typename === "TextFrame") {{
                    try {{ item.createOutline(); }} catch(e){{}}
                }}
            }}
        }}

        function hasRealInk(container) {{
            for (var i = 0; i < container.pageItems.length; i++) {{
                var item = container.pageItems[i];
                if (item.name === "AI_ANCHOR_BOX" || item.name === "BAND_ANCHOR") continue;
                if (item.typename === "PathItem") {{ if (item.filled || item.stroked) return true; }} 
                else if (item.typename === "CompoundPathItem") {{
                    if (item.pathItems) {{
                        for(var p=0; p<item.pathItems.length; p++) {{
                            if (item.pathItems[p].filled || item.pathItems[p].stroked) return true;
                        }}
                    }}
                }} else if (item.typename === "TextFrame") {{
                    if (item.contents && item.contents.replace(/\\s+/g, '') !== '') return true;
                }} else if (item.typename === "PlacedItem" || item.typename === "RasterItem") {{ return true; }} 
                else if (item.typename === "GroupItem") {{ if (hasRealInk(item)) return true; }}
            }}
            return false;
        }}

        // 🌟 整体保形方案：仅删除「完全位于检测框之外」的其它面板内容，绝不修改保留内容的几何
        // box_t = 上边(较大 y)，box_b = 下边(较小 y)
        function deleteFullyOutside(container, box_l, box_r, box_t, box_b) {{
            var tol = 0.5;
            for (var i = container.pageItems.length - 1; i >= 0; i--) {{
                var item = container.pageItems[i];
                if (item.name === "GLOBAL_ANCHOR" || item.name === "CLIP_BOX") continue;

                if (item.typename === "GroupItem") {{
                    deleteFullyOutside(item, box_l, box_r, box_t, box_b);
                    if (item.pageItems.length === 0) {{ try {{ item.remove(); }} catch(e){{}} }}
                    continue;
                }}
                try {{
                    var gb = item.geometricBounds;
                    if (!isFinite(gb[0])) {{ item.remove(); continue; }}
                    var i_l = Math.min(gb[0], gb[2]), i_r = Math.max(gb[0], gb[2]);
                    var i_b = Math.min(gb[1], gb[3]), i_t = Math.max(gb[1], gb[3]);
                    // 完全在框外才删除；跨越边界的内容保留，交给裁剪蒙版处理
                    if (i_r < box_l - tol || i_l > box_r + tol || i_b > box_t + tol || i_t < box_b - tol) {{
                        item.remove();
                    }}
                }} catch(e) {{}}
            }}
        }}

        // 判定检测框区域内是否存在有效内容（有则排版，无则剔除该面板）
        function hasInkInBox(container, box_l, box_r, box_t, box_b) {{
            for (var i = 0; i < container.pageItems.length; i++) {{
                var item = container.pageItems[i];
                if (item.name === "GLOBAL_ANCHOR" || item.name === "CLIP_BOX") continue;
                if (item.typename === "GroupItem") {{
                    if (hasInkInBox(item, box_l, box_r, box_t, box_b)) return true;
                    continue;
                }}
                var ink = false;
                if (item.typename === "PathItem") {{ ink = (item.filled || item.stroked); }}
                else if (item.typename === "CompoundPathItem") {{
                    if (item.pathItems) {{
                        for (var p = 0; p < item.pathItems.length; p++) {{
                            if (item.pathItems[p].filled || item.pathItems[p].stroked) {{ ink = true; break; }}
                        }}
                    }}
                }} else if (item.typename === "TextFrame") {{
                    ink = (item.contents && item.contents.replace(/\\s+/g, '') !== '');
                }} else if (item.typename === "PlacedItem" || item.typename === "RasterItem") {{ ink = true; }}
                if (!ink) continue;
                try {{
                    var gb = item.geometricBounds;
                    var i_l = Math.min(gb[0], gb[2]), i_r = Math.max(gb[0], gb[2]);
                    var i_b = Math.min(gb[1], gb[3]), i_t = Math.max(gb[1], gb[3]);
                    if (i_r >= box_l && i_l <= box_r && i_t >= box_b && i_b <= box_t) return true;
                }} catch(e) {{}}
            }}
            return false;
        }}

        // 判定描边颜色是否为「源稿刀版线」常用的蓝/青色 (例如 RGB 0,174,239 / CMYK 高青低品低黄)。
        function colorIsBlue(col) {{
            if (!col) return false;
            try {{
                var tn = col.typename;
                if (tn === "RGBColor") {{
                    return (col.blue > 100 && col.red < 130 && (col.blue - col.red) > 40);
                }} else if (tn === "CMYKColor") {{
                    return (col.cyan > 45 && col.magenta < 45 && col.yellow < 45);
                }} else if (tn === "SpotColor") {{
                    return colorIsBlue(col.spot.color);
                }}
            }} catch(e) {{}}
            return false;
        }}
        // 仅描边、且描边为蓝/青色 -> 视为源稿刀版线 (排除有填充的蓝色图形)。
        function strokeIsDieline(item) {{
            try {{
                if (item.typename === "PathItem") {{
                    return (!item.filled && item.stroked && colorIsBlue(item.strokeColor));
                }} else if (item.typename === "CompoundPathItem" && item.pathItems && item.pathItems.length > 0) {{
                    var noFill = true, blueStroke = false;
                    for (var cp = 0; cp < item.pathItems.length; cp++) {{
                        if (item.pathItems[cp].filled) noFill = false;
                        if (item.pathItems[cp].stroked && colorIsBlue(item.pathItems[cp].strokeColor)) blueStroke = true;
                    }}
                    return (noFill && blueStroke);
                }}
            }} catch(e) {{}}
            return false;
        }}

        // 删除源稿自带的「面板边框 / 刀版线」等空描边图元（无填充、仅描边、尺寸≈整个面板，
        // 或贯穿整行/整列的细线，或描边为蓝/青色的刀版线）。这些不是图形内容，排版引擎会另绘
        // 干净的刀版线；若保留，它们会在格子边界处与刀版线叠加、并越界到相邻格子，形成杂乱的
        // 线条/重影。仅整体删除这类图元，不修改任何保留内容的几何。
        function removeFramesAndDieLines(container, box_w, box_h) {{
            for (var i = container.pageItems.length - 1; i >= 0; i--) {{
                var item = container.pageItems[i];
                if (item.name === "GLOBAL_ANCHOR" || item.name === "CLIP_BOX") continue;
                if (item.typename === "GroupItem") {{
                    removeFramesAndDieLines(item, box_w, box_h);
                    if (item.pageItems.length === 0) {{ try {{ item.remove(); }} catch(e){{}} }}
                    continue;
                }}
                try {{
                    var emptyStroke = false;
                    if (item.typename === "PathItem") {{
                        emptyStroke = (!item.filled && item.stroked);
                    }} else if (item.typename === "CompoundPathItem") {{
                        if (item.pathItems && item.pathItems.length > 0) {{
                            var noFill = true, gotStroke = false;
                            for (var cp = 0; cp < item.pathItems.length; cp++) {{
                                if (item.pathItems[cp].filled) noFill = false;
                                if (item.pathItems[cp].stroked) gotStroke = true;
                            }}
                            emptyStroke = (noFill && gotStroke);
                        }}
                    }}
                    if (!emptyStroke) continue;
                    // 颜色优先: 凡是蓝/青色描边(源稿刀版线)一律删除, 不依赖尺寸判断
                    if (strokeIsDieline(item)) {{ item.remove(); continue; }}
                    var gb = item.geometricBounds;
                    var w = Math.abs(gb[2] - gb[0]), h = Math.abs(gb[1] - gb[3]);
                    var isFramingBox = (w >= box_w * 0.9 && h >= box_h * 0.9);
                    var isSpanLineH = (w >= box_w * 0.9 && h <= 3.0);
                    var isSpanLineV = (h >= box_h * 0.9 && w <= 3.0);
                    if (isFramingBox || isSpanLineH || isSpanLineV) {{ item.remove(); }}
                }} catch(e) {{}}
            }}
        }}

        // 用一个矩形蒙版把整个面板内容裁剪到检测框，几何零修改（超出部分仅被遮挡，不删点不变形）
        function clipToBox(wrapper, left, top, w, h) {{
            var clipRect = wrapper.pathItems.rectangle(top, left, w, h);
            clipRect.name = "CLIP_BOX";
            clipRect.filled = false;
            clipRect.stroked = false;
            try {{ clipRect.zOrder(ZOrderMethod.BRINGTOFRONT); }} catch(e) {{}}
            clipRect.clipping = true;
            wrapper.clipped = true;
            return clipRect;
        }}

        // 检测面板内是否存在「跨越格子边界」(部分在内、部分在外) 的图元 -> 这类内容必须靠
        // 蒙版裁剪, 不能简单删除整块 (否则会像满版黑条那样溢出到相邻格)。
        function panelOverflows(container, cl, cr, ct, cb, tol) {{
            for (var i = 0; i < container.pageItems.length; i++) {{
                var it = container.pageItems[i];
                if (it.name === "GLOBAL_ANCHOR" || it.name === "CLIP_BOX") continue;
                if (it.typename === "GroupItem") {{
                    if (panelOverflows(it, cl, cr, ct, cb, tol)) return true;
                    continue;
                }}
                try {{
                    var gb = it.geometricBounds;  // [left, top, right, bottom], top>bottom
                    var i_l = Math.min(gb[0], gb[2]), i_r = Math.max(gb[0], gb[2]);
                    var i_b = Math.min(gb[1], gb[3]), i_t = Math.max(gb[1], gb[3]);
                    var insideArea = (i_r > cl && i_l < cr && i_t > cb && i_b < ct);
                    var pokesOut = (i_l < cl - tol || i_r > cr + tol || i_b < cb - tol || i_t > ct + tol);
                    if (insideArea && pokesOut) return true;  // 跨界 -> 需要蒙版
                }} catch(e) {{}}
            }}
            return false;
        }}

        // 烘焙裁剪(去蒙版): 在「内容未跨越格子边界」时, 解除剪切关系并删除裁剪矩形, 使输出不
        // 残留剪切组/蒙版对象; 若存在满版/跨界内容(如满版黑条), 则保留蒙版以防溢出。
        // 描边/文字等几何零修改, 完整保留。cl/cr = 左右边界, ct/cb = 上下边界 (ct 为较大 y)。
        function bakeClip(wrapper, cl, cr, ct, cb) {{
            // 有跨界内容 -> 保留蒙版裁剪, 不烘焙 (正确性优先)
            if (panelOverflows(wrapper, cl, cr, ct, cb, 1.0)) {{
                globalDebugLog.push("  [烘焙裁剪] 检测到跨界内容, 保留蒙版以防溢出");
                return;
            }}
            try {{ wrapper.clipped = false; }} catch(e) {{}}
            for (var i = wrapper.pageItems.length - 1; i >= 0; i--) {{
                try {{
                    if (wrapper.pageItems[i].name === "CLIP_BOX") {{ wrapper.pageItems[i].remove(); }}
                }} catch(e) {{}}
            }}
            deleteFullyOutside(wrapper, cl, cr, ct, cb);
        }}

        {jsx_copy_logic}
        {jsx_dynamic_parts}

        // 🌟 绘制顶部专属刀版线 (防遮挡) 🌟
        debugStep = "7. 绘制顶部刀版线";
        var dieLayer = newDoc.layers.add();
        dieLayer.name = "刀版线";
        var dieColor = (newDoc.documentColorSpace == DocumentColorSpace.CMYK) ? new CMYKColor() : new RGBColor();
        if (newDoc.documentColorSpace == DocumentColorSpace.CMYK) {{ dieColor.magenta = 100; }} 
        else {{ dieColor.red = 255; dieColor.blue = 0; dieColor.green = 255; }} // RGB 品红

        var currentY = gridH;
        for (var r = 0; r < 3; r++) {{
            var currentX = 0;
            for (var c = 0; c < 4; c++) {{
                var w = colWidths[c];
                var h = rowHeights[r];
                if (w <= 0 || h <= 0) {{ currentX += w; continue; }} // 跳过塌缩行/列(如2刀版线无摇盖)
                var panel = dieLayer.pathItems.rectangle(currentY, currentX, w, h);
                panel.filled = false; panel.stroked = true; panel.strokeColor = dieColor; panel.strokeWidth = 1.0; 
                panel.locked = true; // 锁定刀版线防止后续误触
                currentX += w;
            }}
            currentY -= rowHeights[r];
        }}

        // 🔍 诊断：把排版日志写入文件，便于排查"超刀版线/未缩放"问题
        try {{
            var _logFile = new File("{layout_log_path}");
            _logFile.encoding = "UTF-8";
            _logFile.open("w");
            _logFile.write("格子尺寸(pt) colWidths=[" + colWidths.join(",") + "] rowHeights=[" + rowHeights.join(",") + "]\\n");
            _logFile.write(globalDebugLog.join("\\n"));
            _logFile.close();
        }} catch(e) {{}}

        debugStep = "8. 保存新文档与清理缓存";
        var margin = 50 * mm2pt; 
        newDoc.artboards[0].artboardRect = [-margin, gridH + margin, gridW + margin, -margin];

        var destFile = new File("{output_pdf_abs}");
        var saveOpts = new PDFSaveOptions();
        saveOpts.preserveEditability = true; 

        newDoc.saveAs(destFile, saveOpts);
        newDoc.close(SaveOptions.DONOTSAVECHANGES);
        sourceDoc.close(SaveOptions.DONOTSAVECHANGES);
        
        try {{ app.preferences.setBooleanPreference("pasteRemembersLayers", oldPasteLayers); }} catch(e){{}}
        app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;

        globalDebugLog.join("\\n"); 

    }} catch(e) {{
        try {{ app.preferences.setBooleanPreference("pasteRemembersLayers", oldPasteLayers); }} catch(err){{}}
        app.userInteractionLevel = UserInteractionLevel.DISPLAYALERTS;
        "JSX Error | 阶段: [" + debugStep + "] | 信息: " + e.message + " | 代码行号: " + e.line;
    }}
    """

    try:
        result = app.DoJavaScript(jsx_code)
        if isinstance(result, str) and result.startswith("JSX Error"):
            print("\n" + "!"*50)
            print(f"❌ Illustrator JSX 内部报错详情:\n{result}")
            print("!"*50 + "\n")
            return None, None
        return output_pdf, None
    except Exception as e:
        print(f"❌ Python 侧调用错误：{e}")
        return None, None


# ==========================================
# 3. 业务调度层
# ==========================================

def apply_layout_pure_python(pdf_path, layout_instruction, box_size_mm, detections, zoom_factor=2.0, resize=False,
                             margin_mm=0.0):
    pdf_path_abs = os.path.abspath(pdf_path)

    # 自动适配「2 刀版线」(仅 正唛+侧唛、无摇盖): 折叠上下摇盖行, 按单行排版
    has_flap = any('摇盖' in (d.get('class_name', '') or '') for d in detections)
    if not has_flap:
        # 1) 指令里只保留正唛/侧唛内容, 去掉所有摇盖项
        layout_instruction = {k: v for k, v in layout_instruction.items() if '摇盖' not in k}
        # 2) 摇盖高度置 0 -> 上下摇盖行高度归零, 网格塌缩成单行
        _dims = box_size_mm["排版尺寸"] if isinstance(box_size_mm, dict) else box_size_mm
        _dims = list(_dims)
        while len(_dims) < 4:
            _dims.append(0.0)
        _dims = _dims[:4]
        _dims[3] = 0.0  # flap_h = 0
        box_size_mm = {"排版尺寸": _dims}
        print(f"[Layout] 检测到无摇盖 (2 刀版线), 启用单行排版 (摇盖高=0), 指令={list(layout_instruction.keys())}")

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

    if resize:
        result_pdf, _ = apply_layout_illustrator(
            input_pdf=pdf_path_abs,
            output_pdf=final_pdf_path,
            layout_instruction=layout_instruction,
            refined_detections=refined_detections,
            box_size_mm=box_size_mm,
            margin_mm=margin_mm
        )

        final_preview_path = os.path.join(output_dir, f"{base_name_pre}_{timestamp}_final_preview.png")
        if result_pdf and os.path.exists(result_pdf):
            preview_doc = fitz.open(result_pdf)
            preview_page = preview_doc[0]
            preview_page.get_pixmap(matrix=fitz.Matrix(2, 2)).save(final_preview_path)
            preview_doc.close()

        src_doc.close()
        return result_pdf, final_preview_path if result_pdf else None
    else:
        src_doc.close()
        return None, None


if __name__ == "__main__":
    input_pdf = os.path.abspath(r"C:\Users\18858\Desktop\detection_pdf_purecode\result\_nocolor\asdfz_clean_nocolor.pdf")

    detections = [{'box': [205.843994140624, 226.83422851562398, 521.941589355468, 265.593994140624], 'class_name': '侧唛上摇盖'}, {'box': [521.941589355468, 226.83422851562398, 972.479614257812, 266.56219482421795], 'class_name': '正唛上摇盖'}, {'box': [972.699584960938, 226.83422851562398, 1290.169799804688, 264.218017578124], 'class_name': '侧唛上摇盖'}, {'box': [1290.169799804688, 226.83422851562398, 1742.087646484376, 264.218017578124], 'class_name': '正唛上摇盖'}, {'box': [205.18179321289, 264.62219238281193, 521.941589355468, 581.377990722656], 'class_name': '侧唛内容'}, {'box': [521.941589355468, 264.62219238281193, 974.24560546875, 581.377990722656], 'class_name': '正唛内容'}, {'box': [973.273803710938, 264.218017578124, 1290.169799804688, 581.377990722656], 'class_name': '侧唛内容'}, {'box': [1290.169799804688, 264.218017578124, 1742.087646484376, 581.377990722656], 'class_name': '正唛内容'}, {'box': [205.518005371094, 582.554412841796, 522.277770996094, 621.31199645996], 'class_name': '侧唛下摇盖'}, {'box': [521.943786621094, 582.5503845214839, 973.364013671876, 621.31199645996], 'class_name': '正唛下摇盖'}, {'box': [974.24560546875, 582.554412841796, 1291.005981445312, 621.31199645996], 'class_name': '侧唛下摇盖'}, {'box': [1290.1181640625, 582.554412841796, 1742.087646484376, 621.31199645996], 'class_name': '正唛下摇盖'}]

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
        79.80,
        55.88,
        55.88,
        6.84
      ]
    }

    if os.path.exists(input_pdf):
        pdf_path, preview_path = apply_layout_pure_python(
            input_pdf,
            layout_instruction,
            box_size_mm,
            detections,
            zoom_factor=ZOOM_FACTOR,
            resize=True,
            margin_mm=0.0
        )
        if pdf_path:
            print("执行完成！排版文件已生成。")
        else:
            print("执行中断或失败，未生成排版文件。")