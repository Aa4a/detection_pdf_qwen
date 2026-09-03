import os
import fitz
import win32com.client
from pprint import pformat
import tempfile
import json
import time
import uuid


def _safe_remove(path, retries=5, delay=0.3):
    """删除临时文件, 带重试; Illustrator 可能短暂占用句柄导致 WinError 32。
    删不掉也不抛异常 (只是残留一个临时文件, 下次会用新文件名)。"""
    for _ in range(retries):
        try:
            if os.path.exists(path):
                os.remove(path)
            return True
        except Exception:
            time.sleep(delay)
    return False


# ==========================================
# 0. Adobe Illustrator 连接
# ==========================================
def get_illustrator_app():
    """连接 AI. 若 GetActiveObject 拿到的是已崩溃的实例, 自动回退到 Dispatch."""
    app = None
    try:
        app = win32com.client.GetActiveObject("Illustrator.Application")
        # 探测是否可用 (访问任意属性会抛 RPC 错误)
        _ = app.Path
    except Exception:
        app = None

    if app is None:
        app = win32com.client.Dispatch("Illustrator.Application")

    try:
        app.UserInteractionLevel = -1  # -1 = AiDontDisplayAlerts
    except Exception:
        pass
    return app


# ==========================================
# 1a. 从刀版"线段"重建网格 (核心修复)
# ==========================================
# 刀版图的版面往往是由开放的压痕线/折线绘制的, 而不是闭合矩形,
# 因此不能只靠"闭合矩形"识别 (否则会错误地命中标题栏表格)。
# 这里改为: 收集所有横/竖线段, 按位置加权聚类出主干网格线, 再用网格线交点重建格子。

def _cluster_weighted(items, tol):
    """items: list of (pos, weight, lo, hi)。按 pos 排序后, 相邻 (间距<=tol) 的合并,
    pos 用 weight 加权平均, lo/hi 取并集。"""
    items = sorted(items, key=lambda t: t[0])
    out = []
    for pos, w, lo, hi in items:
        if out and pos - out[-1]["pos"] <= tol:
            c = out[-1]
            tot = c["w"] + w
            c["pos"] = (c["pos"] * c["w"] + pos * w) / tot
            c["w"] = tot
            c["lo"] = min(c["lo"], lo)
            c["hi"] = max(c["hi"], hi)
        else:
            out.append({"pos": pos, "w": w, "lo": lo, "hi": hi})
    return out


def _reconstruct_grid_from_lines(rects, page_w, page_h):
    """从已转换为 PDF 坐标的所有路径外框 rects 中, 重建刀版网格。

    返回 final_rows = [[Rect, ...], ...] (每行已按 x 排序), 失败返回 None。
    至少需要 2 列才算有效 (1x1 外框不算成功, 应交后续闭合矩形/Qwen 回退)。
    """
    vseg, hseg = [], []
    # 线宽阈值放宽到 8pt: 部分刀版压痕线比 1pt 描边更粗, 过严会只剩外框
    for x0, y0, x1, y1 in rects:
        w, h = abs(x1 - x0), abs(y1 - y0)
        if w <= 8 and h > 15:                     # 竖线段
            vseg.append(((x0 + x1) / 2.0, h, min(y0, y1), max(y0, y1)))
        elif h <= 8 and w > 15:                   # 横线段
            hseg.append(((y0 + y1) / 2.0, w, min(x0, x1), max(x0, x1)))

    if len(vseg) < 2 or len(hseg) < 2:
        return None

    vcl = _cluster_weighted(vseg, 20.0)
    hcl = _cluster_weighted(hseg, 20.0)

    # 先严后松: 内部折线往往不贯穿整页, 过严会只剩外框 2x2 -> 1 格
    for v_frac, h_frac, tag in (
        (0.40, 0.65, "strict"),
        (0.22, 0.45, "relaxed"),
        (0.12, 0.30, "loose"),
    ):
        vlines = [c for c in vcl if (c["hi"] - c["lo"]) >= v_frac * page_h]
        hlines = [c for c in hcl if (c["hi"] - c["lo"]) >= h_frac * page_w]
        col_xs = sorted(c["pos"] for c in vlines)
        row_ys = sorted(c["pos"] for c in hlines)
        n_cols = len(col_xs) - 1
        n_rows = len(row_ys) - 1
        if n_cols < 2 or n_rows < 1:
            print(f"[AI] 线段重建({tag}): {len(col_xs)} 竖 x {len(row_ys)} 横 "
                  f"-> 列/行不足 (cols={max(0, n_cols)}, rows={max(0, n_rows)}), 跳过")
            continue

        # 刀版通常 2~4 列; 十几列多半是条码/文字装饰线误切
        if n_cols > 6:
            print(f"[AI] 线段重建({tag}): 列数过多 ({n_cols}>6), 跳过")
            continue

        rows = []
        for i in range(n_rows):
            y0, y1 = row_ys[i], row_ys[i + 1]
            # 丢弃高度过扁的伪行 (标题栏/装饰线夹出的缝)
            if (y1 - y0) < 0.04 * page_h:
                continue
            row = [fitz.Rect(col_xs[j], y0, col_xs[j + 1], y1)
                   for j in range(n_cols)]
            rows.append(row)

        if not rows or max(len(r) for r in rows) < 2:
            print(f"[AI] 线段重建({tag}): 有效行不足, 跳过")
            continue

        # 单行结果若高度远小于页面, 多半是底部装饰条而非主面
        if len(rows) == 1:
            rh = max(c.height for c in rows[0])
            if rh < 0.25 * page_h:
                print(f"[AI] 线段重建({tag}): 单行高度过矮 "
                      f"({rh:.1f} < 0.25*page_h), 跳过")
                continue

        print(f"[AI] 线段重建网格({tag}): {len(col_xs)} 条竖线 x {len(row_ys)} 条横线 "
              f"-> {len(rows)} 行 x {n_cols} 列")
        return rows

    return None


def _rows_from_qwen_boxes(qwen_boxes, y_tol=25.0):
    """把 Qwen 检测框按 y 聚类成行, 供 build_12_grid 在 Adobe 失败/质量差时回退使用。"""
    if not qwen_boxes:
        return None
    # 粘合口通常很窄, 留给 _expand_row_to_4 过滤; 这里先去掉明显非版面小框
    rects = sorted((b["rect"] for b in qwen_boxes), key=lambda r: r.y0)
    rows = []
    cur = [rects[0]]
    for r in rects[1:]:
        # 同行: 顶边接近, 或垂直范围有明显重叠
        ref = cur[0]
        y_overlap = min(ref.y1, r.y1) - max(ref.y0, r.y0)
        same_row = abs(r.y0 - ref.y0) < y_tol or y_overlap > 0.35 * min(ref.height, r.height)
        if same_row:
            cur.append(r)
        else:
            rows.append(sorted(cur, key=lambda c: c.x0))
            cur = [r]
    rows.append(sorted(cur, key=lambda c: c.x0))
    # 至少要有一行能扩成 4 格 (>=2 cell)
    if not any(len(r) >= 2 for r in rows):
        return None
    print(f"[Qwen] 检测框分行: 共 {len(rows)} 行")
    for i, row in enumerate(rows):
        info = ", ".join(f"({c.x0:.1f},{c.y0:.1f},{c.width:.1f}x{c.height:.1f})" for c in row)
        print(f"     Row {i} ({len(row)} 格): {info}")
    return rows


def _grid_looks_degenerate(grid, page_w, page_h=None):
    """Adobe 线段网格常见失败形态。"""
    if not grid or len(grid) < 2:
        return True
    main = grid[1]
    widths = [max(1.0, c.width) for c in main]
    heights = [max(1.0, c.height) for c in main]
    if max(widths) > 0.45 * page_w:
        return True
    if max(widths) / min(widths) > 8.0:
        return True
    if main[-1].x1 > page_w * 1.05:
        return True
    # 主面行高度过矮 (装饰条/条码区被当成主面)
    if page_h and max(heights) < 0.25 * page_h:
        return True
    return False


def _grid_qwen_coverage_score(grid, qwen_boxes):
    """每个 Qwen 框被「单个网格格」覆盖的比例, 再取平均。

    注意: 旧算法用交集/min(cell,qwen), 碎小格完全落在大 Qwen 框内会虚高到 1.0。
    这里改为交集/qwen面积, 碎格只能拿到很小的覆盖分。
    """
    if not grid or not qwen_boxes:
        return 0.0
    cells = [c for row in grid for c in row]
    scores = []
    for gb in qwen_boxes:
        g_area = max(1.0, gb["rect"].get_area())
        best = 0.0
        for cell in cells:
            inter = cell & gb["rect"]
            ia = inter.get_area() if not inter.is_empty else 0.0
            if ia <= 0:
                continue
            best = max(best, ia / g_area)
        scores.append(best)
    return sum(scores) / len(scores) if scores else 0.0


def _prefer_better_grid(adobe_pack, qwen_pack, qwen_boxes, page_w, page_h=None):
    """在 Adobe / Qwen 两套 12 格之间选更合理的一套。

    adobe_pack / qwen_pack: (grid, synth, info) 或 (None, None, None)
    """
    a_grid, a_synth, a_info = adobe_pack
    g_grid, g_synth, g_info = qwen_pack

    if a_grid is None and g_grid is None:
        return None, None, None, "none"
    if a_grid is None:
        return g_grid, g_synth, g_info, "qwen-only"
    if g_grid is None:
        return a_grid, a_synth, a_info, "adobe-only"

    a_bad = _grid_looks_degenerate(a_grid, page_w, page_h)
    a_score = _grid_qwen_coverage_score(a_grid, qwen_boxes)
    g_score = _grid_qwen_coverage_score(g_grid, qwen_boxes)
    print(f"[Grid] 质量对比: Adobe cover={a_score:.2f} degenerate={a_bad}; "
          f"Qwen cover={g_score:.2f}")

    # Adobe 几何崩了, 或盖不住 Qwen 大框 (碎格典型) -> 用 Qwen
    if a_bad or g_score >= a_score + 0.12:
        return g_grid, g_synth, g_info, "qwen-better"
    return a_grid, a_synth, a_info, "adobe"


# ==========================================
# 1. 使用 Adobe Illustrator 识别刀版线候选
# ==========================================
def detect_dielines_with_ai(input_path):
    """
    用 Adobe Illustrator (JSX 注入) 极速读取 PathItems，识别刀版线候选矩形并分行。
    将原本的 N 次跨进程通信压缩为 1 次。
    """
    app = get_illustrator_app()
    abs_in = os.path.abspath(input_path)
    doc = app.Open(abs_in)

    try:
        page_w = float(doc.Width)
        page_h = float(doc.Height)
        
        # 1. 准备临时文件路径 (每次唯一, 避免残留锁/并发冲突)
        temp_dir = tempfile.gettempdir()
        json_path = os.path.join(temp_dir, f"ai_dielines_{uuid.uuid4().hex}.json").replace("\\", "/")

        # 2. 构建 JSX 脚本
        jsx_script = f"""
        try {{
            var doc = app.activeDocument;
            var abRect = doc.artboards[0].artboardRect;
            var abLeft = abRect[0], abTop = abRect[1], abRight = abRect[2], abBottom = abRect[3];
            
            var jsonFile = new File("{json_path}");
            jsonFile.open("w");
            jsonFile.encoding = "UTF-8";
            
            jsonFile.write('{{"artboard":[' + abLeft + ',' + abTop + ',' + abRight + ',' + abBottom + '],"paths":[');
            
            var isFirst = true;
            var pathCount = doc.pathItems.length;
            
            for (var i = 0; i < pathCount; i++) {{
                try {{
                    var item = doc.pathItems[i];
                    if (item.hidden || item.guides) continue;
                    
                    var gb = item.geometricBounds;
                    var pLeft = gb[0], pTop = gb[1], pRight = gb[2], pBottom = gb[3];
                    
                    if (!isFirst) jsonFile.write(',');
                    jsonFile.write('[' + pLeft + ',' + pTop + ',' + pRight + ',' + pBottom + ']');
                    isFirst = false;
                }} catch (e) {{
                    // 忽略无效或无边界的路径
                }}
            }}
            jsonFile.write(']}}');
            jsonFile.close();
        }} catch (err) {{
            // 全局捕获
        }}
        """

        print(f"[AI] 页面尺寸: {page_w:.1f} x {page_h:.1f} pt，正在注入 JSX 极速提取路径...")
        
        # 3. 执行 JSX
        app.DoJavaScript(jsx_script)

        # 4. 读取 JSON 数据
        if not os.path.exists(json_path):
            raise Exception("JSX 脚本未能成功生成 JSON 文件，可能遇到权限或AI内部错误。")
            
        with open(json_path, "r", encoding="utf-8") as f:
            ai_data = json.load(f)

        ab_left, ab_top, ab_right, ab_bottom = ai_data["artboard"]
        raw_paths = ai_data["paths"]
        
        print(f"[AI] 极速提取完成: 共 {len(raw_paths)} 个底层路径，开始 Python 逻辑过滤...")

        # AI 画板 (artboard) 往往内嵌在 PDF 页面 (mediabox) 内部, 四周留有出血/边距。
        # 直接用 artboard 原点换算会导致整体偏移, 因此这里按"画板在页面中居中"计算补偿量。
        ai_w = ab_right - ab_left
        ai_h = ab_top - ab_bottom
        try:
            _pdf_page = fitz.open(abs_in)[0]
            pdf_w, pdf_h = _pdf_page.rect.width, _pdf_page.rect.height
        except Exception:
            pdf_w, pdf_h = ai_w, ai_h
        off_x = (pdf_w - ai_w) / 2.0
        off_y = (pdf_h - ai_h) / 2.0
        if abs(off_x) > 0.5 or abs(off_y) > 0.5:
            print(f"[AI] 画板内嵌补偿: off_x={off_x:.1f}, off_y={off_y:.1f} "
                  f"(PDF {pdf_w:.1f}x{pdf_h:.1f} vs 画板 {ai_w:.1f}x{ai_h:.1f})")

        # 坐标转换函数 (AI 画板坐标 -> PDF 页面坐标, 含内嵌补偿)
        def ai_to_pdf_rect(left, top, right, bottom):
            return fitz.Rect(
                left - ab_left + off_x,
                ab_top - top + off_y,
                right - ab_left + off_x,
                ab_top - bottom + off_y,
            )

        # ------------------------------------------------------------------
        # 优先: 用"线段重建网格" (适配开放压痕线绘制的刀版图, 避免误命中标题栏表格)
        # ------------------------------------------------------------------
        all_rects = [tuple(ai_to_pdf_rect(*p)) for p in raw_paths]
        recon_rows = _reconstruct_grid_from_lines(all_rects, pdf_w, pdf_h)
        # 至少一行有 >=2 格才可采用; 1x1 外框会阻断后续回退, 直接丢弃
        if recon_rows and any(len(r) >= 2 for r in recon_rows):
            print(f"[AI] 采用线段重建结果: {len(recon_rows)} 行")
            for i, row in enumerate(recon_rows):
                info = ", ".join(f"({r.x0:.1f},{r.y0:.1f},{r.width:.1f}x{r.height:.1f})" for r in row)
                print(f"     Row {i} ({len(row)} 格): {info}")
            _safe_remove(json_path)
            return recon_rows, (page_w, page_h)
        if recon_rows:
            print(f"[AI] 线段重建结果列数不足 "
                  f"(max_cols={max(len(r) for r in recon_rows)}), 放弃并回退")

        print("[AI] 线段重建未成功, 回退到闭合矩形识别逻辑...")

        min_size = min(page_w, page_h) * 0.05
        page_area = page_w * page_h
        candidates = []

        # 5. 执行 Python 侧的精确尺寸过滤
        for p in raw_paths:
            left, top, right, bottom = p
            w = right - left
            h = top - bottom

            # 保持你原有的优秀过滤逻辑
            if w <= 0 or h <= 0:
                continue
            if max(w, h) < min_size:
                continue
            if w < 2 and h < 150:
                continue
            if h < 2 and w < 150:
                continue
            if (w * h) / page_area >= 0.80:
                continue
            if w >= page_w * 0.95 and h >= page_h * 0.95:
                continue

            candidates.append(ai_to_pdf_rect(left, top, right, bottom))

        print(f"[AI] 过滤后候选刀版矩形: {len(candidates)} 个")

        # 去重: 边界几乎相同的矩形
        def _dedupe(rects, tol=2.0):
            out = []
            for r in rects:
                dup = False
                for o in out:
                    if (abs(r.x0 - o.x0) < tol and abs(r.y0 - o.y0) < tol
                            and abs(r.x1 - o.x1) < tol and abs(r.y1 - o.y1) < tol):
                        dup = True
                        break
                if not dup:
                    out.append(r)
            return out

        before = len(candidates)
        candidates = _dedupe(candidates)
        print(f"[AI] 去重后: {len(candidates)} 个 (移除 {before - len(candidates)} 个重复)")

        # 按 y0 升序 -> 分行 (top 相近视为同一行)
        candidates.sort(key=lambda r: r.y0)
        rows = []
        if candidates:
            cur = [candidates[0]]
            for r in candidates[1:]:
                if r.y0 - cur[-1].y0 < 10:
                    cur.append(r)
                else:
                    rows.append(cur)
                    cur = [r]
            rows.append(cur)

        # 每行按 x 排序
        final_rows = []
        for row in rows:
            sorted_row = sorted(row, key=lambda r: r.x0)
            if sorted_row:
                final_rows.append(sorted_row)

        print(f"[AI] 分行结果: 共 {len(final_rows)} 行")
        for i, row in enumerate(final_rows):
            info = ", ".join(f"({r.x0:.1f},{r.y0:.1f},{r.width:.1f}x{r.height:.1f})" for r in row)
            print(f"     Row {i} ({len(row)} 格): {info}")

        # 清理临时文件 (删不掉不影响结果)
        _safe_remove(json_path)

        return final_rows, (page_w, page_h)
        
    finally:
        try:
            doc.Close(2)  # 2 = aiDoNotSaveChanges
        except Exception:
            pass


# ==========================================
# 2. 构建/补全 3x4 = 12 格网格
# ==========================================
# 规则: 摇盖高度 = 侧唛宽度 / 2
# 列布局: [主面, 侧唛, 主面, 侧唛]
# 行布局: [上摇盖, 主面, 下摇盖]

def _row_y(row):
    """返回行的真实 y 范围 (min y0, max y1), 不用平均值, 避免内嵌小 cell 干扰"""
    y0 = min(c.y0 for c in row)
    y1 = max(c.y1 for c in row)
    return y0, y1


def _clean_row(row, tol=2.0):
    """清理同行内噪音:
       1) 先识别外框 (包住 >=2 个其它 cell 的大矩形) 并移除
       2) 再移除剩余的"被更大 cell 包住"的内嵌小 cell
       3) 按 x0 聚类合并近似重复 (保留面积最大的)
    """
    if len(row) <= 1:
        return row

    def contains(b, a):
        return (b.x0 - tol <= a.x0 and a.x1 <= b.x1 + tol
                and b.y0 - tol <= a.y0 and a.y1 <= b.y1 + tol
                and b.get_area() > a.get_area() * 1.1)

    # 1) 外框: 包住 >= 2 个"有意义的 cell" (w>=50 且 h>=50, 排除噪音细线)
    MIN_LEAF = 50.0

    def is_leaf(c):
        return c.width >= MIN_LEAF and c.height >= MIN_LEAF

    outer_ids = set()
    for i, b in enumerate(row):
        leaf_contained = 0
        for j, a in enumerate(row):
            if j != i and is_leaf(a) and contains(b, a):
                leaf_contained += 1
                if leaf_contained >= 2:
                    break
        if leaf_contained >= 2:
            outer_ids.add(i)

    # 2) 在非外框中, 过滤被其它非外框 cell 包住的小 cell
    remaining_idx = [i for i in range(len(row)) if i not in outer_ids]
    kept_idx = []
    for i in remaining_idx:
        a = row[i]
        is_inner = False
        for j in remaining_idx:
            if i == j:
                continue
            if contains(row[j], a):
                is_inner = True
                break
        if not is_inner:
            kept_idx.append(i)
    kept = [row[i] for i in kept_idx]

    # 3) 按 x0 聚类, 合并近似重复 (针对那些差异在 15pt 以内的近似矩形)
    if len(kept) > 1:
        kept.sort(key=lambda c: c.x0)
        clusters = [[kept[0]]]
        for c in kept[1:]:
            last = clusters[-1][-1]
            if abs(c.x0 - last.x0) < 15.0 and abs(c.x1 - last.x1) < 15.0:
                clusters[-1].append(c)
            else:
                clusters.append([c])
        kept = [max(cl, key=lambda c: c.get_area()) for cl in clusters]

    return kept


def _expand_row_to_4(row):
    """
    将一行扩充为 4 格 [主面, 侧唛, 主面, 侧唛]。

    支持:
      - 过滤极窄的粘口 (糊口)
      - 3格补全: 通过测量相邻格子的空隙，推断缺失的版面并补全
      - 2格补全: 以间隙为侧唛宽度向外补全
    """
    if isinstance(row, _SynthRow) and row.synth_flags is not None:
        return list(row)[:4], row.synth_flags[:4], "fully-synthesized"

    row = sorted(row, key=lambda c: c.x0)

    # 1. 过滤粘口 (通常在最左或最右，宽度远小于主面，通常 < 150 pt)
    filtered_row = []
    for i, c in enumerate(row):
        is_glue_flap = False
        if len(row) >= 3:
            # 检查最左侧
            if i == 0 and c.width < 150 and c.width < row[1].width * 0.4:
                is_glue_flap = True
            # 检查最右侧
            if i == len(row) - 1 and c.width < 150 and c.width < row[-2].width * 0.4:
                is_glue_flap = True
        if not is_glue_flap:
            filtered_row.append(c)
    
    row = filtered_row

    # 2. 如果超过 4 格，按宽度取最大的 4 个
    if len(row) > 4:
        top4 = sorted(row, key=lambda c: c.width, reverse=True)[:4]
        row = sorted(top4, key=lambda c: c.x0)
        
    if len(row) == 4:
        return row, [False] * 4, "detected-4"

    # 3. 如果剩 3 格 (过滤掉粘口或 AI 漏检了 1 格)
    if len(row) == 3:
        r0, r1, r2 = row
        y0 = min(c.y0 for c in row)
        y1 = max(c.y1 for c in row)
        
        gap1 = r1.x0 - r0.x1
        gap2 = r2.x0 - r1.x1
        
        # 若第一格和第二格之间有较大空隙 -> 说明缺口在这里，补全！
        if gap1 > 20 and gap1 >= gap2:
            synth = fitz.Rect(r0.x1, y0, r1.x0, y1)
            return [r0, synth, r1, r2], [False, True, False, False], "expanded-3-to-4 (补全左侧缺口)"
            
        # 若第二格和第三格之间有较大空隙 -> 补全！
        elif gap2 > 20 and gap2 > gap1:
            synth = fitz.Rect(r1.x1, y0, r2.x0, y1)
            return [r0, r1, synth, r2], [False, False, True, False], "expanded-3-to-4 (补全右侧缺口)"
            
        else:
            # 中间没空隙，说明缺的是最外侧的侧唛。直接镜像中间侧唛的宽度在最右侧补一个
            w_guess = r1.width 
            synth = fitz.Rect(r2.x1, y0, r2.x1 + w_guess, y1)
            return [r0, r1, r2, synth], [False, False, False, True], "expanded-3-to-4 (补全最右侧)"

    # 4. 2 格补全逻辑
    if len(row) == 2:
        r0, r1 = row
        y0 = min(r0.y0, r1.y0)
        y1 = max(r0.y1, r1.y1)
        gap = r1.x0 - r0.x1          # 两块之间的水平间隙
        avg_w = (r0.width + r1.width) / 2.0

        # 情况 A: 两块之间有明显空隙 -> 空隙本身就是侧唛 (两块都是主面)
        if gap > avg_w * 0.2:
            side_w = gap
            col0 = fitz.Rect(r0.x0, y0, r0.x1, y1)              # 主面1 (检测)
            col1 = fitz.Rect(r0.x1, y0, r1.x0, y1)              # 侧唛1 (合成, 空隙)
            col2 = fitz.Rect(r1.x0, y0, r1.x1, y1)              # 主面2 (检测)
            col3 = fitz.Rect(r1.x1, y0, r1.x1 + side_w, y1)     # 侧唛2 (合成)
            return [col0, col1, col2, col3], [False, True, False, True], "expanded-2-to-4 (空隙=侧唛)"

        # 情况 B: 两块紧贴 (主面 + 侧唛), 按 [主面,侧唛,主面,侧唛] 向右重复铺成 4 格
        col0 = fitz.Rect(r0.x0, y0, r0.x1, y1)                  # 检测
        col1 = fitz.Rect(r1.x0, y0, r1.x1, y1)                  # 检测
        x = r1.x1
        col2 = fitz.Rect(x, y0, x + r0.width, y1)               # 合成 (复制 col0 宽)
        x += r0.width
        col3 = fitz.Rect(x, y0, x + r1.width, y1)               # 合成 (复制 col1 宽)
        return [col0, col1, col2, col3], [False, False, True, True], "expanded-2-adjacent-to-4 (主面+侧唛重复)"

    return None, None, f"unsupported-{len(row)}-cells"


def _virtual_cols_from_note(note):
    """返回该行扩展时纯属"向右重复补出"的虚拟列索引 (输出/绘制时应跳过)。
    仅 "主面+侧唛紧贴" 这种 2 格刀版会产生右侧 2 个虚拟列, 其余补全(空隙=侧唛、3补4等)
    补出的格子是真实版面的一部分, 不跳过。"""
    if note and "adjacent" in note:
        return {2, 3}
    return set()


def _filter_inner_rows(rows):
    """
    同时过滤:
      A) "外轮廓"行: cell 数 <=2 且包住了 >=3 个其他行 (如整个包装外框)
      B) "内嵌噪音"行: 被某个更大的、cell >=3 的行包住
    """
    def y_range(row):
        return min(c.y0 for c in row), max(c.y1 for c in row)

    ranges = [y_range(r) for r in rows]

    def _row_max_cell_area(row):
        return max(c.width * c.height for c in row)

    # A) 标记外轮廓行
    outline_idx = set()
    for i, (y0_i, y1_i) in enumerate(ranges):
        if len(rows[i]) > 2:
            continue
        area_i = _row_max_cell_area(rows[i])
        contained = 0
        wraps_structural = False  # 是否包住了另一组"主面级"的大结构
        for j, (y0_j, y1_j) in enumerate(ranges):
            if i == j:
                continue
            if y0_i - 2 <= y0_j and y1_j <= y1_i + 2:
                contained += 1
                # 被包住的行里若存在与本行 cell 同量级的大块, 才说明本行是真正的外框
                if _row_max_cell_area(rows[j]) >= 0.3 * area_i:
                    wraps_structural = True
        # 仅当它确实包住了另一组主面级结构时才算外轮廓;
        # 若它只是包住一堆 logo/文字/条码等小细节行, 那它本身就是主面行, 不能删
        if contained >= 3 and wraps_structural:
            outline_idx.add(i)

    # B) 判断内嵌时, parent 必须不是外轮廓且 cell 数 >= 3
    #    容差: 相对 parent 高度的 5% (处理 cell 顶/底稍超出 parent 边界的情况)
    keep_idx = []
    for i, (y0_i, y1_i) in enumerate(ranges):
        if i in outline_idx:
            continue
        is_inner = False
        for j, (y0_j, y1_j) in enumerate(ranges):
            if i == j or j in outline_idx:
                continue
            if len(rows[j]) < 3:
                continue
            h_i = y1_i - y0_i
            h_j = y1_j - y0_j
            y_tol = max(5.0, 0.05 * h_j)
            if h_j > h_i * 1.1 and y0_j - y_tol <= y0_i and y1_i <= y1_j + y_tol:
                is_inner = True
                break
        if not is_inner:
            keep_idx.append(i)

    keep = [rows[i] for i in keep_idx]
    if len(keep) != len(rows):
        outline_n = len(outline_idx)
        print(f"[Grid] 过滤: 外轮廓 {outline_n} 行 + 内嵌噪音 {len(rows) - len(keep) - outline_n} 行, "
              f"剩余 {len(keep)} 行")
    return keep


def _find_stacked_triplet(rows):
    """
    在 rows 中找 3 个纵向相邻的行 (第 i+1 行的 y0 接近第 i 行的 y1)。
    每行的 cell 个数 >=2 才算有效。
    找到则返回 [row_top, row_main, row_bottom], 否则 None。
    """
    valid = [r for r in rows if len(r) >= 2]
    if len(valid) < 3:
        return None

    valid.sort(key=lambda r: _row_y(r)[0])

    for i in range(len(valid) - 2):
        r0, r1, r2 = valid[i], valid[i + 1], valid[i + 2]
        y0_0, y1_0 = _row_y(r0)
        y0_1, y1_1 = _row_y(r1)
        y0_2, y1_2 = _row_y(r2)
        gap1 = abs(y0_1 - y1_0)
        gap2 = abs(y0_2 - y1_1)
        tol = max(10.0, 0.05 * (y1_1 - y0_1))
        if gap1 < tol and gap2 < tol:
            return [r0, r1, r2]
    return None


def _pick_main_row(rows):
    """选主面行: 每行尝试 _expand_row_to_4, 取能成功扩展且扩展后面积最大的那行"""
    scored = []
    for r in rows:
        cells, _, _ = _expand_row_to_4(r)
        if cells is None or len(cells) != 4:
            continue
        area = sum(c.width * c.height for c in cells)
        scored.append((r, area))
    if not scored:
        return None
    return max(scored, key=lambda x: x[1])[0]


class _SynthRow(list):
    """用于标记整行都是 Python 合成的行"""
    synth_flags = None  # 如果整行合成, 这里记录 [True]*4


def _find_stacked_pair_and_synthesize(rows):
    """
    找 2 行纵向堆叠并补全第 3 行 (规则: h_flap = w_side/2)。

    情况 A: 上摇盖 + 主面 (main 在下, flap 在上)  -> 合成下摇盖
    情况 B: 主面 + 下摇盖 (main 在上, flap 在下)  -> 合成上摇盖

    判断哪个是 main: 高度显著更大 (约 >= 1.5x flap)
    """
    valid = [r for r in rows if len(r) >= 2]
    if len(valid) < 2:
        return None
    valid.sort(key=lambda r: _row_y(r)[0])

    for i in range(len(valid) - 1):
        ra, rb = valid[i], valid[i + 1]
        y0a, y1a = _row_y(ra)
        y0b, y1b = _row_y(rb)
        gap = abs(y0b - y1a)
        ha = y1a - y0a
        hb = y1b - y0b
        tol = max(10.0, 0.05 * max(ha, hb))
        if gap >= tol:
            continue

        if ha > hb * 1.5:
            main_raw, bottom_raw = ra, rb
            need = "top"
        elif hb > ha * 1.5:
            top_raw, main_raw = ra, rb
            need = "bottom"
        else:
            continue

        # 扩展 main 到 4 格 (决定列结构)
        main_cells, main_flags, note = _expand_row_to_4(main_raw)
        if main_cells is None:
            continue
        col_xs = [main_cells[0].x0, main_cells[1].x0,
                  main_cells[2].x0, main_cells[3].x0, main_cells[3].x1]
        w_side = min(main_cells[0].width, main_cells[1].width)  # 规则: 侧唛 = 窄的
        h_flap = w_side / 2.0
        main_y0 = min(c.y0 for c in main_cells)
        main_y1 = max(c.y1 for c in main_cells)

        def make_row(y0, y1):
            return [fitz.Rect(col_xs[k], y0, col_xs[k + 1], y1) for k in range(4)]

        if need == "bottom":
            top_cells, top_flags, _ = _expand_row_to_4(top_raw)
            if top_cells is None:
                continue
            synth_row = _SynthRow(make_row(main_y1, main_y1 + h_flap))
            synth_row.synth_flags = [True] * 4
            print(f"[Grid] 找到 2 行堆叠 (上摇盖+主面), 合成下摇盖 (h_flap={h_flap:.1f}pt).")
            return [top_raw, main_raw, synth_row]
        else:
            bottom_cells, bottom_flags, _ = _expand_row_to_4(bottom_raw)
            if bottom_cells is None:
                continue
            synth_row = _SynthRow(make_row(main_y0 - h_flap, main_y0))
            synth_row.synth_flags = [True] * 4
            print(f"[Grid] 找到 2 行堆叠 (主面+下摇盖), 合成上摇盖 (h_flap={h_flap:.1f}pt).")
            return [synth_row, main_raw, bottom_raw]

    return None


def build_12_grid(rows):
    """
    从 AI 识别的若干行, 构建完整的 3x4 = 12 格网格。

    策略:
      1) 若能找到 3 行纵向堆叠 (上摇盖/主面/下摇盖结构完整) -> 直接用
      2) 否则以面积最大的行为"主面行", 上下摇盖按规则 h_flap=w_side/2 合成
      3) 每行若只有 2 格 (仅主面被检测到), 则自动补出 [主面,侧唛,主面,侧唛]

    返回:
        grid:       [[Rect x4] x3]
        synth_flag: [[bool  x4] x3]  True = 合成的
        info: dict
    """
    if not rows:
        return None, None, None

    # 先清理每行内的小内嵌 cell
    rows = [_clean_row(r) for r in rows if r]
    before_n = sum(len(r) for r in rows)
    # 清理完如果行变空就移除
    rows = [r for r in rows if r]
    after_n = sum(len(r) for r in rows)
    if before_n != after_n:
        print(f"[Grid] 清理行内内嵌 cell: {before_n} -> {after_n} 个 cell")

    rows = _filter_inner_rows(rows)

    # 优先: 找 3 行堆叠
    triplet = _find_stacked_triplet(rows)

    # 次优: 找 2 行堆叠 (top+main 或 main+bottom), 合成缺失的一行
    if triplet is None:
        triplet = _find_stacked_pair_and_synthesize(rows)

    if triplet is not None:
        print("[Grid] 找到 3 行纵向堆叠的刀版结构, 直接使用。")
        expanded = []
        synth = []
        notes = []
        for row in triplet:
            cells, flags, note = _expand_row_to_4(row)
            if cells is None:
                print(f"[Grid] 行扩展失败: {note}, 退化为单行合成模式。")
                triplet = None
                break
            expanded.append(cells)
            synth.append(flags)
            notes.append(note)
            print(f"       行扩展: {note}, 合成标志={flags}")
        if triplet is not None:
            main_row = expanded[1]
            w0, w1 = main_row[0].width, main_row[1].width
            # 规则: 侧唛是窄的那个, 主面是宽的 (h_flap = 侧唛宽 / 2)
            w_side = min(w0, w1)
            w_main = max(w0, w1)
            # 列布局: 若 col[0] 是侧唛 (narrower) 则 [侧,主,侧,主]; 否则 [主,侧,主,侧]
            col_pattern = ["侧唛", "主面", "侧唛", "主面"] if w0 <= w1 else ["主面", "侧唛", "主面", "侧唛"]
            info = {
                "w_main_pt": w_main,
                "w_side_pt": w_side,
                "h_panel_pt": main_row[0].height,
                "h_flap_pt": expanded[0][0].height,
                "h_flap_source": "detected",
                "col_pattern": col_pattern,
                "cols_to_skip": sorted(_virtual_cols_from_note(notes[1])),
            }
            return expanded, synth, info

    # 没找到堆叠 -> 以最大行为主面, 上下摇盖合成
    print("[Grid] 未找到 3 行堆叠结构, 采用'主面行 + 合成摇盖'模式。")
    main_raw = _pick_main_row(rows)
    if main_raw is None:
        return None, None, None

    main_cells, main_flags, note = _expand_row_to_4(main_raw)
    if main_cells is None:
        print(f"[Grid] 主面行扩展失败: {note}")
        return None, None, None
    print(f"[Grid] 主面行扩展: {note}")

    col_xs = [main_cells[i].x0 for i in range(4)] + [main_cells[3].x1]
    w0, w1 = main_cells[0].width, main_cells[1].width
    w_side = min(w0, w1)
    w_main = max(w0, w1)
    col_pattern = ["侧唛", "主面", "侧唛", "主面"] if w0 <= w1 else ["主面", "侧唛", "主面", "侧唛"]
    h_flap = w_side / 2.0
    main_y0 = min(c.y0 for c in main_cells)
    main_y1 = max(c.y1 for c in main_cells)

    def make_row(y0, y1):
        return [fitz.Rect(col_xs[i], y0, col_xs[i + 1], y1) for i in range(4)]

    top_row = make_row(main_y0 - h_flap, main_y0)
    bottom_row = make_row(main_y1, main_y1 + h_flap)

    grid = [top_row, main_cells, bottom_row]
    # 摇盖行全部合成; 主面行按扩展逻辑给出的合成标志
    synth = [[True] * 4, main_flags, [True] * 4]

    info = {
        "w_main_pt": w_main,
        "w_side_pt": w_side,
        "h_panel_pt": main_cells[0].height,
        "h_flap_pt": h_flap,
        "h_flap_source": "rule (w_side/2)",
        "col_pattern": col_pattern,
        "cols_to_skip": sorted(_virtual_cols_from_note(note)),
    }
    return grid, synth, info


# ==========================================
# 3. 可视化 12 格检测框
# ==========================================
ROW_LABELS = ["上摇盖", "主面", "下摇盖"]
COL_LABELS = ["主面", "侧唛", "主面", "侧唛"]

CELL_COLORS = {
    "主面":   (0.90, 0.20, 0.20),
    "侧唛":   (0.10, 0.55, 0.90),
    "上摇盖": (0.20, 0.70, 0.30),
    "下摇盖": (0.95, 0.60, 0.10),
}


def _cell_color(row_label, col_label):
    if row_label in ("上摇盖", "下摇盖"):
        return CELL_COLORS[row_label]
    return CELL_COLORS[col_label]


def _draw_dashed_rect(page, rect, color, width, dash_len=6.0):
    """PyMuPDF 低版本不支持 dashes 属性, 这里手动画四条虚线"""
    shape = page.new_shape()
    try:
        # 高版本 fitz 支持在 finish 里传 dashes
        shape.draw_rect(rect)
        shape.finish(color=color, width=width, dashes=f"[{dash_len} {dash_len}] 0")
        shape.commit()
    except Exception:
        # 回退: 用短线段拼接
        def dashed_line(p1, p2):
            x1, y1 = p1
            x2, y2 = p2
            dx, dy = x2 - x1, y2 - y1
            length = (dx * dx + dy * dy) ** 0.5
            if length == 0:
                return
            ux, uy = dx / length, dy / length
            n = int(length // (dash_len * 2))
            for k in range(n + 1):
                sx = x1 + ux * (k * dash_len * 2)
                sy = y1 + uy * (k * dash_len * 2)
                ex = sx + ux * dash_len
                ey = sy + uy * dash_len
                if (ex - x1) * ux + (ey - y1) * uy > length:
                    ex, ey = x2, y2
                page.draw_line((sx, sy), (ex, ey), color=color, width=width)
        dashed_line((rect.x0, rect.y0), (rect.x1, rect.y0))
        dashed_line((rect.x1, rect.y0), (rect.x1, rect.y1))
        dashed_line((rect.x1, rect.y1), (rect.x0, rect.y1))
        dashed_line((rect.x0, rect.y1), (rect.x0, rect.y0))


def _detect_with_qwen(input_path, requirements=""):
    """
    用 Qwen 检测包装版面。返回:
        {
          "has_top": bool, "has_bottom": bool,
          "boxes": [{"rect": fitz.Rect(PDF坐标), "class_name": str}, ...]
        }
    或 None (调用失败)。

    说明: Qwen 返回的 box_norm 是相对"整张 PDF 页面图像"的归一化坐标 (0~1),
    与 fitz 打开 PDF 的 mediabox 坐标系一致, 因此乘以页面宽高即得 PDF 坐标,
    可直接与 Adobe 刀版线重建出的网格 (同为 PDF 坐标) 做重叠匹配。
    """
    try:
        from qwen_detection import pdf_to_image, get_packaging_detections
    except Exception as e:
        print(f"[Qwen] 模块导入失败: {e}")
        return None

    try:
        print("[Qwen] 正在请求版面检测框 (用于语义分类融合)...")
        img = pdf_to_image(input_path, zoom_factor=2)
        results = get_packaging_detections(img, requirements)
        if not results:
            print("[Qwen] 未返回结果")
            return None

        try:
            pr = fitz.open(input_path)[0].rect
            pdf_w, pdf_h = pr.width, pr.height
        except Exception:
            pdf_w, pdf_h = 1.0, 1.0

        boxes = []
        for r in results:
            nb = r.get("box_norm")
            lb = r.get("class_name")
            if not nb or not lb or len(nb) != 4:
                continue
            x0, y0, x1, y1 = nb
            boxes.append({
                "rect": fitz.Rect(x0 * pdf_w, y0 * pdf_h, x1 * pdf_w, y1 * pdf_h),
                "class_name": lb,
            })

        labels = [b["class_name"] for b in boxes]
        has_top = any("上摇盖" in lb for lb in labels)
        has_bottom = any("下摇盖" in lb for lb in labels)
        print(f"[Qwen] 检测到 {len(boxes)} 个框, has_top={has_top}, has_bottom={has_bottom}")
        return {"has_top": has_top, "has_bottom": has_bottom, "boxes": boxes}
    except Exception as e:
        print(f"[Qwen] 调用失败: {e}")
        return None


def _rule_label(r_idx, col_label):
    """行列规则推导的语义标签 (Qwen 无匹配时的回退)。"""
    if r_idx == 0:
        return "正唛上摇盖" if col_label == "主面" else "侧唛上摇盖"
    if r_idx == 1:
        return "正唛内容" if col_label == "主面" else "侧唛内容"
    return "正唛下摇盖" if col_label == "主面" else "侧唛下摇盖"


def _fuse_labels(grid, rows_to_skip, cols_to_skip, col_labels, qwen_boxes, min_overlap=0.35):
    """语义融合: 每个可见网格 cell 用与之重叠最大的 Qwen 框的标签 (Adobe 定边界, Qwen 定语义);
    无足够重叠时回退到行列规则标签。

    返回 {(r_idx, c_idx): {"class_name": str, "source": "qwen"/"rule", "overlap": float}}。
    """
    label_map = {}
    for r_idx in range(3):
        if r_idx in rows_to_skip:
            continue
        for c_idx in range(4):
            if c_idx in cols_to_skip:
                continue
            cell = grid[r_idx][c_idx]
            cell_area = max(1.0, cell.get_area())
            best_lb, best_ov = None, 0.0
            for gb in qwen_boxes:
                inter = cell & gb["rect"]
                ia = inter.get_area() if not inter.is_empty else 0.0
                if ia <= 0:
                    continue
                # 重叠度: 交集 / 两者较小面积 (对尺寸差异更鲁棒)
                ov = ia / max(1.0, min(cell_area, gb["rect"].get_area()))
                if ov > best_ov:
                    best_ov, best_lb = ov, gb["class_name"]
            rule = _rule_label(r_idx, col_labels[c_idx])
            if best_lb and best_ov >= min_overlap:
                label_map[(r_idx, c_idx)] = {"class_name": best_lb, "source": "qwen", "overlap": best_ov}
            else:
                label_map[(r_idx, c_idx)] = {"class_name": rule, "source": "rule", "overlap": best_ov}
    return label_map


def _build_final_detection_output(grid, rows_to_skip, col_labels, zoom_factor=2.0,
                                  cols_to_skip=None, label_map=None):
    """
    将最终网格转成可直接复用的检测结果:
      - 坐标: zoom_factor 图像坐标 (非 Adobe 坐标)
      - 格式: [{'box': [x0,y0,x1,y1], 'class_name': '...', 'label_source': 'qwen'/'rule'}, ...]
    分类来源:
      - 优先采用 label_map (Qwen 语义融合) 的标签;
      - 无 label_map 时回退到行列规则。
    """
    cols_to_skip = set(cols_to_skip or [])
    output = []
    for r_idx in range(3):
        if r_idx in rows_to_skip:
            continue
        for c_idx in range(4):
            if c_idx in cols_to_skip:
                continue
            col_label = col_labels[c_idx]
            rect = grid[r_idx][c_idx]

            source = "rule"
            if label_map and (r_idx, c_idx) in label_map:
                class_name = label_map[(r_idx, c_idx)]["class_name"]
                source = label_map[(r_idx, c_idx)]["source"]
            else:
                class_name = _rule_label(r_idx, col_label)

            if class_name is None:
                continue

            output.append({
                "box": [
                    rect.x0 * zoom_factor,
                    rect.y0 * zoom_factor,
                    rect.x1 * zoom_factor,
                    rect.y1 * zoom_factor,
                ],
                "class_name": class_name,
                "label_source": source,
            })
    return output


def visualize_dielines(input_path, output_path, use_qwen=True, qwen_requirements=""):
    """
    用 Adobe 识别刀版线, Python 补全 12 格 (摇盖缺失时按 h_flap = w_side/2 合成),
    然后在原 PDF 上叠加彩色检测框与标签, 保存为新 PDF。

    Args:
        use_qwen: 是否用 Qwen 先确定行结构 (是否有上/下摇盖), 避免合成不存在的行

    Returns:
        dict:
            {
              "success": bool,
              "final_boxes": list[{"box":[x0,y0,x1,y1], "class_name": str}],  # zoomfactor=2 图像坐标
              "dieline_size_mm": {
                  "w_main_mm": float, "w_side_mm": float, "h_panel_mm": float, "h_flap_mm": float
              },
              ...
            }
    """
    print("=" * 60)
    print(f"输入: {input_path}")
    print("=" * 60)

    # 0) 可选: Qwen 检测 -> 行结构先验 + 语义检测框 (用于分类融合)
    gem = _detect_with_qwen(input_path, qwen_requirements) if use_qwen else None
    if gem is not None:
        qwen_has_top, qwen_has_bottom = gem["has_top"], gem["has_bottom"]
        qwen_boxes = gem["boxes"]
    else:
        qwen_has_top, qwen_has_bottom = True, True  # 默认全 3 行
        qwen_boxes = []

    # 1) AI 检测构建 12 格; 同时用 Qwen 框建一套, 质量差时改用 Qwen 几何
    #    (线段重建常把外框/装饰当侧唛, 把真正 4 面吞成超大一格)
    rows, (page_w, page_h) = detect_dielines_with_ai(input_path)
    adobe_pack = build_12_grid(rows)

    qwen_pack = (None, None, None)
    if qwen_boxes:
        gem_rows = _rows_from_qwen_boxes(qwen_boxes)
        if gem_rows:
            qwen_pack = build_12_grid(gem_rows)

    grid, synth, info, source = _prefer_better_grid(
        adobe_pack, qwen_pack, qwen_boxes, page_w, page_h
    )
    if source.startswith("qwen"):
        print(f"[Grid] 采用 Qwen 检测框几何 ({source})")
    elif source == "adobe":
        print("[Grid] 采用 Adobe 刀版线几何")

    if grid is None:
        print("错误: 未识别到 4 格完整的主面行, 无法推断网格。")
        return {
            "success": False,
            "error": "未识别到4格完整的主面行, 无法推断网格",
            "final_boxes": [],
            "dieline_size_mm": {},
        }

    # 2) 根据 Qwen 先验裁掉不存在的摇盖行
    rows_to_skip = set()
    if not qwen_has_top:
        rows_to_skip.add(0)
        print("[Qwen] 上摇盖不存在, 跳过绘制第 0 行")
    if not qwen_has_bottom:
        rows_to_skip.add(2)
        print("[Qwen] 下摇盖不存在, 跳过绘制第 2 行")

    CM = 72 / 2.54
    print("\n网格信息 (由主面行推导):")
    print(f"  主面宽 w_main = {info['w_main_pt']:.1f} pt ({info['w_main_pt']/CM:.2f} cm)")
    print(f"  侧唛宽 w_side = {info['w_side_pt']:.1f} pt ({info['w_side_pt']/CM:.2f} cm)")
    print(f"  主面高 h_panel = {info['h_panel_pt']:.1f} pt ({info['h_panel_pt']/CM:.2f} cm)")
    print(f"  摇盖高 h_flap  = {info['h_flap_pt']:.1f} pt ({info['h_flap_pt']/CM:.2f} cm)  [来源: {info.get('h_flap_source', '?')}]")

    col_labels = info.get("col_pattern", COL_LABELS)
    print(f"  列布局: {col_labels}")

    cols_to_skip = set(info.get("cols_to_skip", []))
    if cols_to_skip:
        print(f"  跳过虚拟列 (仅 2 个真实刀版, 不向右补): {sorted(cols_to_skip)}")

    # 语义融合: Adobe 网格 (精确边界) + Qwen 框 (语义分类)
    label_map = _fuse_labels(grid, rows_to_skip, cols_to_skip, col_labels, qwen_boxes)
    n_gem = sum(1 for v in label_map.values() if v["source"] == "qwen")
    print(f"  语义融合: {n_gem}/{len(label_map)} 个格采用 Qwen 分类, 其余回退行列规则")

    n_cols_drawn = 4 - len(cols_to_skip)
    n_rows_drawn = 3 - len(rows_to_skip)
    print(f"\n最终 {n_rows_drawn * n_cols_drawn} 格 (跳过 {len(rows_to_skip)} 行, {len(cols_to_skip)} 列):")
    for r_idx in range(3):
        if r_idx in rows_to_skip:
            print(f"  --- 第 {r_idx} 行 ({ROW_LABELS[r_idx]}) 不存在, 跳过 ---")
            continue
        for c_idx in range(4):
            if c_idx in cols_to_skip:
                continue
            rect = grid[r_idx][c_idx]
            tag = "合成" if synth[r_idx][c_idx] else "检测"
            print(f"  [{r_idx},{c_idx}] {ROW_LABELS[r_idx]}/{col_labels[c_idx]} "
                  f"({tag}) x0={rect.x0:.1f} y0={rect.y0:.1f} "
                  f"{rect.width:.1f}x{rect.height:.1f}")

    final_detections = _build_final_detection_output(
        grid=grid,
        rows_to_skip=rows_to_skip,
        col_labels=col_labels,
        zoom_factor=2.0,
        cols_to_skip=cols_to_skip,
        label_map=label_map,
    )

    # 用 fitz 在原 PDF 上绘制
    doc = fitz.open(input_path)
    page = doc[0]

    # ==========================================
    # 修复：注册中文字体以支持中文标签显示
    # ==========================================
    try:
        # 优先加载 Windows 系统的微软雅黑
        page.insert_font(fontname="msyh", fontfile="C:/Windows/Fonts/msyh.ttc")
        chn_font = "msyh"
    except Exception:
        try:
            # 备选：Windows 黑体
            page.insert_font(fontname="simhei", fontfile="C:/Windows/Fonts/simhei.ttf")
            chn_font = "simhei"
        except Exception:
            # 终极备选：PyMuPDF 自带的 CJK 字体
            page.insert_font(fontname="cjk", fontbuffer=fitz.Font("cjk").buffer)
            chn_font = "cjk"

    font_size = max(10.0, min(page_w, page_h) * 0.012)
    line_width = max(1.0, min(page_w, page_h) * 0.002)

    drawn = 0
    for r_idx in range(3):
        if r_idx in rows_to_skip:
            continue
        for c_idx in range(4):
            if c_idx in cols_to_skip:
                continue
            rect = grid[r_idx][c_idx]
            row_label = ROW_LABELS[r_idx]
            col_label = col_labels[c_idx]
            is_synth = synth[r_idx][c_idx]
            color = _cell_color(row_label, col_label)

            if is_synth:
                # 合成: 虚线 + 稍粗
                _draw_dashed_rect(page, rect, color, line_width * 1.5,
                                  dash_len=max(4.0, line_width * 6))
            else:
                # 检测: 实线
                page.draw_rect(rect, color=color, width=line_width)
            drawn += 1

            tag = "合成" if is_synth else "检测"
            fused = label_map.get((r_idx, c_idx), {})
            cls = fused.get("class_name", _rule_label(r_idx, col_label))
            src = "G" if fused.get("source") == "qwen" else "R"
            label = f"[{r_idx},{c_idx}] {cls} ({tag}/{src})"
            size_txt = f"{rect.width:.1f} x {rect.height:.1f} pt"

            tx = rect.x0 + font_size * 0.4
            ty = rect.y0 + font_size * 1.2
            
            # 使用动态获取的中文字体变量 chn_font 替代原先硬编码的 "helv"
            page.insert_text((tx, ty), label, fontsize=font_size,
                             color=color, fontname=chn_font)
            page.insert_text((tx, ty + font_size * 1.2), size_txt,
                             fontsize=font_size * 0.85, color=color, fontname=chn_font)

    doc.save(output_path)
    doc.close()
    print(f"\n可视化完成! 输出文件: {output_path}  (共绘制 {drawn} 格)")
    print("图例: 实线 = AI 检测到的刀版线, 虚线 = Python 按规则合成")
    pt_to_mm = 25.4 / 72.0
    result = {
        "success": True,
        "input_path": input_path,
        "output_path": output_path,
        "final_boxes": final_detections,
        "dieline_size_mm": {
            "w_main_mm": info["w_main_pt"] * pt_to_mm,
            "w_side_mm": info["w_side_pt"] * pt_to_mm,
            "h_panel_mm": info["h_panel_pt"] * pt_to_mm,
            "h_flap_mm": info["h_flap_pt"] * pt_to_mm,
        },
        "rows_skipped": sorted(list(rows_to_skip)),
    }

    return result

# ==========================================
# 批量处理
# ==========================================
def batch_visualize(input_dir, output_dir=None, suffix="_vis"):
    """对 input_dir 下所有 PDF 做刀版线可视化, 结果保存到 output_dir。"""
    input_dir = os.path.abspath(input_dir)
    if output_dir is None:
        output_dir = input_dir + suffix
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    # 过滤: 跳过已有的可视化输出
    pdfs = sorted([f for f in os.listdir(input_dir)
                   if f.lower().endswith(".pdf") and not f.lower().endswith(f"{suffix}.pdf")])
    print(f"\n########## 批量处理 ##########")
    print(f"输入目录: {input_dir}")
    print(f"输出目录: {output_dir}")
    print(f"共 {len(pdfs)} 个 PDF\n")

    ok, fail = [], []
    for i, name in enumerate(pdfs, 1):
        in_path = os.path.join(input_dir, name)
        out_path = os.path.join(output_dir, name.replace(".pdf", f"{suffix}.pdf"))
        print(f"\n########## [{i}/{len(pdfs)}] {name} ##########")
        try:
            result = visualize_dielines(in_path, out_path)
            if isinstance(result, dict) and result.get("success"):
                ok.append(name)
            else:
                fail.append((name, "未能推断 12 格网格"))
        except Exception as e:
            import traceback
            print(f"!! 异常: {e}")
            traceback.print_exc()
            fail.append((name, str(e)))

    print("\n" + "=" * 60)
    print(f"批量完成: 成功 {len(ok)} 个, 失败 {len(fail)} 个")
    if ok:
        print("成功列表:")
        for n in ok:
            print(f"  + {n}")
    if fail:
        print("失败列表:")
        for name, err in fail:
            print(f"  - {name}: {err}")


# ==========================================
# 调用示例
# ==========================================
if __name__ == "__main__":
    result = visualize_dielines(
        input_path=r"C:\Users\Administrator\Desktop\detection_pdf_purecode\layout_result\AyxPytzjGtXNVoMxAwOE_clean.pdf",
        output_path=r"C:\Users\Administrator\Desktop\detection_pdf_purecode\layout_result\AyxPytzjGtXNVoMxAwOE_clean_vis.pdf",
    )
    print(result)
