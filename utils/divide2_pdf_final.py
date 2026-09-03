import cv2
import numpy as np
import fitz  # PyMuPDF
import os
import win32com.client  # 新增：用于调用 Adobe Illustrator


def ensure_clean_dir(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f"📂 [系统] 新建目录: {directory}/")
    else:
        print(f"📂 [系统] 使用现有目录: {directory}/ (未删除旧文件)")


def imread_unicode(path, flags=cv2.IMREAD_COLOR):
    """支持 Windows 中文/非 ASCII 路径的图像读取。"""
    try:
        data = np.fromfile(path, dtype=np.uint8)
        if data.size == 0:
            return None
        return cv2.imdecode(data, flags)
    except Exception:
        return None


def imwrite_unicode(path, img):
    """支持 Windows 中文/非 ASCII 路径的图像写入。"""
    ext = os.path.splitext(path)[1]
    if not ext:
        ext = ".png"
    ok, buf = cv2.imencode(ext, img)
    if not ok:
        return False
    buf.tofile(path)
    return True


def save_debug_img(debug_dir, filename, img, msg=""):
    path = os.path.join(debug_dir, filename)
    imwrite_unicode(path, img)
    if msg:
        print(f"   -> 👁️ 生成调试图: {filename} ({msg})")


def make_pdf_screenshot(pdf_path, debug_dir, zoom_factor):
    doc = fitz.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom_factor, zoom_factor))

    img_path = os.path.join(debug_dir, "00_original_screenshot.png")
    pix.save(img_path)
    doc.close()
    return img_path


def has_grid_structure(img_crop, min_width_ratio, min_height_px):
    if img_crop is None or img_crop.size == 0:
        return False
    h, w = img_crop.shape[:2]
    gray = cv2.cvtColor(img_crop, cv2.COLOR_BGR2GRAY)
    edges = cv2.Canny(gray, 50, 150)
    kernel = np.ones((3, 3), np.uint8)
    edges_dilated = cv2.dilate(edges, kernel, iterations=1)
    contours, _ = cv2.findContours(edges_dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    for cnt in contours:
        x_c, y_c, w_c, h_c = cv2.boundingRect(cnt)
        if w_c > (w * min_width_ratio) and h_c > min_height_px:
            return True
    return False


def process_image_and_detect(image_path, debug_dir, kernel_height=3, kernel_width=500):
    base_name = os.path.splitext(os.path.basename(image_path))[0]
    print("🔍 [视觉分析] 开始图像处理...")

    original_img = imread_unicode(image_path)
    if original_img is None:
        raise FileNotFoundError(f"无法读取图像（可能是路径含特殊字符或文件不存在）: {image_path}")
    img_h, img_w = original_img.shape[:2]
    img_area = img_h * img_w

    gray = cv2.cvtColor(original_img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 220, 255, cv2.THRESH_BINARY_INV)
    contours, _ = cv2.findContours(thresh, cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)

    all_rects = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        area = w * h
        if area < img_area * 0.95 and area > img_area * 0.0005:
            all_rects.append(cnt)

    mask = np.zeros((img_h, img_w), dtype=np.uint8)
    cv2.drawContours(mask, all_rects, -1, 255, thickness=-1)
    save_debug_img(debug_dir, f"03_mask_clean_{base_name}.png", mask, "去噪掩膜")

    kernel = np.ones((kernel_height, kernel_width), np.uint8)
    dilated_mask = cv2.dilate(mask, kernel, iterations=1)
    save_debug_img(debug_dir, f"04_mask_dilated_{base_name}.png", dilated_mask, "膨胀后")

    group_contours, _ = cv2.findContours(dilated_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    final_boxes = []
    debug_blocks_img = original_img.copy()

    for i, grp_cnt in enumerate(group_contours):
        x, y, w, h = cv2.boundingRect(grp_cnt)
        final_boxes.append((x, y, w, h))
        cv2.rectangle(debug_blocks_img, (x, y), (x + w, y + h), (0, 255, 0), 2)
        cv2.putText(debug_blocks_img, str(i), (x, y - 5), cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)

    save_debug_img(debug_dir, f"05_detected_blocks_{base_name}.png", debug_blocks_img, "区块识别结果")
    final_boxes.sort(key=lambda b: b[1])

    return final_boxes, (img_h, img_w), original_img


def find_text_top_in_gap(original_img, y_start, y_end, white_threshold=220, min_row_pixels=8):
    """在两区块之间的水平条带 [y_start, y_end) 内做文本边缘检测。

    主区块检测会按面积过滤掉很小的对象（如“外箱 39.5*28*40cm”这类标题文字），
    导致中点切割线穿过标题文字造成截断。这里直接在间隙内逐行统计非白像素，
    返回最靠上的有内容行 y（全图坐标）。无内容返回 None。
    """
    y_start = max(int(y_start), 0)
    y_end = min(int(y_end), original_img.shape[0])
    if y_end - y_start < 2:
        return None

    strip = original_img[y_start:y_end]
    gray = cv2.cvtColor(strip, cv2.COLOR_BGR2GRAY)
    mask = gray < white_threshold
    # 每行非白像素数量，过滤抗锯齿/噪点
    row_counts = mask.sum(axis=1)
    rows_with_content = np.where(row_counts >= min_row_pixels)[0]
    if rows_with_content.size == 0:
        return None
    return y_start + int(rows_with_content[0])


def calculate_cuts_and_debug(boxes, img_dims, original_img, debug_dir, base_name, text_margin_px=8):
    img_h, img_w = img_dims
    cut_lines = [0]

    for i in range(len(boxes) - 1):
        curr_box_bottom = boxes[i][1] + boxes[i][3]
        next_box_top = boxes[i + 1][1]
        mid_point = int((curr_box_bottom + next_box_top) / 2)

        # 在两区块间隙内检测文本（如标题），若有则把切割线移到文本上方，避免截断文字
        text_top = find_text_top_in_gap(original_img, curr_box_bottom, next_box_top)
        if text_top is not None:
            candidate = max(curr_box_bottom + 1, text_top - text_margin_px)
            cut = min(candidate, mid_point)
        else:
            cut = mid_point

        cut_lines.append(cut)

    cut_lines.append(img_h)

    debug_cut_img = original_img.copy()
    for y_line in cut_lines:
        cv2.line(debug_cut_img, (0, y_line), (img_w, y_line), (255, 0, 0), 3)

    save_debug_img(debug_dir, f"06_cut_lines_{base_name}.png", debug_cut_img, "切割线预览 (蓝线)")
    return cut_lines


def split_pdf_via_illustrator(pdf_path, cut_lines, img_h, original_img_ref, debug_dir, grid_min_width_ratio,
                              grid_min_height_px, trim_whitespace=True, trim_margin_pt=5.0):
    """
    [更新核心] 使用 JSX 驱动 Adobe Illustrator
    新增：自动屏蔽“缺失字体”、“缺失链接”等弹窗，实现全自动静默处理
    """
    print("✂️ [执行切割] 正在唤起 Adobe Illustrator 进行无蒙版智能裁切...")

    try:
        app = win32com.client.GetActiveObject("Illustrator.Application")
    except Exception:
        app = win32com.client.Dispatch("Illustrator.Application")

    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    generated_files = []
    rejected_viz = []

    for i in range(len(cut_lines) - 1):
        y1_pixel = cut_lines[i]
        y2_pixel = cut_lines[i + 1]

        # --- 视觉过滤校验 ---
        roi_img = original_img_ref[y1_pixel:y2_pixel, 0:original_img_ref.shape[1]]
        if not has_grid_structure(roi_img, grid_min_width_ratio, grid_min_height_px):
            print(f"   -> 🚫 [跳过] Part {i + 1}: 纯文本或空白 (无网格结构)")
            debug_rej = roi_img.copy()
            cv2.line(debug_rej, (0, 0), (debug_rej.shape[1], debug_rej.shape[0]), (0, 0, 255), 5)
            rejected_viz.append(debug_rej)
            continue

        # --- 使用千分位坐标传递位置 ---
        norm_y1 = int((y1_pixel / img_h) * 1000)
        norm_y2 = int((y2_pixel / img_h) * 1000)

        output_name = os.path.join(debug_dir, f"RESULT_{base_name}_part_{i + 1}.pdf")

        # 强制转换为绝对路径并替换为正斜杠
        pdf_path_js = os.path.abspath(pdf_path).replace("\\", "/")
        output_name_js = os.path.abspath(output_name).replace("\\", "/")

        # --- 构建并执行 JSX 脚本 ---
        # 增加 try...finally 结构，确保脚本哪怕崩溃也能恢复用户的弹窗设置
        jsx_script = f"""
        var originalInteractionLevel = app.userInteractionLevel;
        // 关键：静默模式，屏蔽缺失字体/链接弹窗
        app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS; 

        try {{
            var doc = app.open(new File("{pdf_path_js}"));

            // 解锁全部以确保可以被删除
            app.executeMenuCommand('unlockAll');
            app.executeMenuCommand('showAll');

            var abRect = doc.artboards[0].artboardRect; // [left, top, right, bottom]
            var docW = abRect[2] - abRect[0];
            var docH = abRect[1] - abRect[3];

            // 基于千分位的坐标还原计算
            var top_pt = abRect[1] - ({norm_y1} / 1000) * docH;
            var bottom_pt = abRect[1] - ({norm_y2} / 1000) * docH;

            // 1. 设置新画板范围
            var newAbRect = [abRect[0], top_pt, abRect[2], bottom_pt];
            doc.artboards[0].artboardRect = newAbRect;

            // 2. 遍历并删除完全在画板外部的对象
            for (var i = 0; i < doc.layers.length; i++) {{
                var layer = doc.layers[i];
                if (layer.locked) layer.locked = false;
                if (!layer.visible) layer.visible = true;

                // 倒序遍历对象，防止删除时索引错乱
                for (var j = layer.pageItems.length - 1; j >= 0; j--) {{
                    var item = layer.pageItems[j];
                    try {{
                        if (item.locked) item.locked = false;
                        if (item.hidden) item.hidden = false;

                        var b = item.geometricBounds; 
                        var isOutside = (b[0] > newAbRect[2] || b[2] < newAbRect[0] || b[3] > newAbRect[1] || b[1] < newAbRect[3]);

                        if (isOutside) {{
                            item.remove(); // 物理删除
                        }}
                    }} catch(e) {{
                        // 忽略空组或参考线异常
                    }}
                }}
            }}

            // 3. （可选）将画板收缩到剩余对象的紧致包围盒，去除多余白边
            var doTrim = {str(bool(trim_whitespace)).lower()};
            var trimMargin = {trim_margin_pt};
            var trimDebug = "";
            if (doTrim) {{
                try {{
                    var hasItems = false;
                    var tL = null, tT = null, tR = null, tB = null;
                    var leafCount = 0;

                    // 递归遍历叶子对象（穿透 GroupItem），用 visibleBounds ∩ newAbRect 作为有效内容包围盒
                    function processItem(it) {{
                        var vb = null;
                        try {{ vb = it.visibleBounds; }} catch(eb) {{ return; }}
                        if (!vb) return;
                        if (vb[2] <= vb[0] || vb[1] <= vb[3]) return;

                        // 整体在画板外，直接丢弃
                        if (vb[0] >= newAbRect[2] || vb[2] <= newAbRect[0] ||
                            vb[3] >= newAbRect[1] || vb[1] <= newAbRect[3]) return;

                        // GroupItem 递归下钻，避免把整组当成一个大 bbox
                        if (it.typename === "GroupItem") {{
                            try {{
                                var kids = it.pageItems;
                                if (kids && kids.length > 0) {{
                                    for (var k = 0; k < kids.length; k++) processItem(kids[k]);
                                    return;
                                }}
                            }} catch(eg) {{}}
                        }}

                        // 叶子节点：取与画板的交集作为可见 bbox
                        var ix0 = Math.max(vb[0], newAbRect[0]);
                        var iy1 = Math.min(vb[1], newAbRect[1]);
                        var ix1 = Math.min(vb[2], newAbRect[2]);
                        var iy0 = Math.max(vb[3], newAbRect[3]);
                        if (ix1 - ix0 < 0.01 || iy1 - iy0 < 0.01) return;

                        if (tL === null || ix0 < tL) tL = ix0;
                        if (tT === null || iy1 > tT) tT = iy1;
                        if (tR === null || ix1 > tR) tR = ix1;
                        if (tB === null || iy0 < tB) tB = iy0;
                        leafCount++;
                        hasItems = true;
                    }}

                    var items = doc.pageItems;
                    var n = items.length;
                    trimDebug += "pageItems=" + n + "\\n";
                    for (var ii = 0; ii < n; ii++) processItem(items[ii]);
                    trimDebug += "leafCount=" + leafCount + " tL=" + tL + " tT=" + tT + " tR=" + tR + " tB=" + tB + "\\n";
                    if (hasItems) {{
                        var tightRect = [tL - trimMargin, tT + trimMargin, tR + trimMargin, tB - trimMargin];
                        // 限制在原 newAbRect 范围内
                        if (tightRect[0] < newAbRect[0]) tightRect[0] = newAbRect[0];
                        if (tightRect[1] > newAbRect[1]) tightRect[1] = newAbRect[1];
                        if (tightRect[2] > newAbRect[2]) tightRect[2] = newAbRect[2];
                        if (tightRect[3] < newAbRect[3]) tightRect[3] = newAbRect[3];
                        if (tightRect[2] - tightRect[0] > 1 && tightRect[1] - tightRect[3] > 1) {{
                            doc.artboards[0].artboardRect = tightRect;
                            trimDebug += "applied=" + tightRect.join(",") + "\\n";
                        }} else {{
                            trimDebug += "skipped: rect too small\\n";
                        }}
                    }}
                }} catch(et) {{
                    trimDebug += "TRIM ERROR: " + et + "\\n";
                }}
                // 把调试信息写到与输出 PDF 同目录的 .trim.log
                try {{
                    var logFile = new File("{output_name_js}".replace(/\\.pdf$/, ".trim.log"));
                    logFile.encoding = "UTF-8";
                    logFile.open("w");
                    logFile.write(trimDebug);
                    logFile.close();
                }} catch(el) {{}}
            }}

            // 4. 导出 PDF 设置
            var saveOpts = new PDFSaveOptions();
            saveOpts.preserveEditability = false;
            saveOpts.viewAfterSaving = false;

            // 5. 保存并关闭
            var saveFile = new File("{output_name_js}");
            doc.saveAs(saveFile, saveOpts);
            doc.close(SaveOptions.DONOTSAVECHANGES);

        }} catch(err) {{
            // 记录潜在的执行错误
            // alert("Error: " + err); // 调试时可开启
        }} finally {{
            // 恢复 AI 默认的用户交互弹窗设置，千万不能漏掉
            app.userInteractionLevel = originalInteractionLevel;
        }}
        """

        print(f"   -> ✅ [调用 AI] Part {i + 1}: 正在智能清除多余对象并保存 (已屏蔽弹窗)...")
        app.DoJavaScript(jsx_script)
        generated_files.append(output_name)

    if rejected_viz:
        try:
            max_w = max(img.shape[1] for img in rejected_viz)
            resized_rej = [cv2.resize(img, (max_w, int(img.shape[0] * (max_w / img.shape[1])))) for img in rejected_viz]
            rejection_summary = np.vstack(resized_rej)
            save_debug_img(debug_dir, f"07_rejected_parts_{base_name}.png", rejection_summary, "过滤部分预览")
        except Exception as e:
            print("生成过滤预览图失败:", e)

    return generated_files


def trim_pdf_whitespace(pdf_path, margin_pt=5.0, zoom=2.0, white_threshold=245):
    """渲染每页找到非白像素包围盒，重建为 MediaBox 起点为 (0,0) 的新 PDF。

    用 PyMuPDF 的 show_pdf_page(clip=...) 重建页面，避免 /MediaBox 非零原点
    在 Adobe Illustrator 中显示错位的问题（AI 把画板放在 (0,0) 而内容仍在
    原坐标，会导致画板和内容分离）。

    - margin_pt: 内容周围保留的留白（PDF 点）。
    - zoom: 渲染检测时的缩放倍率，越大越精确，越慢。
    - white_threshold: 灰度阈值，低于该值视为有内容（0-255）。
    返回 True 表示有页面被裁剪。
    """
    try:
        src = fitz.open(pdf_path)
    except Exception as e:
        print(f"   -> ⚠️ 无法打开 PDF 用于去白边: {pdf_path} ({e})")
        return False

    new_doc = fitz.open()
    changed = False
    try:
        for page_index in range(src.page_count):
            page = src[page_index]
            # fitz 坐标（top-down，origin 在原 mediabox 的左上）
            page_w = page.rect.width
            page_h = page.rect.height
            if page_w <= 0 or page_h <= 0:
                continue

            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            gray = cv2.cvtColor(img[:, :, :3], cv2.COLOR_RGB2GRAY) if pix.n >= 3 else img[:, :, 0]

            mask = gray < white_threshold
            if not mask.any():
                # 整页空白，原样复制
                new_doc.insert_pdf(src, from_page=page_index, to_page=page_index)
                continue

            ys, xs = np.where(mask)
            x0_px, x1_px = int(xs.min()), int(xs.max()) + 1
            y0_px, y1_px = int(ys.min()), int(ys.max()) + 1

            # 像素 -> fitz 坐标（top-down）
            clip_x0 = max(x0_px / zoom - margin_pt, 0)
            clip_y0 = max(y0_px / zoom - margin_pt, 0)
            clip_x1 = min(x1_px / zoom + margin_pt, page_w)
            clip_y1 = min(y1_px / zoom + margin_pt, page_h)

            if clip_x1 - clip_x0 < 1 or clip_y1 - clip_y0 < 1:
                new_doc.insert_pdf(src, from_page=page_index, to_page=page_index)
                continue

            clip_rect = fitz.Rect(clip_x0, clip_y0, clip_x1, clip_y1)
            new_w = clip_rect.width
            new_h = clip_rect.height
            new_page = new_doc.new_page(width=new_w, height=new_h)
            # 把源页 clip 区域的内容贴到新页（新页 mediabox 自动是 (0,0,new_w,new_h)）
            new_page.show_pdf_page(new_page.rect, src, page_index, clip=clip_rect)
            changed = True
    except Exception as e:
        print(f"   -> ⚠️ trim_pdf_whitespace 异常: {e}")
        new_doc.close()
        src.close()
        return False

    if changed:
        tmp_path = pdf_path + ".trim.tmp"
        try:
            new_doc.save(tmp_path, deflate=True, garbage=3)
        finally:
            new_doc.close()
            src.close()
        os.replace(tmp_path, pdf_path)
    else:
        new_doc.close()
        src.close()
    return changed


def divide_pdf(input_pdf_path, debug_output_dir, zoom_factor, grid_width_ratio, grid_height_px,
               trim_whitespace=True, trim_margin_pt=5.0):
    base_name = os.path.splitext(os.path.basename(input_pdf_path))[0]
    ensure_clean_dir(debug_output_dir)

    if not os.path.exists(input_pdf_path):
        print(f"❌ 错误: 找不到文件 {input_pdf_path}")
        return

    screenshot_path = make_pdf_screenshot(input_pdf_path, debug_output_dir, zoom_factor)
    boxes, dims, orig_img = process_image_and_detect(screenshot_path, debug_output_dir)

    if not boxes:
        print("⚠️ 未检测到内容，停止。")
        return

    cut_lines = calculate_cuts_and_debug(boxes, dims, orig_img, debug_output_dir, base_name)

    # 1) JSX 阶段先按对象 bounds 紧致化画板（对宽度方向常常无效——因为页边框/标题等对象会撑满整页）。
    # 2) Python 阶段再用渲染像素法重写 MediaBox/CropBox，做最终的去白边（权威，AI/Acrobat/WPS 都一致）。
    generated_files = split_pdf_via_illustrator(
        input_pdf_path,
        cut_lines,
        dims[0],
        orig_img,
        debug_output_dir,
        grid_width_ratio,
        grid_height_px,
        trim_whitespace=trim_whitespace,
        trim_margin_pt=trim_margin_pt,
    )

    if trim_whitespace and generated_files:
        print("✂️ [后处理] 渲染像素法去除每个分片的纯白边距 (重写 MediaBox/CropBox)...")
        for f in generated_files:
            if os.path.exists(f):
                try:
                    ok = trim_pdf_whitespace(f, margin_pt=trim_margin_pt)
                    print(f"   -> {'✅ 已裁白' if ok else '➖ 无需裁白'}: {os.path.basename(f)}")
                except Exception as e:
                    print(f"   -> ⚠️ 裁白失败 {os.path.basename(f)}: {e}")

    return generated_files


if __name__ == "__main__":
    base_dir = r"C:\Users\Administrator\Desktop\detection_pdf_purecode\utils"
    file_name = r"C:\Users\Administrator\Desktop\detection_pdf_purecode\utils\zxc.pdf"

    input_pdf = os.path.join(base_dir, file_name)
    folder_name = os.path.splitext(file_name)[0]
    debug_dir = os.path.join(base_dir, os.path.basename(folder_name))

    os.makedirs(os.path.dirname(debug_dir), exist_ok=True)

    CONFIG = {
        "zoom_factor": 2.0,
        "grid_min_width_ratio": 0.15,
        "grid_min_height_px": 40,
        "trim_whitespace": True,
        "trim_margin_pt": 5.0,
    }

    print(f"\n======== 处理文件: {file_name} ========")

    final_files = divide_pdf(
        input_pdf_path=input_pdf,
        debug_output_dir=debug_dir,
        zoom_factor=CONFIG["zoom_factor"],
        grid_width_ratio=CONFIG["grid_min_width_ratio"],
        grid_height_px=CONFIG["grid_min_height_px"],
        trim_whitespace=CONFIG["trim_whitespace"],
        trim_margin_pt=CONFIG["trim_margin_pt"],
    )

    print("   -> 完整文件列表:", final_files)