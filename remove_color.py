import os
import cv2
import numpy as np
import win32com.client

def check_image_black_ratio(image_path, BLACK_THRESHOLD=0.2):
    if not os.path.exists(image_path):
        print("[X] 找不到预览图片。")
        return False
    try:
        img = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), -1)
        if img is None: return False
        if len(img.shape) == 3:
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        else:
            gray = img
        ret, binary = cv2.threshold(gray, 254, 255, cv2.THRESH_BINARY)
        total_pixels = binary.size
        black_pixels = total_pixels - cv2.countNonZero(binary)
        ratio = black_pixels / total_pixels
        percentage = ratio * 100
        print(f"[统计] 黑色占比(非白区域): {percentage:.2f}% (阈值: {BLACK_THRESHOLD * 100:.2f}%)")
        return ratio >= BLACK_THRESHOLD
    except Exception as e:
        print(f"[X] 检测出错: {e}")
        return False

def remove_background_in_ai_debug(input_file, output_file, debug_log_file, cells=None):
    """
    去除"铺满刀版格的底色"。

    cells: 可选, 刀版格列表 [[x0,y0,x1,y1], ...] (PDF 点坐标, 原点在画板左上, y 向下)。
           传入后只删除"填满某个刀版格"的填充 (底色), 横跨多格的黑色色带等会被保留。
           不传时回退为"铺满整个刀版区域"判断。
    """
    input_path = os.path.abspath(input_file).replace("\\", "/")
    output_path = os.path.abspath(output_file).replace("\\", "/")
    log_path = os.path.abspath(debug_log_file).replace("\\", "/")

    # 将刀版格转成 JS 数组字面量 (PDF 点坐标)
    if cells:
        cells_js = "[" + ",".join(
            "[{:.4f},{:.4f},{:.4f},{:.4f}]".format(c[0], c[1], c[2], c[3]) for c in cells
        ) + "]"
    else:
        cells_js = "[]"

    jsx_code = f"""
    var inputPath = "{input_path}";
    var outputPath = "{output_path}";
    var logPath = "{log_path}";
    var pdfCells = {cells_js}; // [[x0,y0,x1,y1], ...] PDF点 (画板左上为原点, y 向下)

    var logFile = new File(logPath);
    if(logFile.exists) logFile.remove();
    
    function writeLog(msg) {{
        var f = new File(logPath);
        f.encoding = "UTF-8";
        f.open("a");
        f.writeln(msg);
        f.close();
    }}

    app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

    try {{
        if (new File(inputPath).exists) {{
            writeLog("=== 开始处理 PDF ===");
            var doc = app.open(new File(inputPath));
            var items = doc.pathItems;
            var ab = doc.artboards[0].artboardRect; // [left, top, right, bottom]
            var pageW = doc.width;
            var pageH = doc.height;
            var pageArea = pageW * pageH;
            var minSize = Math.min(pageW, pageH) * 0.05;

            writeLog("画板尺寸: " + pageW + " x " + pageH);
            writeLog("PathItems 总数量: " + items.length);

            // ===== 1. 计算刀版格 (优先用传入的检测结果) =====
            // 把传入的 PDF 点坐标刀版格转换到 Adobe 坐标 (y 向上);
            // 同时求并集得到整个刀版区域, 作为没有传入刀版格时的回退判断。
            writeLog("\\n--- 1. 计算刀版格 / 刀版区域 ---");
            var abLeft = ab[0], abTop = ab[1];
            var cells = [];
            for (var ci = 0; ci < pdfCells.length; ci++) {{
                var pc = pdfCells[ci]; // [x0,y0,x1,y1] PDF点, y 向下
                // Adobe: left=x0+abLeft, top=abTop-y0, right=x1+abLeft, bottom=abTop-y1
                cells.push([pc[0] + abLeft, abTop - pc[1], pc[2] + abLeft, abTop - pc[3]]);
            }}
            writeLog("传入刀版格数量: " + cells.length);

            // 刀版区域包围盒 (来自检测的刀版线尺寸过滤, 用于回退)
            var dieL = null, dieT = null, dieR = null, dieB = null, dieCount = 0;
            for (var i = 0; i < items.length; i++) {{
                try {{
                    var it = items[i];
                    if (it.hidden || it.guides) continue;
                    var gb = it.geometricBounds;
                    var l = gb[0], t = gb[1], r = gb[2], b = gb[3];
                    var w = r - l, h = t - b;
                    if (w <= 0 || h <= 0) continue;
                    if (Math.max(w, h) < minSize) continue;
                    if (w < 2 && h < 150) continue;
                    if (h < 2 && w < 150) continue;
                    if ((w * h) / pageArea >= 0.80) continue;
                    if (w >= pageW * 0.95 && h >= pageH * 0.95) continue;
                    if (dieL === null) {{ dieL = l; dieT = t; dieR = r; dieB = b; }}
                    else {{
                        if (l < dieL) dieL = l;
                        if (t > dieT) dieT = t;
                        if (r > dieR) dieR = r;
                        if (b < dieB) dieB = b;
                    }}
                    dieCount++;
                }} catch(e) {{ continue; }}
            }}
            if (dieL === null) {{ dieL = ab[0]; dieT = ab[1]; dieR = ab[2]; dieB = ab[3]; }}
            var dieW = dieR - dieL, dieH = dieT - dieB;
            writeLog("刀版区域: L=" + dieL.toFixed(1) + " T=" + dieT.toFixed(1) + " R=" + dieR.toFixed(1) + " B=" + dieB.toFixed(1) + " (" + dieW.toFixed(1) + " x " + dieH.toFixed(1) + ")");

            // ===== 2. 删除"铺满某个刀版格"的底色 =====
            // 规则: 一个填充若覆盖某个刀版格的宽和高各 >= CELL_COVER, 视为该格底色 -> 删除。
            // 黑色色带横跨多格、纵向只占一条, 无法填满任一格, 因此被保留。
            writeLog("\\n--- 2. 删除铺满刀版格的底色 ---");
            var CELL_COVER = 0.90; // 单格覆盖阈值
            var REGION_COVER = 0.85; // 回退: 整刀版区域覆盖阈值
            var deletedCount = 0;
            for (var j = items.length - 1; j >= 0; j--) {{
                try {{
                    var curr = items[j];
                    if (!curr.filled || curr.locked || curr.hidden) continue;

                    var cb = curr.geometricBounds; // [left, top, right, bottom]
                    var cl = cb[0], ct = cb[1], cr = cb[2], cbm = cb[3];

                    var matched = false;
                    var mInfo = "";

                    if (cells.length > 0) {{
                        for (var ci2 = 0; ci2 < cells.length; ci2++) {{
                            var ce = cells[ci2];
                            var ceL = ce[0], ceT = ce[1], ceR = ce[2], ceB = ce[3];
                            var ceW = ceR - ceL, ceH = ceT - ceB;
                            if (ceW <= 0 || ceH <= 0) continue;
                            var ow = Math.min(cr, ceR) - Math.max(cl, ceL);
                            var oh = Math.min(ct, ceT) - Math.max(cbm, ceB);
                            if (ow <= 0 || oh <= 0) continue;
                            if ((ow / ceW) >= CELL_COVER && (oh / ceH) >= CELL_COVER) {{
                                matched = true;
                                mInfo = "格#" + ci2 + " cW=" + (ow / ceW).toFixed(2) + " cH=" + (oh / ceH).toFixed(2);
                                break;
                            }}
                        }}
                    }} else {{
                        // 回退: 铺满整个刀版区域
                        var ovW = Math.min(cr, dieR) - Math.max(cl, dieL);
                        var ovH = Math.min(ct, dieT) - Math.max(cbm, dieB);
                        if (ovW > 0 && ovH > 0 && (ovW / dieW) >= REGION_COVER && (ovH / dieH) >= REGION_COVER) {{
                            matched = true;
                            mInfo = "刀版区域";
                        }}
                    }}

                    if (matched) {{
                        writeLog("删除底色 idx[" + j + "] 命中 " + mInfo);
                        curr.remove();
                        deletedCount++;
                    }}
                }} catch(e) {{
                    writeLog("索引 [" + j + "] 处理出错: " + e.message);
                }}
            }}

            writeLog("总计删除底色数量: " + deletedCount);
            if (deletedCount > 0) {{
                var pdfOptions = new PDFSaveOptions();
                pdfOptions.compatibility = PDFCompatibility.ACROBAT7;
                pdfOptions.preserveEditability = true;
                doc.saveAs(new File(outputPath), pdfOptions);
                writeLog("已保存新 PDF: " + outputPath);
            }} else {{
                writeLog("未找到铺满刀版区域的底色, 不生成新文件。");
            }}
            
            doc.close(SaveOptions.DONOTSAVECHANGES);
            writeLog("=== 处理结束 ===");
        }} else {{
            writeLog("错误: 找不到输入文件 " + inputPath);
        }}
    }} catch(e) {{
        writeLog("全局严重错误: " + e.message);
    }}
    """

    print("[启动] 启动 Illustrator (Debug Mode)...")
    try:
        ai = win32com.client.Dispatch("Illustrator.Application")
        ai.DoJavaScript(jsx_code)
        print(f"[OK] 处理完成，请查看日志文件: {log_path}")
    except Exception as e:
        print(f"[X] AI 处理发生错误: {e}")

def re_color_pdf(INPUT_PDF, preview_img_path, OUTPUT_FOLDER, cells=None):
    out_name = os.path.basename(INPUT_PDF).replace(".pdf", "_nocolor.pdf")
    OUTPUT_PDF = os.path.join(OUTPUT_FOLDER, out_name)
    DEBUG_LOG = os.path.join(OUTPUT_FOLDER, "ai_debug_log.txt")

    if check_image_black_ratio(preview_img_path):
        remove_background_in_ai_debug(INPUT_PDF, OUTPUT_PDF, DEBUG_LOG, cells=cells)
        if os.path.exists(OUTPUT_PDF):
            return OUTPUT_PDF
    return INPUT_PDF

if __name__ == "__main__":
    import fitz

    INPUT_PDF = r"C:\Users\Administrator\Desktop\detection_pdf_purecode\layout_result\AFFohtMumE_clean.pdf"
    OUTPUT_FOLDER = r"C:\Users\18858\Desktop\detection_pdf_purecode"

    # 检测结果 (final_boxes, zoom_factor=2 图像坐标)
    detections = [{'box': [205.843994140624, 226.83422851562398, 521.941589355468, 265.593994140624], 'class_name': '侧唛上摇盖'}, {'box': [521.941589355468, 226.83422851562398, 972.479614257812, 266.56219482421795], 'class_name': '正唛上摇盖'}, {'box': [972.699584960938, 226.83422851562398, 1290.169799804688, 264.218017578124], 'class_name': '侧唛上摇盖'}, {'box': [1290.169799804688, 226.83422851562398, 1742.087646484376, 264.218017578124], 'class_name': '正唛上摇盖'}, {'box': [205.18179321289, 264.62219238281193, 521.941589355468, 581.377990722656], 'class_name': '侧唛内容'}, {'box': [521.941589355468, 264.62219238281193, 974.24560546875, 581.377990722656], 'class_name': '正唛内容'}, {'box': [973.273803710938, 264.218017578124, 1290.169799804688, 581.377990722656], 'class_name': '侧唛内容'}, {'box': [1290.169799804688, 264.218017578124, 1742.087646484376, 581.377990722656], 'class_name': '正唛内容'}, {'box': [205.518005371094, 582.554412841796, 522.277770996094, 621.31199645996], 'class_name': '侧唛下摇盖'}, {'box': [521.943786621094, 582.5503845214839, 973.364013671876, 621.31199645996], 'class_name': '正唛下摇盖'}, {'box': [974.24560546875, 582.554412841796, 1291.005981445312, 621.31199645996], 'class_name': '侧唛下摇盖'}, {'box': [1290.1181640625, 582.554412841796, 1742.087646484376, 621.31199645996], 'class_name': '正唛下摇盖'}]

    # final_boxes 为 zoom_factor=2 图像坐标, 还原为 PDF 点坐标 (除以 2)
    cells_pt = [[v / 2.0 for v in d["box"]] for d in detections if "box" in d]

    # 生成预览图 (与 pipeline.py 一致, zoom_factor=2)
    preview_dir = os.path.join(OUTPUT_FOLDER, "result", "_preview")
    os.makedirs(preview_dir, exist_ok=True)
    preview_base = os.path.splitext(os.path.basename(INPUT_PDF))[0]
    preview_raw_path = os.path.join(preview_dir, f"{preview_base}_preview.png")
    _doc = fitz.open(INPUT_PDF)
    _doc[0].get_pixmap(matrix=fitz.Matrix(2, 2)).save(preview_raw_path)
    _doc.close()

    ori_pdf_path = re_color_pdf(INPUT_PDF, preview_raw_path, OUTPUT_FOLDER, cells=cells_pt)
    print("   -> 最终路径: ", ori_pdf_path)