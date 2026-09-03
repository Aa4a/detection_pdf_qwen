import win32com.client as win32
import os
import time
import subprocess
import pythoncom
import traceback
import fitz


# ==========================================
# 1. 辅助工具函数
# ==========================================

def force_kill_process(process_name):
    """强制关闭进程"""
    try:
        subprocess.run(f"taskkill /f /im {process_name}",
                       shell=True, capture_output=True, text=True)
    except Exception:
        pass


def safe_excel_cleanup(excel_obj=None):
    """安全清理 Excel 对象"""
    try:
        if excel_obj:
            excel_obj.DisplayAlerts = False
            for wb in excel_obj.Workbooks:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            excel_obj.Quit()
            time.sleep(0.5)
    except Exception:
        pass
    finally:
        try:
            del excel_obj
        except Exception:
            pass
        force_kill_process("EXCEL.EXE")


def _merge_pdfs(pdf_paths, output_path):
    """按顺序合并多个单页 PDF 为一个多页 PDF。"""
    if len(pdf_paths) == 1:
        # 单页直接拷贝, 避免无意义重写
        import shutil
        shutil.copy2(pdf_paths[0], output_path)
        return output_path

    out = fitz.open()
    try:
        for p in pdf_paths:
            src = fitz.open(p)
            try:
                out.insert_pdf(src)
            finally:
                src.close()
        out.save(output_path)
    finally:
        out.close()
    return output_path


def _resolve_sheet_print_range(ws):
    """
    计算工作表真实打印范围。
    空表 (无单元格值且无有效 Shape) 返回 None。
    """
    XL_VALUES, XL_BY_ROWS, XL_BY_COLS, XL_PREVIOUS, XL_PART = -4163, 1, 2, 2, 2
    used = ws.UsedRange
    rng = used

    try:
        last_cell_row = ws.Cells.Find(
            "*", ws.Cells(1, 1), XL_VALUES, XL_PART, XL_BY_ROWS, XL_PREVIOUS
        )
        last_cell_col = ws.Cells.Find(
            "*", ws.Cells(1, 1), XL_VALUES, XL_PART, XL_BY_COLS, XL_PREVIOUS
        )
        if last_cell_row is None or last_cell_col is None:
            raise ValueError("no value cells")
        last_row = last_cell_row.Row
        last_col = last_cell_col.Column
        max_row, max_col = last_row, last_col

        # 1) 有值单元格的 MergeArea
        for r in range(1, last_row + 1):
            for c in range(1, last_col + 1):
                cell = ws.Cells(r, c)
                try:
                    v = cell.Value
                except Exception:
                    continue
                if v is None or str(v).strip() == "":
                    continue
                if cell.MergeCells:
                    ma = cell.MergeArea
                    max_row = max(max_row, ma.Row + ma.Rows.Count - 1)
                    max_col = max(max_col, ma.Column + ma.Columns.Count - 1)

        # 2) 内容行带内的其它合并区
        used_last_col = used.Column + used.Columns.Count - 1
        scan_rows = max_row
        for r in range(1, scan_rows + 1):
            c = 1
            while c <= used_last_col:
                cell = ws.Cells(r, c)
                if cell.MergeCells:
                    ma = cell.MergeArea
                    if ma.Row == r and ma.Column == c:
                        max_row = max(max_row, ma.Row + ma.Rows.Count - 1)
                        max_col = max(max_col, ma.Column + ma.Columns.Count - 1)
                    c = ma.Column + ma.Columns.Count
                else:
                    c += 1

        # 3) Shape 锚点
        try:
            for i in range(1, ws.Shapes.Count + 1):
                shp = ws.Shapes(i)
                try:
                    if float(shp.Height) < 2:
                        continue
                    br = shp.BottomRightCell
                    max_row = max(max_row, br.Row)
                    max_col = max(max_col, br.Column)
                except Exception:
                    continue
        except Exception:
            pass

        rng = ws.Range(ws.Cells(1, 1), ws.Cells(max_row, max_col))
        print(f"   [{ws.Name}] 真实内容范围: {rng.Address} "
              f"(Find值止于 R{last_row}C{last_col}, UsedRange: {used.Address})")
        return rng
    except Exception as e:
        # 无单元格值时, 若有有效 Shape 仍导出 (纯图形唛头)
        try:
            shape_max_r, shape_max_c = 0, 0
            for i in range(1, ws.Shapes.Count + 1):
                shp = ws.Shapes(i)
                try:
                    if float(shp.Height) < 2:
                        continue
                    br = shp.BottomRightCell
                    shape_max_r = max(shape_max_r, br.Row)
                    shape_max_c = max(shape_max_c, br.Column)
                except Exception:
                    continue
            if shape_max_r > 0 and shape_max_c > 0:
                rng = ws.Range(ws.Cells(1, 1), ws.Cells(shape_max_r, shape_max_c))
                print(f"   [{ws.Name}] 无单元格值, 按 Shape 范围: {rng.Address}")
                return rng
        except Exception:
            pass
        print(f"   [{ws.Name}] 跳过空表 ({e})")
        return None


def _prepare_sheet_page_setup(ws, rng):
    """隐藏残留 Shape, 配置单页导出 PageSetup。"""
    try:
        hidden_shapes = 0
        for i in range(1, ws.Shapes.Count + 1):
            shp = ws.Shapes(i)
            try:
                if float(shp.Height) < 2:
                    shp.Visible = False
                    hidden_shapes += 1
            except Exception:
                continue
        if hidden_shapes:
            print(f"   [{ws.Name}] 已隐藏残留 Shape: {hidden_shapes} 个")
    except Exception:
        pass

    ws.ResetAllPageBreaks()
    ws.PageSetup.PrintArea = rng.Address
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = 1
    ws.PageSetup.PaperSize = 8  # A3

    if rng.Width > rng.Height:
        ws.PageSetup.Orientation = 2  # 横向
    else:
        ws.PageSetup.Orientation = 1  # 纵向

    ws.PageSetup.LeftHeader = ""
    ws.PageSetup.CenterHeader = ""
    ws.PageSetup.RightHeader = ""
    ws.PageSetup.LeftFooter = ""
    ws.PageSetup.CenterFooter = ""
    ws.PageSetup.RightFooter = ""

    for margin in ['Left', 'Right', 'Top', 'Bottom', 'Header', 'Footer']:
        setattr(ws.PageSetup, f"{margin}Margin", 0)


def _ai_cleanup_pdf(ai, input_pdf, output_pdf):
    """Illustrator: 解组、去白边、按内容裁切画板。"""
    js_pdf_path = input_pdf.replace("\\", "/")
    js_output_pdf = output_pdf.replace("\\", "/")

    jsx_script = f"""
    (function () {{
        var logs = [];
        function addLog(msg) {{ logs.push(msg); }}

        try {{
            var originalInteraction = app.userInteractionLevel;
            app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

            var f = new File("{js_pdf_path}");
            if (!f.exists) throw "中间 PDF 文件未找到";

            var doc = app.open(f);

            // --- A. 暴力解组与释放蒙版 ---
            app.executeMenuCommand('selectall');
            for(var i=0; i<8; i++) {{
                app.executeMenuCommand('releaseMask');
                app.executeMenuCommand('ungroup');
            }}
            addLog("解组完成");

            // --- B. 智能清理 (删除背景和空框) ---
            var items = doc.pageItems;
            var delCount = 0;

            for (var i = items.length - 1; i >= 0; i--) {{
                var item = items[i];
                var removeIt = false;
                try {{
                    if (item.typename === "PathItem") {{
                        if (!item.filled && !item.stroked) {{
                            removeIt = true;
                        }}
                        else if (item.filled) {{
                            var c = item.fillColor;
                            var isWhite = false;

                            if (c.typename === "RGBColor" && c.red > 250 && c.green > 250 && c.blue > 250) isWhite = true;
                            else if (c.typename === "CMYKColor" && c.cyan==0 && c.magenta==0 && c.yellow==0 && c.black==0) isWhite = true;
                            else if (c.typename === "GrayColor" && c.gray == 0) isWhite = true;

                            var isBigWhite = (item.width > 200 && item.height > 200);
                            var isThinWhite = (Math.min(item.width, item.height) < 3);
                            if (isWhite && !item.stroked && (isBigWhite || isThinWhite)) {{
                                removeIt = true;
                            }}
                        }}
                    }}
                }} catch(e) {{}}

                if (removeIt) {{
                    try {{
                        item.remove();
                        delCount++;
                    }} catch(e) {{}}
                }}
            }}
            addLog("清理对象数: " + delCount);

            // --- C. 根据剩余内容计算真实边界 ---
            var finalItems = doc.pageItems;
            if (finalItems.length > 0) {{
                var minX = 99999, maxX = -99999, minY = 99999, maxY = -99999;
                var hasContent = false;

                for (var j = 0; j < finalItems.length; j++) {{
                    var it = finalItems[j];
                    var skipInvisible = false;
                    try {{
                        if (it.typename === "PathItem" && it.filled && !it.stroked) {{
                            var fc = it.fillColor;
                            if ((fc.typename === "RGBColor" && fc.red > 250 && fc.green > 250 && fc.blue > 250) ||
                                (fc.typename === "CMYKColor" && fc.cyan==0 && fc.magenta==0 && fc.yellow==0 && fc.black==0) ||
                                (fc.typename === "GrayColor" && fc.gray == 0)) {{
                                skipInvisible = true;
                            }}
                        }}
                    }} catch(e) {{}}
                    
                    try {{
                        if (!it.guides && !it.hidden && !skipInvisible) {{
                            var b = it.geometricBounds;
                            if (b && b.length === 4) {{
                                if (b[0] < minX) minX = b[0];
                                if (b[1] > maxY) maxY = b[1];
                                if (b[2] > maxX) maxX = b[2];
                                if (b[3] < minY) minY = b[3];
                                hasContent = true;
                            }}
                        }}
                    }} catch(e) {{
                        // 忽略某些引发异常的幽灵对象(0面积对象等)
                    }}
                }}

                if (hasContent) {{
                    try {{
                        var pad = 10;
                        var newRect = [minX - pad, maxY + pad, maxX + pad, minY - pad];
                        doc.artboards[0].artboardRect = newRect;
                        addLog("画板裁切成功");
                    }} catch(e) {{
                        addLog("画板裁切异常:" + e);
                    }}
                }}
            }}

            addLog("准备保存");
            var outF = new File("{js_output_pdf}");
            var opts = new PDFSaveOptions();
            opts.compatibility = PDFCompatibility.ACROBAT7;
            
            // 关键修复：关闭编辑保留，避免触发保存警告导致被自动Cancel
            opts.preserveEditability = false; 

            // 在保存前恢复交互级别，确保后续如有不可抗力弹窗，至少不会被粗暴抛出 "cancelled"
            app.userInteractionLevel = originalInteraction; 
            
            try {{
                doc.saveAs(outF, opts);
                addLog("保存完毕");
            }} catch(err) {{
                throw "SaveAs失败: " + err;
            }}

            doc.close(SaveOptions.DONOTSAVECHANGES);
            return "Success";

        }} catch (e) {{
            try {{ 
                if(doc) doc.close(SaveOptions.DONOTSAVECHANGES); 
                app.userInteractionLevel = originalInteraction;
            }} catch(x) {{}}
            return "Error: " + e + " | Logs: " + logs.join(",");
        }}
    }})();
    """

    result = ai.DoJavaScript(jsx_script)
    if result and "Error" in str(result):
        raise RuntimeError(str(result))
    if not os.path.exists(output_pdf):
        raise RuntimeError(f"AI 未生成输出 PDF | Logs: {result}")
    return output_pdf


# ==========================================
# 2. 核心转换函数
# ==========================================

def excel_to_ai_pdf(input_excel_path, output_folder=None, kill_excel_start=True):
    """
    将 Excel 文件通过 AI 转换为 PDF。
    - 单 sheet → 单页 PDF
    - 多 sheet → 每张可见工作表一页, 合并为多页 PDF
    集成特性：
    1. Excel 端：强制单页、清除边距、清除页眉页脚。
    2. AI 端：暴力解散编组、删除幽灵路径、删除白色大背景、根据真实内容裁切画板。
    """

    if not os.path.exists(input_excel_path):
        print(f"❌ 错误：文件不存在 -> {input_excel_path}")
        return None

    input_excel_path = os.path.abspath(input_excel_path)
    file_dir = os.path.dirname(input_excel_path)
    base_name = os.path.splitext(os.path.basename(input_excel_path))[0]

    if output_folder:
        save_dir = os.path.abspath(output_folder)
    else:
        save_dir = os.path.join(file_dir, "output")
    os.makedirs(save_dir, exist_ok=True)

    timestamp = time.strftime("%Y%m%d_%H%M%S")
    temp_xlsx_path = os.path.join(file_dir, f"temp_fix_{timestamp}.xlsx")
    final_output_pdf = os.path.join(save_dir, f"{base_name}_{timestamp}.pdf")
    temp_files = [temp_xlsx_path]

    print(f"=== 开始处理: {base_name} ===")

    if kill_excel_start:
        force_kill_process("EXCEL.EXE")

    excel = None
    XL_SHEET_VISIBLE = -1
    XL_WORKSHEET = -4167

    try:
        # --- 1. Excel: 清洗并按 sheet 导出 ---
        print("🔨 [1/3] Excel 格式化与导出...")
        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            wb_src = excel.Workbooks.Open(input_excel_path, UpdateLinks=0, ReadOnly=True)
        except Exception:
            wb_src = excel.Workbooks.Open(input_excel_path, CorruptLoad=1, UpdateLinks=0)

        wb_src.SaveAs(temp_xlsx_path, FileFormat=51)
        wb_src.Close(SaveChanges=False)
        safe_excel_cleanup(excel)

        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb_clean = excel.Workbooks.Open(temp_xlsx_path, UpdateLinks=0)

        sheet_pdfs = []  # [(sheet_name, temp_pdf_path), ...]
        for idx in range(1, wb_clean.Sheets.Count + 1):
            ws = wb_clean.Sheets(idx)
            try:
                if ws.Type != XL_WORKSHEET:
                    print(f"   [{ws.Name}] 跳过非工作表 (Type={ws.Type})")
                    continue
                if ws.Visible != XL_SHEET_VISIBLE:
                    print(f"   [{ws.Name}] 跳过隐藏表")
                    continue
            except Exception as e:
                print(f"   [Sheets({idx})] 跳过: {e}")
                continue

            rng = _resolve_sheet_print_range(ws)
            if rng is None:
                continue

            _prepare_sheet_page_setup(ws, rng)

            sheet_pdf = os.path.join(
                file_dir, f"temp_sheet{idx}_{timestamp}.pdf"
            )
            temp_files.append(sheet_pdf)
            ws.ExportAsFixedFormat(0, sheet_pdf, 0, True, False)
            if not os.path.exists(sheet_pdf):
                print(f"   [{ws.Name}] 导出失败, 跳过")
                continue
            sheet_pdfs.append((ws.Name, sheet_pdf))
            print(f"   [{ws.Name}] 已导出中间 PDF")

        wb_clean.Close(SaveChanges=False)
        safe_excel_cleanup(excel)
        excel = None

        if not sheet_pdfs:
            print("❌ 没有可导出的工作表")
            return None

        print(f"   共 {len(sheet_pdfs)} 个 sheet 待 AI 处理")

        # --- 2. Illustrator: 逐页清洗 ---
        print("🚀 [2/3] Illustrator 智能去白边与裁切...")
        try:
            ai = win32.GetActiveObject("Illustrator.Application")
        except Exception:
            ai = win32.Dispatch("Illustrator.Application")
            time.sleep(3)

        cleaned_pdfs = []
        for i, (sheet_name, sheet_pdf) in enumerate(sheet_pdfs, 1):
            cleaned = os.path.join(
                file_dir, f"temp_clean{i}_{timestamp}.pdf"
            )
            temp_files.append(cleaned)
            print(f"   AI 处理 [{i}/{len(sheet_pdfs)}] {sheet_name} ...")
            try:
                _ai_cleanup_pdf(ai, sheet_pdf, cleaned)
                cleaned_pdfs.append(cleaned)
            except Exception as e:
                print(f"   ❌ [{sheet_name}] AI 处理失败: {e}")
                return None

        # --- 3. 合并多页 ---
        print(f"📎 [3/3] 合并 {len(cleaned_pdfs)} 页 PDF...")
        _merge_pdfs(cleaned_pdfs, final_output_pdf)

        print(f"✅ 成功生成: {os.path.basename(final_output_pdf)} "
              f"({len(cleaned_pdfs)} 页)")
        return final_output_pdf

    except Exception as e:
        print(f"\n❌ 程序异常: {e}")
        traceback.print_exc()
        return None

    finally:
        if excel:
            safe_excel_cleanup(excel)
        for p in temp_files:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass


# ==========================================
# 3. 运行入口
# ==========================================
if __name__ == "__main__":
    excel_path = r"Z:\public\excel_1000\excel_1000\AAYocfTEOD.xlsx"
    excel_to_ai_pdf(excel_path)
