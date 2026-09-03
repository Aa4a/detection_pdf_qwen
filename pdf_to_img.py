import os

import fitz

def pdf_to_picture(pdf_path, zoom_factor=2, output_dir=None):
    """
    参考 pipeline 的逐页渲染逻辑：
    - 多页 PDF 全部渲染，不只第 1 页
    - 输出命名为 <pdf名>_page000.png / <pdf名>_page001.png ...
    """
    if output_dir is None:
        output_dir = os.path.dirname(pdf_path)
    os.makedirs(output_dir, exist_ok=True)

    page_paths = []
    doc = fitz.open(pdf_path)
    try:
        stem = os.path.splitext(os.path.basename(pdf_path))[0]
        for page_index in range(len(doc)):
            page = doc[page_index]
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom_factor, zoom_factor))
            out_path = os.path.join(output_dir, f"{stem}_page{page_index:03d}.png")
            if os.path.exists(out_path):
                try:
                    os.remove(out_path)
                except Exception:
                    pass
            pix.save(out_path)
            page_paths.append(out_path)
    finally:
        doc.close()
    return page_paths

if __name__ == "__main__":
    # base_dir = r"C:\Users\18858\Desktop\01192\curve_match_slices\testzxc"
    # for name in os.listdir(base_dir):
    #     if not name.lower().endswith(".pdf"):
    #         continue
    #     pdf_path = os.path.join(base_dir, name)
    #     page_images = pdf_to_picture(pdf_path, zoom_factor=2)
    #     print(f"{pdf_path} -> {len(page_images)} 页")
    #     for img in page_images:
    #         print(img)

    pdf_path = r"C:\Users\18858\Desktop\test_vis\测试PDF\IN8800_doc1.pdf"
    page_images = pdf_to_picture(pdf_path, zoom_factor=2)