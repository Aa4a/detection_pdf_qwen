import pikepdf
import sys

def remove_illustrator_private_data(input_pdf, output_pdf):
    try:
        # 打开 PDF 文件
        print(f"正在读取: {input_pdf}")
        pdf = pikepdf.Pdf.open(input_pdf)
        
        # 核心逻辑：Illustrator 通常将私有数据存储在根目录的 /PieceInfo 中
        # 以及页面级别的 /PieceInfo 中
        
        # 1. 移除全局层面的私有数据
        if '/PieceInfo' in pdf.Root:
            del pdf.Root['/PieceInfo']
            print("已清理全局 Illustrator 私有数据。")
            
        # 2. 遍历每一页，移除页面级别的私有数据
        for i, page in enumerate(pdf.pages):
            if '/PieceInfo' in page:
                del page['/PieceInfo']
                print(f"已清理第 {i+1} 页的 Illustrator 私有数据。")

        # 3. 尝试清理可能触发警告的 XMP 元数据（可选，通常删除 PieceInfo 就足够了）
        with pdf.open_metadata() as meta:
            # 这里可以清除一些特定的 Adobe 标记，但为保持最大兼容性，我们先不动主体元数据
            pass

        # 保存为新的 PDF
        pdf.save(output_pdf)
        print(f"处理完成！干净的 PDF 已保存至: {output_pdf}")
        print("现在使用 Illustrator 打开此文件，将直接读取最新 PDF 视图，不会再弹窗。")

    except Exception as e:
        print(f"处理失败: {e}")

if __name__ == "__main__":
    # 使用示例
    input_file = r"C:\Users\18858\Desktop\detection_pdf_purecode\test_rize_500\CDGoynvNwsJ.pdf"   # 替换成你的报错 PDF 路径
    output_file = r"C:\Users\18858\Desktop\detection_pdf_purecode\test_rize_500\CDGoynvNwsJ_clear.pdf"  # 替换成你想保存的路径
    
    remove_illustrator_private_data(input_file, output_file)