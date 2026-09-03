import os
import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFont

# ==========================================
# 可视化 detections 的 box 在原 PDF 上的位置
# box 坐标为 zoom=2 的图像坐标
# ==========================================

ZOOM_FACTOR = 2.0

input_pdf = os.path.abspath(r"C:\Users\Administrator\Desktop\detection_pdf_purecode\output\5-662Z 923759内外箱唛头 -_20260812_111007.pdf")

detections = [
    {"class_name": "正唛内容-1", "box": [39, 495, 729, 1059]},
    {"class_name": "正唛内容-2", "box": [804, 496, 1491, 1051]},
]

# 每个类别一个颜色 (RGB)
COLORS = [
    (255, 0, 0),      # 红
    (0, 128, 255),    # 蓝
    (0, 200, 0),      # 绿
    (255, 128, 0),    # 橙
    (200, 0, 200),    # 紫
    (0, 200, 200),    # 青
]


def get_font(size):
    """尝试加载支持中文的字体, 失败则退回默认字体"""
    candidates = [
        r"C:\Windows\Fonts\msyh.ttc",     # 微软雅黑
        r"C:\Windows\Fonts\simhei.ttf",   # 黑体
        r"C:\Windows\Fonts\simsun.ttc",   # 宋体
    ]
    for fp in candidates:
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                continue
    return ImageFont.load_default()


def main():
    if not os.path.exists(input_pdf):
        print(f"❌ 找不到 PDF: {input_pdf}")
        return

    doc = fitz.open(input_pdf)
    page = doc[0]

    # 按 zoom=2 渲染, 使像素坐标与 box 坐标一致
    pix = page.get_pixmap(matrix=fitz.Matrix(ZOOM_FACTOR, ZOOM_FACTOR))
    mode = "RGBA" if pix.alpha else "RGB"
    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples).convert("RGB")

    draw = ImageDraw.Draw(img)
    font = get_font(28)

    print(f"页面渲染尺寸 (px): {pix.width} x {pix.height}")

    # 相同 box 的标签纵向错开，避免互相遮挡
    label_stack = {}

    for i, det in enumerate(detections):
        x0, y0, x1, y1 = det['box']
        color = COLORS[i % len(COLORS)]
        label = det['class_name']

        # 画框（同坐标会叠画；用不同颜色可区分）
        draw.rectangle([x0, y0, x1, y1], outline=color, width=4)

        # 画标签背景 + 文字
        text = f"{i}:{label}"
        tb = draw.textbbox((0, 0), text, font=font)
        tw, th = tb[2] - tb[0], tb[3] - tb[1]
        key = (int(x0), int(y0), int(x1), int(y1))
        stack = label_stack.get(key, 0)
        label_stack[key] = stack + 1
        ty = max(0, y0 - (th + 8) * (stack + 1) - 2)
        draw.rectangle([x0, ty, x0 + tw + 8, ty + th + 6], fill=color)
        draw.text((x0 + 4, ty + 2), text, fill=(255, 255, 255), font=font)

        print(f"[{i}] {label}: box={det['box']}  (w={x1-x0:.1f}, h={y1-y0:.1f})")

    out_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "detections_visualized.png"
    )
    img.save(out_path)
    doc.close()
    print(f"\n[OK] saved: {out_path}")


if __name__ == "__main__":
    main()
