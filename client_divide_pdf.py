"""调用 /divide_pdf 的最小示例。"""

import json

import requests


API_URL = "http://127.0.0.1:8003/divide_pdf"
PDF_URL = "http://192.168.0.64:9000/sun/example.pdf"  # 换成可下载的 PDF URL


response = requests.post(
    API_URL,
    json={
        "pdf_url": PDF_URL,
        "zoom_factor": 2.0,
        "grid_width_ratio": 0.15,
        "grid_height_px": 40,
        "trim_whitespace": True,
        "trim_margin_pt": 5.0,
    },
    timeout=1800,
)
response.raise_for_status()
data = response.json()
print(json.dumps(data, ensure_ascii=False, indent=2))

if data["status"] == "success":
    for item in data["result"]["outputs"]:
        print(f"part {item['part']}: {item['output_url']}")
