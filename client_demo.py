# -*- coding: utf-8 -*-
"""
客户端调用脚本 (对接远程 MinIO + 本服务)

流程 (与既有 call_service.py 一致):
    1) 把本地 PDF 上传到 MinIO 输入桶, 生成 1 小时预签名下载 URL (pdf_url)
       (若直接传入 http(s) 预签名 URL, 则跳过上传)
    2) POST {pdf_url} 到服务 /detect 或 /pipeline
    3) 服务处理完把结果上传 MinIO, 返回 output_url (预签名) + console_url
    4) 用 output_url 下载结果

用法:
    # 上传本地 PDF 后调用
    python client_demo.py detect    "C:/path/to/input.pdf"
    python client_demo.py pipeline  "C:/path/to/input.pdf"

    # ai_layout: 给定检测框直接排版
    #   不带 detections.json 时, 会先调 /detect 拿检测框再排版
    python client_demo.py ai_layout "C:/path/to/input.pdf"
    python client_demo.py ai_layout "C:/path/to/input.pdf" "C:/path/to/detections.json"

    # 直接用已有的预签名 URL 调用 (跳过上传)
    python client_demo.py pipeline "http://192.168.0.64:9000/inputfile/nan_part1.pdf?X-Amz-..."

环境变量 (均有默认值, 可覆盖):
    API_BASE            服务地址, 默认 http://192.168.0.48:8003
    MINIO_ENDPOINT      MinIO 地址, 默认 192.168.0.64:9000
    MINIO_ACCESS_KEY    访问密钥, 必须通过环境变量设置
    MINIO_SECRET_KEY    私有密钥, 必须通过环境变量设置
    MINIO_INPUT_BUCKET  输入桶, 默认 inputfile
"""

import json
import os
import sys
from datetime import timedelta

import requests
from minio import Minio

API_BASE = os.environ.get("API_BASE", "http://192.168.0.48:8003")

MINIO_ENDPOINT = os.environ.get("MINIO_ENDPOINT", "192.168.0.64:9000")
MINIO_ACCESS_KEY = os.environ.get("MINIO_ACCESS_KEY", "")
MINIO_SECRET_KEY = os.environ.get("MINIO_SECRET_KEY", "")
MINIO_SECURE = os.environ.get("MINIO_SECURE", "false").lower() in ("1", "true", "yes", "on")
INPUT_BUCKET = os.environ.get("MINIO_INPUT_BUCKET", "inputfile")

# pipeline 排版参数 (按需修改; 不改就用这些值随请求一起发给服务)
LAYOUT_INSTRUCTION = {
    "正唛内容": [1, 3],
    "侧唛内容": [2, 4],
    "正唛上摇盖": [1, 3],
    "侧唛上摇盖": [2, 4],
    "正唛下摇盖": [1, 3],
    "侧唛下摇盖": [2, 4],
}
BOX_SIZE_MM = {"排版尺寸": [79.87, 55.87, 55.87, 20]}
MARGIN_MM = 0.0


def _minio():
    return Minio(MINIO_ENDPOINT, access_key=MINIO_ACCESS_KEY,
                 secret_key=MINIO_SECRET_KEY, secure=MINIO_SECURE)


def upload_and_presign(local_path):
    """上传本地 PDF 到输入桶, 返回 1 小时有效的预签名下载 URL。"""
    client = _minio()
    if not client.bucket_exists(INPUT_BUCKET):
        client.make_bucket(INPUT_BUCKET)
    object_name = os.path.basename(local_path)
    client.fput_object(INPUT_BUCKET, object_name, local_path, content_type="application/pdf")
    url = client.presigned_get_object(INPUT_BUCKET, object_name, expires=timedelta(hours=1))
    print(f"[上传] {INPUT_BUCKET}/{object_name}")
    return url


def download_url(url, save_path):
    r = requests.get(url, timeout=600)
    r.raise_for_status()
    os.makedirs(os.path.dirname(os.path.abspath(save_path)), exist_ok=True)
    with open(save_path, "wb") as f:
        f.write(r.content)
    print(f"[下载] {save_path}")


def resolve_pdf_url(source):
    """source 是 http(s) URL 则直接用; 否则当作本地文件上传后取预签名 URL。"""
    if source.lower().startswith(("http://", "https://")):
        return source
    if not os.path.isfile(source):
        print(f"文件不存在: {source}")
        sys.exit(1)
    return upload_and_presign(source)


def call(endpoint, source, extra_payload=None, out_dir="client_out"):
    pdf_url = resolve_pdf_url(source)
    payload = {"pdf_url": pdf_url}
    if extra_payload:
        payload.update(extra_payload)

    url = f"{API_BASE}/{endpoint}"
    print(f"POST {url}  (pdf_url={pdf_url[:90]}...)")
    r = requests.post(url, json=payload, timeout=3600)
    print(f"HTTP {r.status_code}")
    data = r.json()
    print(json.dumps(data, ensure_ascii=False, indent=2))

    if data.get("output_url"):
        suffix = "vis.pdf" if endpoint == "detect" else "output.pdf"
        download_url(data["output_url"], os.path.join(out_dir, f"{endpoint}_{suffix}"))
    return data


def detect_boxes(pdf_url):
    """调用 /detect 拿到检测框 (result.final_boxes)。"""
    r = requests.post(f"{API_BASE}/detect", json={"pdf_url": pdf_url}, timeout=3600)
    r.raise_for_status()
    data = r.json()
    if data.get("status") != "success":
        print(f"检测失败: {data.get('error')}")
        sys.exit(1)
    return data.get("result", {}).get("final_boxes", [])


def do_ai_layout(source, detections_file=None, out_dir="client_out"):
    """给定检测框直接排版; 未提供 detections 时先调 /detect 获取。"""
    pdf_url = resolve_pdf_url(source)

    if detections_file:
        with open(detections_file, "r", encoding="utf-8") as f:
            detections = json.load(f)
        # 兼容 {"final_boxes": [...]} 或直接 [...]
        if isinstance(detections, dict):
            detections = detections.get("final_boxes", detections.get("detections", []))
    else:
        print("[ai_layout] 未提供 detections, 先调用 /detect 获取检测框...")
        detections = detect_boxes(pdf_url)
    print(f"[ai_layout] 检测框数量: {len(detections)}")

    payload = {
        "pdf_url": pdf_url,
        "detections": detections,
        "layout_instruction": LAYOUT_INSTRUCTION,
        "box_size_mm": BOX_SIZE_MM,
        "margin_mm": MARGIN_MM,
    }
    url = f"{API_BASE}/ai_layout"
    print(f"POST {url}  (detections={len(detections)})")
    r = requests.post(url, json=payload, timeout=3600)
    print(f"HTTP {r.status_code}")
    data = r.json()
    print(json.dumps(data, ensure_ascii=False, indent=2))
    if data.get("output_url"):
        download_url(data["output_url"], os.path.join(out_dir, "ai_layout_output.pdf"))
    return data


def main():
    if len(sys.argv) < 3:
        print("用法: python client_demo.py [detect|pipeline|ai_layout] <本地PDF路径 或 预签名URL> [detections.json]")
        sys.exit(1)
    mode, source = sys.argv[1], sys.argv[2]
    if mode not in ("detect", "pipeline", "ai_layout"):
        print(f"未知模式: {mode} (可选 detect / pipeline / ai_layout)")
        sys.exit(1)

    if mode == "ai_layout":
        detections_file = sys.argv[3] if len(sys.argv) > 3 else None
        do_ai_layout(source, detections_file)
        return

    extra = None
    if mode == "pipeline":
        extra = {
            "layout_instruction": LAYOUT_INSTRUCTION,
            "box_size_mm": BOX_SIZE_MM,
            "margin_mm": MARGIN_MM,
        }
    call(mode, source, extra_payload=extra)


if __name__ == "__main__":
    main()
