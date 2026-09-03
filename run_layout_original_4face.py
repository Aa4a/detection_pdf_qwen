# -*- coding: utf-8 -*-
"""
原尺寸 4 面排版调用脚本

流程:
  1) 上传本地 PDF 到 MinIO 输入桶, 生成预签名 URL
  2) POST /detect 获取 final_boxes + dieline_size_mm
  3) 用检测出的刀版尺寸作为 box_size_mm (原尺寸)
  4) POST /ai_layout, 正唛->面1/3, 侧唛->面2/4
  5) 下载检测可视化 PDF 与排版结果 PDF/预览图

用法:
  python run_layout_original_4face.py "C:/path/to/input.pdf"
  python run_layout_original_4face.py "C:/path/to/input.pdf" --out-dir client_out

环境变量 (均可选, 有默认值):
  API_BASE              默认 http://127.0.0.1:8003
  MINIO_ENDPOINT        默认 192.168.0.64:9000
    MINIO_ACCESS_KEY      访问密钥, 必须通过环境变量设置
    MINIO_SECRET_KEY      私有密钥, 必须通过环境变量设置
  MINIO_SECURE          默认 false
  MINIO_INPUT_BUCKET    默认 inputfile
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from datetime import timedelta

import requests
from minio import Minio

API_BASE = os.environ.get("API_BASE", "http://127.0.0.1:8003")
MINIO_ENDPOINT = os.environ.get("MINIO_ENDPOINT", "192.168.0.64:9000")
MINIO_ACCESS_KEY = os.environ.get("MINIO_ACCESS_KEY", "")
MINIO_SECRET_KEY = os.environ.get("MINIO_SECRET_KEY", "")
MINIO_SECURE = os.environ.get("MINIO_SECURE", "false").lower() in ("1", "true", "yes", "on")
INPUT_BUCKET = os.environ.get("MINIO_INPUT_BUCKET", "inputfile")

# 4 面排版: 正唛落在 1/3, 侧唛落在 2/4
LAYOUT_INSTRUCTION = {
    "正唛内容": [1, 3],
    "侧唛内容": [2, 4],
    "正唛上摇盖": [1, 3],
    "侧唛上摇盖": [2, 4],
    "正唛下摇盖": [1, 3],
    "侧唛下摇盖": [2, 4],
}


def _minio() -> Minio:
    return Minio(
        MINIO_ENDPOINT,
        access_key=MINIO_ACCESS_KEY,
        secret_key=MINIO_SECRET_KEY,
        secure=MINIO_SECURE,
    )


def upload_and_presign(local_path: str) -> str:
    client = _minio()
    if not client.bucket_exists(INPUT_BUCKET):
        client.make_bucket(INPUT_BUCKET)
    object_name = os.path.basename(local_path)
    client.fput_object(
        INPUT_BUCKET, object_name, local_path, content_type="application/pdf"
    )
    url = client.presigned_get_object(
        INPUT_BUCKET, object_name, expires=timedelta(hours=2)
    )
    print(f"[上传] {INPUT_BUCKET}/{object_name}")
    return url


def download_url(url: str, save_path: str) -> str:
    r = requests.get(url, timeout=600)
    r.raise_for_status()
    os.makedirs(os.path.dirname(os.path.abspath(save_path)), exist_ok=True)
    with open(save_path, "wb") as f:
        f.write(r.content)
    print(f"[下载] {save_path} ({os.path.getsize(save_path)} bytes)")
    return save_path


def dieline_to_box_size_mm(dieline: dict) -> dict:
    """检测刀版尺寸 -> 排版用 box_size_mm (原尺寸)。"""
    return {
        "排版尺寸": [
            float(dieline.get("w_main_mm") or 0),
            float(dieline.get("w_side_mm") or 0),
            float(dieline.get("h_panel_mm") or 0),
            float(dieline.get("h_flap_mm") or 0),
        ]
    }


def call_detect(pdf_url: str) -> dict:
    url = f"{API_BASE.rstrip('/')}/detect"
    print(f"[1/2] POST {url}")
    r = requests.post(url, json={"pdf_url": pdf_url}, timeout=3600)
    print(f"HTTP {r.status_code}")
    data = r.json()
    print(
        json.dumps(
            {
                k: data.get(k)
                for k in ("status", "error", "elapsed_ms", "object_name", "output_url")
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    if data.get("status") != "success":
        raise RuntimeError(f"/detect 失败: {data.get('error')}")
    return data


def call_ai_layout(pdf_url: str, detections: list, box_size_mm: dict) -> dict:
    url = f"{API_BASE.rstrip('/')}/ai_layout"
    payload = {
        "pdf_url": pdf_url,
        "detections": detections,
        "layout_instruction": LAYOUT_INSTRUCTION,
        "box_size_mm": box_size_mm,
        "margin_mm": 0.0,
    }
    print(f"[2/2] POST {url}  (detections={len(detections)}, box_size_mm={box_size_mm})")
    r = requests.post(url, json=payload, timeout=3600)
    print(f"HTTP {r.status_code}")
    data = r.json()
    print(
        json.dumps(
            {
                k: data.get(k)
                for k in (
                    "status",
                    "error",
                    "elapsed_ms",
                    "object_name",
                    "output_url",
                    "console_url",
                )
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    if data.get("status") != "success":
        raise RuntimeError(f"/ai_layout 失败: {data.get('error')}")
    return data


def run(local_pdf: str, out_dir: str) -> None:
    if not os.path.isfile(local_pdf):
        raise FileNotFoundError(local_pdf)

    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(local_pdf))[0]

    pdf_url = upload_and_presign(local_pdf)

    det = call_detect(pdf_url)
    result = det.get("result") or {}
    boxes = result.get("final_boxes") or []
    dieline = result.get("dieline_size_mm") or {}

    print(f"boxes: {len(boxes)}")
    print(f"dieline_size_mm: {dieline}")

    if det.get("output_url"):
        download_url(
            det["output_url"], os.path.join(out_dir, f"{stem}_detect_vis.pdf")
        )

    if not result.get("success") or not boxes:
        raise RuntimeError("检测未成功或 final_boxes 为空, 无法排版")

    box_size_mm = dieline_to_box_size_mm(dieline)
    print(f"[原尺寸] box_size_mm = {box_size_mm}")

    lay = call_ai_layout(pdf_url, boxes, box_size_mm)

    if lay.get("output_url"):
        download_url(
            lay["output_url"],
            os.path.join(out_dir, f"{stem}_layout_original_4face.pdf"),
        )

    preview = ((lay.get("result") or {}).get("preview") or {})
    if preview.get("output_url"):
        download_url(
            preview["output_url"],
            os.path.join(out_dir, f"{stem}_layout_original_4face_preview.png"),
        )

    print("DONE")
    if lay.get("console_url"):
        print(f"MinIO 控制台: {lay['console_url']}")


def main() -> None:
    parser = argparse.ArgumentParser(description="原尺寸 4 面排版 (detect -> ai_layout)")
    parser.add_argument("pdf", help="本地 PDF 路径")
    parser.add_argument(
        "--out-dir",
        default="client_out",
        help="结果下载目录 (默认 client_out)",
    )
    args = parser.parse_args()
    try:
        run(os.path.abspath(args.pdf), os.path.abspath(args.out_dir))
    except Exception as e:
        print(f"[失败] {type(e).__name__}: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
