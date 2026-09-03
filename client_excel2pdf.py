# -*- coding: utf-8 -*-
"""
Excel 转 PDF 接口的客户端调用脚本 (对接远程 MinIO + 本服务 /excel2pdf)

流程:
    1) 把本地 Excel 上传到 MinIO 输入桶, 生成 1 小时预签名下载 URL (excel_url)
       (若直接传入 http(s) 预签名 URL, 则跳过上传)
    2) POST {excel_url} 到服务 /excel2pdf
    3) 服务用 Excel + Illustrator 转成单页 PDF, 上传 MinIO, 返回 output_url + console_url
    4) 用 output_url 下载结果 PDF

用法:
    python client_excel2pdf.py "C:/path/to/book.xlsx"
    python client_excel2pdf.py "http://192.168.0.64:9000/inputfile/book.xlsx?X-Amz-..."

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

XLSX_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _minio():
    return Minio(MINIO_ENDPOINT, access_key=MINIO_ACCESS_KEY,
                 secret_key=MINIO_SECRET_KEY, secure=MINIO_SECURE)


def upload_and_presign(local_path):
    """上传本地 Excel 到输入桶, 返回 1 小时有效的预签名下载 URL。"""
    client = _minio()
    if not client.bucket_exists(INPUT_BUCKET):
        client.make_bucket(INPUT_BUCKET)
    object_name = os.path.basename(local_path)
    client.fput_object(INPUT_BUCKET, object_name, local_path, content_type=XLSX_CT)
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


def resolve_url(source):
    """source 是 http(s) URL 则直接用; 否则当作本地文件上传后取预签名 URL。"""
    if source.lower().startswith(("http://", "https://")):
        return source
    if not os.path.isfile(source):
        print(f"文件不存在: {source}")
        sys.exit(1)
    return upload_and_presign(source)


def call(source, out_dir="client_out"):
    excel_url = resolve_url(source)
    payload = {"excel_url": excel_url}

    url = f"{API_BASE}/excel2pdf"
    print(f"POST {url}  (excel_url={excel_url[:90]}...)")
    r = requests.post(url, json=payload, timeout=3600)
    print(f"HTTP {r.status_code}")
    data = r.json()
    print(json.dumps(data, ensure_ascii=False, indent=2))

    if data.get("output_url"):
        download_url(data["output_url"], os.path.join(out_dir, "excel2pdf_output.pdf"))
    return data


def main():
    if len(sys.argv) < 2:
        print("用法: python client_excel2pdf.py <本地Excel路径 或 预签名URL>")
        sys.exit(1)
    call(sys.argv[1])


if __name__ == "__main__":
    main()
