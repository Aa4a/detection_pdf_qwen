# -*- coding: utf-8 -*-
"""
MinIO 文件传输封装 (对接远程 MinIO 服务器)

契约与既有服务一致:
    - 输入: 服务收到的是 MinIO 预签名下载 URL (pdf_url), 直接 HTTP GET 即可, 无需凭证
    - 输出: 结果文件上传到 MinIO 桶, 返回 7 天有效的预签名下载 URL + 控制台 URL

配置 (环境变量优先, 否则用下方默认值):
    MINIO_ENDPOINT          MinIO S3 地址, 形如 "192.168.0.64:9000" (不带 http://)
    MINIO_ACCESS_KEY        访问密钥 (默认 admin)
    MINIO_SECRET_KEY        私有密钥 (上传结果时必填, 请通过环境变量设置)
    MINIO_SECURE            是否 https, "true"/"false" (默认 false)
    MINIO_BUCKET            结果输出桶 (默认 sun)
    MINIO_CONSOLE_ENDPOINT  控制台地址 (默认 192.168.0.64:9001)
    MINIO_PRESIGN_DAYS      输出预签名 URL 有效天数 (默认 7)
"""

import os
from datetime import timedelta

import requests
from minio import Minio


def _env_bool(name, default=False):
    val = os.environ.get(name)
    if val is None:
        return default
    return val.strip().lower() in ("1", "true", "yes", "on")


MINIO_ENDPOINT = os.environ.get("MINIO_ENDPOINT", "192.168.0.64:9000")
MINIO_ACCESS_KEY = os.environ.get("MINIO_ACCESS_KEY", "")
MINIO_SECRET_KEY = os.environ.get("MINIO_SECRET_KEY", "")  # 必填: 通过环境变量提供
MINIO_SECURE = _env_bool("MINIO_SECURE", False)
DEFAULT_BUCKET = os.environ.get("MINIO_BUCKET", "sun")
MINIO_CONSOLE_ENDPOINT = os.environ.get("MINIO_CONSOLE_ENDPOINT", "192.168.0.64:9001")

PRESIGN_EXPIRES = timedelta(days=int(os.environ.get("MINIO_PRESIGN_DAYS", "7")))


_client = None


def get_client():
    """返回全局单例 MinIO 客户端。"""
    global _client
    if _client is None:
        _client = Minio(
            MINIO_ENDPOINT,
            access_key=MINIO_ACCESS_KEY,
            secret_key=MINIO_SECRET_KEY,
            secure=MINIO_SECURE,
        )
    return _client


def ensure_bucket(bucket=None):
    """确保桶存在, 不存在则创建。"""
    bucket = bucket or DEFAULT_BUCKET
    client = get_client()
    if not client.bucket_exists(bucket):
        client.make_bucket(bucket)
    return bucket


def download_url(url, local_path):
    """从预签名 URL 直接 HTTP 下载到本地 (无需 MinIO 凭证)。"""
    os.makedirs(os.path.dirname(os.path.abspath(local_path)), exist_ok=True)
    with requests.get(url, stream=True, timeout=300) as r:
        r.raise_for_status()
        with open(local_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=1 << 20):
                if chunk:
                    f.write(chunk)
    return local_path


def upload(local_path, object_name, bucket=None, content_type="application/octet-stream"):
    """上传本地文件到 MinIO, 返回 object_name。"""
    bucket = bucket or DEFAULT_BUCKET
    ensure_bucket(bucket)
    get_client().fput_object(bucket, object_name, local_path, content_type=content_type)
    return object_name


def presigned_url(object_name, bucket=None, expires=None):
    """生成用于下载的预签名 URL (客户端无需 MinIO 凭证即可下载)。"""
    bucket = bucket or DEFAULT_BUCKET
    expires = expires or PRESIGN_EXPIRES
    return get_client().presigned_get_object(bucket, object_name, expires=expires)


def console_url(object_name, bucket=None):
    """生成 MinIO 控制台浏览 URL。"""
    bucket = bucket or DEFAULT_BUCKET
    scheme = "https" if MINIO_SECURE else "http"
    return f"{scheme}://{MINIO_CONSOLE_ENDPOINT}/browser/{bucket}/{object_name}"


def guess_content_type(path):
    """根据扩展名推断 content-type。"""
    ext = os.path.splitext(path)[1].lower()
    return {
        ".pdf": "application/pdf",
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
    }.get(ext, "application/octet-stream")
