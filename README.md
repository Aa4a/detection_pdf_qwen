# PDF 包装刀版检测与排版（Qwen 自部署版）

Windows 上的包装刀版处理服务：从 PDF 提取刀版线，使用自部署 Qwen 视觉模型辅助识别面板与摇盖，再生成检测可视化或自动排版结果。服务通过 FastAPI 暴露接口，并可通过 MinIO 接收输入和返回结果。

## 效果示例

### 示例一

![W1.2193 外箱检测结果](docs/images/example-w1.png)

### 示例二

![testd 检测结果](docs/images/example-testd.png)

## 运行要求

- Windows 10/11 或 Windows Server
- Python 3.10+
- Adobe Illustrator（检测和排版接口需要）
- Microsoft Excel 与 Microsoft Print to PDF（Excel 转 PDF 接口需要）
- 提供 OpenAI 兼容接口的自部署 Qwen 视觉模型
- MinIO（使用文件服务接口时需要）

## 安装

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

程序直接读取系统环境变量；PowerShell 启动示例：

```powershell
$env:QWEN_BASE_URL = "http://127.0.0.1:8000/v1"
$env:QWEN_API_KEY = ""
$env:QWEN_MODEL = "qwen-vl"
$env:QWEN_TIMEOUT = "180"

$env:MINIO_ENDPOINT = "127.0.0.1:9000"
$env:MINIO_ACCESS_KEY = ""
$env:MINIO_SECRET_KEY = ""
$env:MINIO_BUCKET = "sun"

python api_server.py
```

默认接口文档地址：<http://127.0.0.1:8003/docs>。

## Qwen 接口约定

`qwen_detection.py` 使用 OpenAI 兼容的 `chat/completions` 多模态接口，适用于常见的 vLLM、SGLang 等部署。模型需接受 `image_url` 内容，并返回以下 JSON：

```json
{
  "detections": [
    {"label": "第一正唛上摇盖", "box_2d": [100, 120, 300, 480]}
  ]
}
```

`box_2d` 顺序为 `[ymin, xmin, ymax, xmax]`，坐标范围为 `0..1000`。

## 主要接口

| 方法 | 路径 | 功能 |
| --- | --- | --- |
| GET | `/health` | 健康检查 |
| POST | `/detect` | 刀版检测与可视化 |
| POST | `/pipeline` | 清理、检测、去底色与自动排版 |
| POST | `/ai_layout` | 使用已有检测框直接排版 |
| POST | `/divide_pdf` | PDF 视觉区域切分 |
| POST | `/excel2pdf` | Excel 转单页 PDF |

检测请求中的模型字段为：

```json
{
  "pdf_url": "https://example.com/input.pdf",
  "use_qwen": true,
  "qwen_requirements": "按实际刀线边界检测"
}
```

## 核心文件

| 文件 | 用途 |
| --- | --- |
| `qwen_detection.py` | 自部署 Qwen 视觉模型适配器 |
| `qwen_adobe_detection.py` | Illustrator 刀版线提取与 Qwen 结果融合 |
| `api_server.py` | FastAPI 服务入口 |
| `pipeline.py` | 完整处理流水线 |
| `ai_toolkit_modifications.py` | 自动排版 |
| `utils/divide2_pdf_final.py` | PDF 区域切分 |
| `minio_helper.py` | MinIO 上传、下载与预签名 URL |

## 安全说明

仓库不包含模型密钥、MinIO 密钥、测试 PDF、日志或临时输出。请只在本机环境变量或未纳入 Git 的 `.env` 中配置凭据。
