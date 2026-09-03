# PDF 刀版检测 / 排版 / Excel 转 PDF 服务接口文档

基于 FastAPI 的处理服务，文件传输统一走 MinIO：
**输入**为 MinIO 预签名下载 URL（服务直接 HTTP 下载，无需凭证）；
**输出**结果文件由服务上传回 MinIO，并返回 7 天有效的预签名下载 URL。

---

## 1. 基本信息

| 项目 | 值 |
| --- | --- |
| 服务地址 | `http://192.168.0.48:8003` |
| 协议 | HTTP / JSON |
| 交互式文档 | `http://192.168.0.48:8003/docs` (Swagger UI) |
| OpenAPI | `http://192.168.0.48:8003/openapi.json` |
| MinIO 地址 | `192.168.0.64:9000` (控制台 `192.168.0.64:9001`) |
| 结果输出桶 | `sun` (可按请求覆盖) |

> 底层依赖 Adobe Illustrator / Excel（COM 单实例），服务内部对所有任务**串行**执行，同一时刻只处理一个请求。

---

## 2. 统一响应结构

除 `/health` 外，所有接口返回同一结构：

```json
{
  "status": "success | error",
  "output_url": "结果文件预签名下载 URL (成功时)",
  "console_url": "MinIO 控制台浏览 URL",
  "bucket": "结果输出桶名",
  "object_name": "结果文件在桶内的对象名",
  "result": { "各接口的业务结果 (见下)" },
  "elapsed_ms": 8804,
  "error": "错误信息 (失败时), 成功时为 null"
}
```

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `status` | string | `success` 或 `error` |
| `output_url` | string\|null | 结果文件预签名下载 URL，有效期 7 天 |
| `console_url` | string\|null | MinIO 控制台浏览地址 |
| `bucket` | string | 结果输出桶 |
| `object_name` | string\|null | 结果对象名（形如 `原名_8位随机.pdf`） |
| `result` | object | 各接口特有的业务结果 |
| `elapsed_ms` | int | 服务端处理耗时（毫秒） |
| `error` | string\|null | 失败原因 |

---

## 3. 接口列表

### 3.1 健康检查

```
GET /health
```

响应：

```json
{ "status": "ok" }
```

---

### 3.2 刀版线检测 `POST /detect`

用 Adobe Illustrator 识别刀版线并补全 12 格，输出叠加彩色检测框的可视化 PDF。

**请求体**

| 字段 | 类型 | 必填 | 默认 | 说明 |
| --- | --- | --- | --- | --- |
| `pdf_url` | string | 是 | — | 输入 PDF 的 MinIO 预签名下载 URL |
| `bucket` | string | 否 | `sun` | 结果输出桶 |
| `use_qwen` | bool | 否 | `true` | 是否用 Qwen 辅助判定行结构 |
| `qwen_requirements` | string | 否 | `""` | 传给 Qwen 的额外要求描述 |

```json
{
  "pdf_url": "http://192.168.0.64:9000/inputfile/box.pdf?X-Amz-Algorithm=..."
}
```

**`result` 字段**

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `success` | bool | 检测是否成功 |
| `final_boxes` | array | 检测到的 12 格，每项 `{box:[x0,y0,x1,y1], class_name, label_source}`（zoom=2 图像坐标） |
| `dieline_size_mm` | object | `{w_main_mm, w_side_mm, h_panel_mm, h_flap_mm}` |
| `rows_skipped` | array | 被跳过的行号 |

`output_url` 指向可视化 PDF。

---

### 3.3 完整排版流水线 `POST /pipeline`

端到端：规范化 PDF（剥离 Illustrator 私有数据）→ 刀版线检测 → 去除铺满刀版格的底色 → 排版，输出排版后的 PDF。

**请求体**

| 字段 | 类型 | 必填 | 默认 | 说明 |
| --- | --- | --- | --- | --- |
| `pdf_url` | string | 是 | — | 输入 PDF 的 MinIO 预签名下载 URL |
| `bucket` | string | 否 | `sun` | 结果输出桶 |
| `layout_instruction` | object | 否 | 见下 | 排版指令：每类版面 → 目标格号列表 |
| `box_size_mm` | object | 否 | 见下 | 排版尺寸 `{"排版尺寸": [L, W, H, 摇盖高]}`（mm） |
| `margin_mm` | float | 否 | `0.0` | 排版留白（mm） |
| `use_qwen` | bool | 否 | `true` | 是否用 Qwen 辅助 |
| `qwen_requirements` | string | 否 | `""` | Qwen 额外要求 |
| `strip_illustrator` | bool | 否 | `true` | 是否先剥离 Illustrator 私有数据 |

`layout_instruction` / `box_size_mm` 不传时使用的服务端默认值：

```json
{
  "layout_instruction": {
    "正唛内容":  [1, 3],
    "侧唛内容":  [2, 4],
    "正唛上摇盖": [1, 3],
    "侧唛上摇盖": [2, 4],
    "正唛下摇盖": [1, 3],
    "侧唛下摇盖": [2, 4]
  },
  "box_size_mm": { "排版尺寸": [79.87, 55.87, 55.87, 20] }
}
```

**完整请求示例**

```json
{
  "pdf_url": "http://192.168.0.64:9000/inputfile/box.pdf?X-Amz-...",
  "layout_instruction": {
    "正唛内容":  [1, 3],
    "侧唛内容":  [2, 4],
    "正唛上摇盖": [1, 3],
    "侧唛上摇盖": [2, 4],
    "正唛下摇盖": [1, 3],
    "侧唛下摇盖": [2, 4]
  },
  "box_size_mm": { "排版尺寸": [79.87, 55.87, 55.87, 20] },
  "margin_mm": 0.0
}
```

**`result` 字段**

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `detection` | object | 检测结果 `{success, final_boxes, dieline_size_mm, rows_skipped}` |
| `preview` | object | 排版预览图（PNG），`{object_name, output_url, console_url}`（若生成） |

`output_url` 指向排版后的 PDF。

---

### 3.4 Excel 转 PDF `POST /excel2pdf`

后端使用 `Excel转PDF_免费离线版_20260820`（`Excel转PDF_免费离线.exe`）：
Microsoft Excel + Windows 自带 **Microsoft Print to PDF**，单页 A4。不需要 PDFCreator / PDF-XChange / WPS。需本机已安装桌面版 Microsoft Excel。

**请求体**

| 字段 | 类型 | 必填 | 默认 | 说明 |
| --- | --- | --- | --- | --- |
| `excel_url` | string | 是 | — | 输入 Excel（.xlsx/.xls）的 MinIO 预签名下载 URL |
| `bucket` | string | 否 | `sun` | 结果输出桶 |
| `include_hidden` | bool | 否 | `false` | `true` 时同时转换隐藏但有内容的 Sheet |

```json
{
  "excel_url": "http://192.168.0.64:9000/inputfile/book.xlsx?X-Amz-..."
}
```

**`result` 字段**

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `success` | bool | 转换是否成功 |
| `source_name` | string | 源 Excel 文件名 |
| `engine` | string | 算法包名称 |
| `include_hidden` | bool | 是否包含隐藏 Sheet |

`output_url` 指向生成的 PDF。

**响应示例**

```json
{
  "status": "success",
  "output_url": "http://192.168.0.64:9000/sun/book_7c8c695d.pdf?X-Amz-...",
  "console_url": "http://192.168.0.64:9001/browser/sun/book_7c8c695d.pdf",
  "bucket": "sun",
  "object_name": "book_7c8c695d.pdf",
  "result": { "success": true, "source_name": "book.xlsx" },
  "elapsed_ms": 8804,
  "error": null
}
```

---

### 3.5 给定检测框排版 `POST /ai_layout`

用调用方提供的检测框直接排版。
与 `/pipeline` 的区别：**本接口不做检测**，`detections` 由调用方传入（通常先调 `/detect` 拿到 `result.final_boxes` 再传进来），也不做去底色。

- `excel=false`（默认）：`ai_toolkit_modifications`（贴刀版线）
- `excel=true`：`manual_box_layout`（严格按检测框，适合 `/excel2pdf` 产出的源稿）

**请求体**

| 字段 | 类型 | 必填 | 默认 | 说明 |
| --- | --- | --- | --- | --- |
| `pdf_url` | string | 是 | — | 输入 PDF 的 MinIO 预签名下载 URL |
| `detections` | array | 是 | — | 检测框列表，每项 `{box:[x0,y0,x1,y1], class_name}`（zoom=2 图像坐标） |
| `bucket` | string | 否 | `sun` | 结果输出桶 |
| `layout_instruction` | object | 否 | 同 `/pipeline` 默认 | 排版指令 |
| `box_size_mm` | object | 否 | 同 `/pipeline` 默认 | 排版尺寸 `{"排版尺寸": [L, W, H, 摇盖高]}`（mm） |
| `margin_mm` | float | 否 | `0.0` | 排版留白（mm） |
| `excel` | bool | 否 | `false` | `true` 时走手动框严格排版（`manual_box_layout`） |

`detections` 中 `class_name` 取值：`正唛内容 / 侧唛内容 / 正唛上摇盖 / 侧唛上摇盖 / 正唛下摇盖 / 侧唛下摇盖`。

**请求示例**

```json
{
  "pdf_url": "http://192.168.0.64:9000/inputfile/box.pdf?X-Amz-...",
  "detections": [
    { "box": [1166.8, 1983.3, 3179.4, 2989.6], "class_name": "侧唛上摇盖" },
    { "box": [3179.4, 1983.3, 5345.0, 2989.6], "class_name": "正唛上摇盖" },
    { "box": [1166.8, 2989.6, 3179.4, 9310.8], "class_name": "侧唛内容" },
    { "box": [3179.4, 2989.6, 5345.0, 9310.8], "class_name": "正唛内容" }
  ],
  "box_size_mm": { "排版尺寸": [335, 240, 510, 120] },
  "margin_mm": 0.0
}
```

**`result` 字段**

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `detections_count` | int | 传入的检测框数量 |
| `preview` | object | 排版预览图（PNG），`{object_name, output_url, console_url}`（若生成） |

`output_url` 指向排版后的 PDF。

---

## 4. 调用流程

1. 客户端把本地文件（PDF/Excel）上传到 MinIO 输入桶（如 `inputfile`）；
2. 生成一个短期有效的预签名下载 URL（`pdf_url` / `excel_url`）；
3. POST 到对应接口；
4. 服务下载 → 处理 → 上传结果 → 返回 `output_url`；
5. 客户端用 `output_url` 直接下载结果。

### 4.1 curl 示例

```bash
curl -X POST http://192.168.0.48:8003/excel2pdf \
  -H "Content-Type: application/json" \
  -d '{"excel_url":"http://192.168.0.64:9000/inputfile/book.xlsx?X-Amz-..."}'
```

### 4.2 Python 客户端

仓库已提供两个客户端脚本，二者都支持传**本地文件路径**（自动上传并预签名）或**已有预签名 URL**：

```bash
# 检测 / 排版
python client_demo.py detect   "C:\path\to\input.pdf"
python client_demo.py pipeline "C:\path\to\input.pdf"

# 给定检测框排版 (不带 detections.json 时会先调 /detect 获取)
python client_demo.py ai_layout "C:\path\to\input.pdf"
python client_demo.py ai_layout "C:\path\to\input.pdf" "C:\path\to\detections.json"

# Excel 转 PDF
python client_excel2pdf.py "C:\path\to\book.xlsx"

# /ai_layout 内联检测框的完整示例 (检测框写死在脚本里, 不读任何文件)
python example_ai_layout.py
```

`client_demo.py` 顶部可修改 `LAYOUT_INSTRUCTION` / `BOX_SIZE_MM` / `MARGIN_MM`，`pipeline` 与 `ai_layout` 模式会随请求发送。

> **关于 `detections` 的说明**：`/ai_layout` 的 `detections` 是**请求体里的 JSON 数组字段**，不是必须的文件。
> 来源随意：可先调 `/detect` 拿 `result.final_boxes` 转手传入、写死在代码里（见 `example_ai_layout.py`）、或从本地文件读取。
> `client_demo.py ai_layout` 的第三个参数 `detections.json` 只是一个可选便捷入口，支持 `[...]` 或 `{"final_boxes": [...]}` 两种格式；不传时脚本会自动先调 `/detect` 获取。

---

## 5. 服务端启动与配置

```powershell
$env:API_PORT="8003"
$env:PYTHONIOENCODING="utf-8"
$env:MINIO_ENDPOINT="192.168.0.64:9000"
$env:MINIO_ACCESS_KEY=""
$env:MINIO_SECRET_KEY=""
$env:MINIO_BUCKET="sun"
$env:MINIO_CONSOLE_ENDPOINT="192.168.0.64:9001"
python api_server.py
```

**环境变量**

| 变量 | 默认 | 说明 |
| --- | --- | --- |
| `API_HOST` | `0.0.0.0` | 监听地址 |
| `API_PORT` | `8001` | 监听端口（当前部署用 `8003`） |
| `MINIO_ENDPOINT` | `192.168.0.64:9000` | MinIO S3 地址 |
| `MINIO_ACCESS_KEY` | （空） | 访问密钥，必须通过环境变量设置 |
| `MINIO_SECRET_KEY` | （空） | 私有密钥，**上传结果必填** |
| `MINIO_SECURE` | `false` | 是否 https |
| `MINIO_BUCKET` | `sun` | 结果输出桶 |
| `MINIO_CONSOLE_ENDPOINT` | `192.168.0.64:9001` | 控制台地址 |
| `MINIO_PRESIGN_DAYS` | `7` | 输出预签名 URL 有效天数 |

---

## 6. 错误处理

- 请求参数缺失/格式错误：HTTP `422`（FastAPI 校验）。
- 处理过程出错：HTTP `200`，但响应体 `status="error"`，`error` 字段为原因，`output_url=null`。
- 常见错误原因：输入 URL 无法下载、Illustrator/Excel 未安装或未授权、检测失败（非规整刀版图）、排版参数不合法等。

---

## 7. 备注

- 服务内所有 AI 任务串行执行，高并发请求会排队；
- 结果对象名格式为 `原文件名_8位随机后缀.扩展名`，避免覆盖；
- 输入预签名 URL 建议短有效期（如 1 小时），输出默认 7 天。
```
