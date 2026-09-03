# -*- coding: utf-8 -*-
"""
FastAPI 服务: 刀版线检测 + 完整排版流水线 (对接远程 MinIO)

统一契约 (与既有服务一致):
    请求:  { "pdf_url": "<MinIO 预签名下载 URL>", ...可选参数 }
    响应:  {
              "status": "success" | "error",
              "output_url": "<结果 PDF 的预签名下载 URL>",
              "console_url": "<MinIO 控制台浏览 URL>",
              "bucket": "...",
              "object_name": "...",
              "result": { ...业务结果... },
              "elapsed_ms": 1234,
              "error": null | "..."
           }

接口:
    POST /detect     仅检测刀版线 + 可视化 (qwen_adobe_detection.visualize_dielines)
    POST /pipeline   完整流水线: 清理 -> 检测 -> 去底色 -> 排版 (pipeline.run_pipeline)
    GET  /health     健康检查

说明:
    - 输入 pdf_url 是预签名 URL, 服务直接 HTTP 下载, 无需 MinIO 凭证
    - 结果上传 MinIO 需要凭证, 请通过环境变量 MINIO_SECRET_KEY 提供
    - 仅需传 pdf_url 即可运行, 排版参数不传则用下方默认值
    - 底层依赖 Adobe Illustrator (COM, 单实例) -> 单线程执行器 + 锁, 所有任务串行
"""

import os
import shutil
import sys
import tempfile
import threading
import time
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Dict, List, Optional
from urllib.parse import unquote, urlparse

# 底层脚本 (excel_to_pdf_utils / pipeline 等) 有 emoji 打印, Windows 控制台默认 GBK 会
# 抛 UnicodeEncodeError, 这里把标准输出/错误统一改成 UTF-8 (无法编码时替换而非崩溃)。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# 给每一行 stdout/stderr 自动加时间戳前缀 (uvicorn 日志与所有 print 均生效)。
from datetime import datetime as _dt


class _TimestampedStream:
    """按行为标准输出/错误加 [时间] 前缀的包装器。"""

    def __init__(self, stream):
        self._stream = stream
        self._at_line_start = True
        self._ts_wrapped = True

    def write(self, data):
        if not data:
            return 0
        parts = []
        for line in data.splitlines(keepends=True):
            if self._at_line_start:
                parts.append(_dt.now().strftime("[%Y-%m-%d %H:%M:%S] "))
            parts.append(line)
            self._at_line_start = line.endswith(("\n", "\r"))
        self._stream.write("".join(parts))
        return len(data)

    def flush(self):
        self._stream.flush()

    def __getattr__(self, name):
        return getattr(self._stream, name)


# 防止重复包装 (python api_server.py 以 __main__ 加载一次, uvicorn 又按 "api_server:app"
# 再导入一次, 顶层代码会重跑; 用属性标记跨模块判断, isinstance 因类身份不同不可靠)。
if not getattr(sys.stdout, "_ts_wrapped", False):
    sys.stdout = _TimestampedStream(sys.stdout)
if not getattr(sys.stderr, "_ts_wrapped", False):
    sys.stderr = _TimestampedStream(sys.stderr)

import pythoncom
from fastapi import FastAPI
from pydantic import BaseModel, Field

import minio_helper
from ai_toolkit_modifications import apply_layout_pure_python as ai_apply_layout
from qwen_adobe_detection import visualize_dielines
from manual_box_layout import apply_layout_manual_boxes
from pipeline import run_pipeline
from utils.divide2_pdf_final import divide_pdf

# Excel→PDF: 免费离线版 (Microsoft Print to PDF, 无需 PDFCreator)
_EXCEL2PDF_ROOT = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "Excel转PDF_免费离线版_20260820",
)
_EXCEL2PDF_EXE = os.path.join(_EXCEL2PDF_ROOT, "Excel转PDF_免费离线.exe")
_EXCEL2PDF_PRINTER = "Microsoft Print to PDF"


def excel_to_pdf_latest(input_excel_path, output_folder=None, include_hidden=False):
    """调用免费离线版 EXE (Windows 自带 Microsoft Print to PDF)。

    返回生成的 PDF 路径; 失败返回 None。
    """
    import subprocess
    from datetime import datetime
    from pathlib import Path

    import win32print

    exe = Path(_EXCEL2PDF_EXE)
    root = Path(_EXCEL2PDF_ROOT)
    if not exe.is_file():
        raise FileNotFoundError(f"找不到 Excel 转 PDF EXE: {exe}")

    printers = {
        str(item[2])
        for item in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )
    }
    if _EXCEL2PDF_PRINTER not in printers:
        raise RuntimeError(
            f"没有找到打印机 {_EXCEL2PDF_PRINTER}。请在 Windows 功能中启用“Microsoft Print to PDF”。"
        )

    src = Path(input_excel_path).resolve()
    if not src.is_file():
        raise FileNotFoundError(f"Excel 不存在: {src}")

    out_dir = Path(output_folder) if output_folder else (src.parent / "output")
    out_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="excel2pdf_svc_") as tmp:
        tmp_path = Path(tmp)
        in_dir = tmp_path / "in"
        work_out = tmp_path / "out"
        in_dir.mkdir()
        work_out.mkdir()
        shutil.copy2(src, in_dir / src.name)

        cmd = [str(exe), "--input", str(in_dir), "--output", str(work_out)]
        if include_hidden:
            cmd.append("--include-hidden")

        print(
            f"[excel2pdf] 免费离线版: {exe.name}  printer={_EXCEL2PDF_PRINTER}  src={src.name}",
            flush=True,
        )
        proc = subprocess.run(
            cmd,
            cwd=str(root),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if proc.stdout:
            print(proc.stdout, flush=True)
        if proc.stderr:
            print(proc.stderr, flush=True)

        pdfs = sorted(work_out.glob("*.pdf"))
        if not pdfs:
            tail = (proc.stdout or "").strip().splitlines()[-8:]
            print(f"[excel2pdf] 未生成 PDF (exit={proc.returncode}) tail={tail}", flush=True)
            return None

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dest = out_dir / f"{src.stem}_{stamp}.pdf"
        shutil.copy2(pdfs[0], dest)
        print(f"[excel2pdf] 成功: {dest} (exit={proc.returncode})", flush=True)
        return str(dest)


app = FastAPI(title="PDF 刀版检测 / 排版服务", version="2.0.0")

# ------------------------------------------------------------------
# 排版默认参数 (仅传 pdf_url 时使用, 可被请求覆盖)
# ------------------------------------------------------------------
DEFAULT_LAYOUT_INSTRUCTION = {
    "正唛内容": [1, 3],
    "侧唛内容": [2, 4],
    "正唛上摇盖": [1, 3],
    "侧唛上摇盖": [2, 4],
    "正唛下摇盖": [1, 3],
    "侧唛下摇盖": [2, 4],
}
DEFAULT_BOX_SIZE_MM = {"排版尺寸": [79.87, 55.87, 55.87, 20]}


# ------------------------------------------------------------------
# COM 串行执行器: Illustrator 是单实例, 所有依赖 AI 的任务必须串行
# ------------------------------------------------------------------
def _com_init():
    pythoncom.CoInitialize()


_executor = ThreadPoolExecutor(max_workers=1, initializer=_com_init)
_job_lock = threading.Lock()


def _run_serial(fn, *args, **kwargs):
    with _job_lock:
        future = _executor.submit(fn, *args, **kwargs)
        return future.result()


# ------------------------------------------------------------------
# 请求 / 响应模型
# ------------------------------------------------------------------
class DetectRequest(BaseModel):
    pdf_url: str = Field(..., description="MinIO 预签名下载 URL")
    bucket: Optional[str] = Field(None, description="结果输出桶, 缺省用服务端默认桶")
    use_qwen: bool = Field(True, description="是否用 Qwen 辅助行结构判定")
    qwen_requirements: str = Field("", description="传给 Qwen 的额外要求描述")


class PipelineRequest(BaseModel):
    pdf_url: str = Field(..., description="MinIO 预签名下载 URL")
    bucket: Optional[str] = Field(None, description="结果输出桶, 缺省用服务端默认桶")
    layout_instruction: Optional[Dict[str, Any]] = Field(None, description="排版指令, 不传用默认")
    box_size_mm: Optional[Dict[str, Any]] = Field(None, description='排版尺寸, 不传用默认')
    margin_mm: float = Field(0.0, description="排版留白 (mm)")
    use_qwen: bool = Field(True)
    qwen_requirements: str = Field("")
    strip_illustrator: bool = Field(True, description="是否先剥离 Illustrator 私有数据")


class Excel2PdfRequest(BaseModel):
    excel_url: str = Field(..., description="MinIO 预签名下载 URL (指向 .xlsx/.xls)")
    bucket: Optional[str] = Field(None, description="结果输出桶, 缺省用服务端默认桶")
    include_hidden: bool = Field(
        False,
        description="True=同时转换隐藏但有内容的 Sheet; 默认只转可见且有内容的 Sheet",
    )


class DividePdfRequest(BaseModel):
    pdf_url: str = Field(..., description="MinIO 预签名下载 URL（指向 PDF）")
    bucket: Optional[str] = Field(None, description="结果输出桶，缺省使用服务端默认桶")
    zoom_factor: float = Field(2.0, gt=0, le=8, description="首页渲染缩放倍数")
    grid_width_ratio: float = Field(0.15, gt=0, le=1, description="网格最小宽度比例")
    grid_height_px: int = Field(40, ge=1, description="网格最小高度（像素）")
    trim_whitespace: bool = Field(True, description="是否裁掉分片四周白边")
    trim_margin_pt: float = Field(5.0, ge=0, description="裁白边后保留边距（PDF 点）")


class AiLayoutRequest(BaseModel):
    pdf_url: str = Field(..., description="MinIO 预签名下载 URL")
    detections: List[Dict[str, Any]] = Field(
        ..., description="检测框列表, 每项 {box:[x0,y0,x1,y1], class_name}, 坐标为 zoom=2 图像坐标"
    )
    bucket: Optional[str] = Field(None, description="结果输出桶, 缺省用服务端默认桶")
    layout_instruction: Optional[Dict[str, Any]] = Field(None, description="排版指令, 不传用默认")
    box_size_mm: Optional[Dict[str, Any]] = Field(None, description="排版尺寸, 不传用默认")
    margin_mm: float = Field(0.0, description="排版留白 (mm)")
    excel: bool = Field(
        False,
        description="True=Excel 源稿手动框排版 (manual_box_layout, 严格按检测框);"
                    " False=刀版贴合排版 (ai_toolkit_modifications)",
    )


class ServiceResponse(BaseModel):
    status: str
    output_url: Optional[str] = None
    console_url: Optional[str] = None
    bucket: Optional[str] = None
    object_name: Optional[str] = None
    result: Dict[str, Any] = {}
    elapsed_ms: int = 0
    error: Optional[str] = None


# ------------------------------------------------------------------
# 工具函数
# ------------------------------------------------------------------
def _new_workdir():
    d = os.path.join(tempfile.gettempdir(), "pdf_svc", uuid.uuid4().hex)
    os.makedirs(d, exist_ok=True)
    return d


def _name_from_url(pdf_url, default="input.pdf"):
    """从 URL 路径解析出文件名 (去掉查询串)。"""
    path = urlparse(pdf_url).path
    name = os.path.basename(unquote(path))
    return name or default


def _upload_output(local_path, bucket, stem_hint=None):
    """上传结果文件并返回 (object_name, output_url, console_url)。"""
    stem = stem_hint or os.path.splitext(os.path.basename(local_path))[0]
    ext = os.path.splitext(local_path)[1] or ".bin"
    object_name = f"{stem}_{uuid.uuid4().hex[:8]}{ext}"
    minio_helper.upload(
        local_path, object_name, bucket=bucket,
        content_type=minio_helper.guess_content_type(local_path),
    )
    return (
        object_name,
        minio_helper.presigned_url(object_name, bucket=bucket),
        minio_helper.console_url(object_name, bucket=bucket),
    )


# ------------------------------------------------------------------
# 路由
# ------------------------------------------------------------------
@app.on_event("startup")
def _startup():
    try:
        minio_helper.ensure_bucket()
    except Exception as e:
        print(f"[启动] 警告: 无法确认 MinIO 桶存在 (稍后请求时会重试): {e}")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/detect", response_model=ServiceResponse)
def detect(req: DetectRequest):
    """仅检测刀版线 + 生成可视化 PDF。"""
    t0 = time.time()
    bucket = req.bucket or minio_helper.DEFAULT_BUCKET
    workdir = _new_workdir()
    in_name = _name_from_url(req.pdf_url)
    stem = os.path.splitext(in_name)[0]
    try:
        input_pdf = os.path.join(workdir, in_name)
        minio_helper.download_url(req.pdf_url, input_pdf)

        vis_pdf = os.path.join(workdir, f"{stem}_vis.pdf")
        result = _run_serial(
            visualize_dielines,
            input_path=input_pdf,
            output_path=vis_pdf,
            use_qwen=req.use_qwen,
            qwen_requirements=req.qwen_requirements,
        )

        summary = {
            "success": bool(result.get("success")),
            "final_boxes": result.get("final_boxes", []),
            "dieline_size_mm": result.get("dieline_size_mm", {}),
            "rows_skipped": result.get("rows_skipped", []),
        }

        resp = ServiceResponse(status="success", result=summary, bucket=bucket)
        if summary["success"] and os.path.exists(vis_pdf):
            resp.object_name, resp.output_url, resp.console_url = _upload_output(vis_pdf, bucket, stem)
        else:
            resp.status = "error"
            resp.error = result.get("error", "检测失败")
        resp.elapsed_ms = int((time.time() - t0) * 1000)
        return resp
    except Exception as e:
        print(f"[错误] 请求处理失败: {type(e).__name__}: {e}", flush=True)
        traceback.print_exc()
        return ServiceResponse(
            status="error", bucket=bucket, error=f"{type(e).__name__}: {e}",
            elapsed_ms=int((time.time() - t0) * 1000),
        )
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


@app.post("/pipeline", response_model=ServiceResponse)
def pipeline(req: PipelineRequest):
    """完整流水线: 清理 -> 检测 -> 去底色 -> 排版。"""
    t0 = time.time()
    bucket = req.bucket or minio_helper.DEFAULT_BUCKET
    workdir = _new_workdir()
    in_name = _name_from_url(req.pdf_url)
    stem = os.path.splitext(in_name)[0]
    try:
        input_pdf = os.path.join(workdir, in_name)
        minio_helper.download_url(req.pdf_url, input_pdf)

        output_dir = os.path.join(workdir, "result")
        result = _run_serial(
            run_pipeline,
            input_pdf=input_pdf,
            layout_instruction=req.layout_instruction or DEFAULT_LAYOUT_INSTRUCTION,
            box_size_mm=req.box_size_mm or DEFAULT_BOX_SIZE_MM,
            output_dir=output_dir,
            margin_mm=req.margin_mm,
            use_qwen=req.use_qwen,
            qwen_requirements=req.qwen_requirements,
            strip_illustrator=req.strip_illustrator,
            cleaned_dir=os.path.join(output_dir, "_cleaned"),
        )

        det = result.get("detection") or {}
        summary = {
            "detection": {
                "success": bool(det.get("success")),
                "final_boxes": det.get("final_boxes", []),
                "dieline_size_mm": det.get("dieline_size_mm", {}),
                "rows_skipped": det.get("rows_skipped", []),
            },
        }

        layout_pdf = result.get("layout_pdf")
        layout_preview = result.get("layout_preview")

        resp = ServiceResponse(status="success", result=summary, bucket=bucket)
        if layout_pdf and os.path.exists(layout_pdf):
            resp.object_name, resp.output_url, resp.console_url = _upload_output(layout_pdf, bucket, stem)
        else:
            resp.status = "error"
            resp.error = "未生成排版 PDF (检测或排版失败)"

        # 预览图作为附加结果一并上传
        if layout_preview and os.path.exists(layout_preview):
            p_obj, p_url, p_console = _upload_output(layout_preview, bucket, f"{stem}_preview")
            summary["preview"] = {"object_name": p_obj, "output_url": p_url, "console_url": p_console}

        resp.elapsed_ms = int((time.time() - t0) * 1000)
        return resp
    except Exception as e:
        print(f"[错误] 请求处理失败: {type(e).__name__}: {e}", flush=True)
        traceback.print_exc()
        return ServiceResponse(
            status="error", bucket=bucket, error=f"{type(e).__name__}: {e}",
            elapsed_ms=int((time.time() - t0) * 1000),
        )
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


@app.post("/excel2pdf", response_model=ServiceResponse)
def excel2pdf(req: Excel2PdfRequest):
    """Excel -> PDF (免费离线版: Microsoft Print to PDF; 多 sheet 合并为一页 A4)。

    后端: Excel转PDF_免费离线版_20260820/Excel转PDF_免费离线.exe
    需本机已安装 Microsoft Excel, 并启用 Windows 自带“Microsoft Print to PDF”。
    """
    t0 = time.time()
    bucket = req.bucket or minio_helper.DEFAULT_BUCKET
    workdir = _new_workdir()
    in_name = _name_from_url(req.excel_url, default="input.xlsx")
    stem = os.path.splitext(in_name)[0]
    try:
        input_xlsx = os.path.join(workdir, in_name)
        minio_helper.download_url(req.excel_url, input_xlsx)

        output_dir = os.path.join(workdir, "output")
        out_pdf = _run_serial(
            excel_to_pdf_latest,
            input_excel_path=input_xlsx,
            output_folder=output_dir,
            include_hidden=req.include_hidden,
        )

        resp = ServiceResponse(status="success", bucket=bucket)
        if out_pdf and os.path.exists(out_pdf):
            resp.object_name, resp.output_url, resp.console_url = _upload_output(out_pdf, bucket, stem)
            resp.result = {
                "success": True,
                "source_name": in_name,
                "engine": "Excel转PDF_免费离线版_20260820",
                "include_hidden": req.include_hidden,
            }
        else:
            resp.status = "error"
            resp.error = "Excel 转 PDF 失败 (未生成输出; 请确认已装 Excel 且启用 Microsoft Print to PDF)"
            resp.result = {"success": False, "engine": "Excel转PDF_免费离线版_20260820"}
        resp.elapsed_ms = int((time.time() - t0) * 1000)
        return resp
    except Exception as e:
        print(f"[错误] 请求处理失败: {type(e).__name__}: {e}", flush=True)
        traceback.print_exc()
        return ServiceResponse(
            status="error", bucket=bucket, error=f"{type(e).__name__}: {e}",
            elapsed_ms=int((time.time() - t0) * 1000),
        )
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


@app.post("/divide_pdf", response_model=ServiceResponse)
def divide_pdf_endpoint(req: DividePdfRequest):
    """按首页视觉区块切分 PDF，并返回全部分片的 MinIO 下载地址。"""
    t0 = time.time()
    bucket = req.bucket or minio_helper.DEFAULT_BUCKET
    workdir = _new_workdir()
    in_name = _name_from_url(req.pdf_url)
    stem = os.path.splitext(in_name)[0]
    try:
        input_pdf = os.path.join(workdir, in_name)
        minio_helper.download_url(req.pdf_url, input_pdf)

        output_dir = os.path.join(workdir, "divide_output")
        generated_files = _run_serial(
            divide_pdf,
            input_pdf_path=input_pdf,
            debug_output_dir=output_dir,
            zoom_factor=req.zoom_factor,
            grid_width_ratio=req.grid_width_ratio,
            grid_height_px=req.grid_height_px,
            trim_whitespace=req.trim_whitespace,
            trim_margin_pt=req.trim_margin_pt,
        ) or []

        outputs = []
        for index, local_path in enumerate(generated_files, start=1):
            if not os.path.isfile(local_path):
                continue
            object_name, output_url, console_url = _upload_output(
                local_path, bucket, f"{stem}_part_{index}"
            )
            outputs.append({
                "part": index,
                "object_name": object_name,
                "output_url": output_url,
                "console_url": console_url,
            })

        if not outputs:
            return ServiceResponse(
                status="error",
                bucket=bucket,
                result={"source_name": in_name, "outputs": []},
                elapsed_ms=int((time.time() - t0) * 1000),
                error="PDF 切分未生成有效分片",
            )

        first = outputs[0]
        return ServiceResponse(
            status="success",
            bucket=bucket,
            object_name=first["object_name"],
            output_url=first["output_url"],
            console_url=first["console_url"],
            result={
                "source_name": in_name,
                "part_count": len(outputs),
                "outputs": outputs,
            },
            elapsed_ms=int((time.time() - t0) * 1000),
        )
    except Exception as e:
        print(f"[错误] PDF 切分失败: {type(e).__name__}: {e}", flush=True)
        traceback.print_exc()
        return ServiceResponse(
            status="error", bucket=bucket, error=f"{type(e).__name__}: {e}",
            elapsed_ms=int((time.time() - t0) * 1000),
        )
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


@app.post("/ai_layout", response_model=ServiceResponse)
def ai_layout(req: AiLayoutRequest):
    """给定检测框直接排版。

    - excel=false (默认): ai_toolkit_modifications.apply_layout_pure_python (贴刀版)
    - excel=true: manual_box_layout.apply_layout_manual_boxes (严格按检测框, 适合 Excel 转 PDF 源稿)

    与 /pipeline 的区别: 本接口不做检测, detections 由调用方提供
    (可先调 /detect 拿到 result.final_boxes 再传进来)。
    """
    t0 = time.time()
    bucket = req.bucket or minio_helper.DEFAULT_BUCKET
    workdir = _new_workdir()
    in_name = _name_from_url(req.pdf_url)
    stem = os.path.splitext(in_name)[0]
    try:
        input_pdf = os.path.join(workdir, in_name)
        minio_helper.download_url(req.pdf_url, input_pdf)

        if not req.detections:
            return ServiceResponse(
                status="error", bucket=bucket, error="detections 为空, 无法排版",
                elapsed_ms=int((time.time() - t0) * 1000),
            )

        layout_instruction = req.layout_instruction or DEFAULT_LAYOUT_INSTRUCTION
        box_size_mm = req.box_size_mm or DEFAULT_BOX_SIZE_MM
        layout_out = os.path.join(workdir, "result")
        os.makedirs(layout_out, exist_ok=True)

        if req.excel:
            print(f"[ai_layout] excel=true -> manual_box_layout, dets={len(req.detections)}", flush=True)
            layout_pdf, layout_preview = _run_serial(
                apply_layout_manual_boxes,
                input_pdf,
                layout_instruction,
                box_size_mm,
                req.detections,
                zoom_factor=2.0,
                margin_mm=req.margin_mm,
                out_dir=layout_out,
            )
            engine = "manual_box_layout"
        else:
            print(f"[ai_layout] excel=false -> ai_toolkit_modifications, dets={len(req.detections)}", flush=True)
            import ai_toolkit_modifications as _ai_mod
            _ai_mod.output_dir = layout_out
            layout_pdf, layout_preview = _run_serial(
                ai_apply_layout,
                input_pdf,
                layout_instruction,
                box_size_mm,
                req.detections,
                zoom_factor=2.0,
                resize=True,
                margin_mm=req.margin_mm,
            )
            engine = "ai_toolkit_modifications"

        summary: Dict[str, Any] = {
            "detections_count": len(req.detections),
            "excel": req.excel,
            "engine": engine,
        }
        resp = ServiceResponse(status="success", result=summary, bucket=bucket)
        if layout_pdf and os.path.exists(layout_pdf):
            resp.object_name, resp.output_url, resp.console_url = _upload_output(layout_pdf, bucket, stem)
        else:
            resp.status = "error"
            resp.error = "未生成排版 PDF"
            print(f"[错误] ai_layout 未生成排版 PDF: layout_pdf={layout_pdf!r}, "
                  f"exists={bool(layout_pdf) and os.path.exists(layout_pdf)}, engine={engine}", flush=True)

        if layout_preview and os.path.exists(layout_preview):
            p_obj, p_url, p_console = _upload_output(layout_preview, bucket, f"{stem}_preview")
            summary["preview"] = {"object_name": p_obj, "output_url": p_url, "console_url": p_console}

        resp.elapsed_ms = int((time.time() - t0) * 1000)
        return resp
    except Exception as e:
        print(f"[错误] 请求处理失败: {type(e).__name__}: {e}", flush=True)
        traceback.print_exc()
        return ServiceResponse(
            status="error", bucket=bucket, error=f"{type(e).__name__}: {e}",
            elapsed_ms=int((time.time() - t0) * 1000),
        )
    finally:
        shutil.rmtree(workdir, ignore_errors=True)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "api_server:app",
        host=os.environ.get("API_HOST", "0.0.0.0"),
        port=int(os.environ.get("API_PORT", "8003")),
        reload=False,
    )
