# -*- coding: utf-8 -*-
"""财务报表整理网页。本机处理报表，公网地址写在 web/public_url.txt。"""

from __future__ import annotations

import base64
import io
import socket
import tempfile
import threading
import time
import zipfile
from collections import defaultdict
from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from generate_financial_export import DEFAULT_TEMPLATE, UNIT_NAME, Exporter

ROOT = Path(__file__).resolve().parent
WEB = ROOT / "web"
HOST = "0.0.0.0"
PORT = int(__import__("os").environ.get("PORT", "8765"))
MAX_BYTES = 25 * 1024 * 1024
SOURCE_SUFFIXES = {".xls", ".xlsx", ".et"}
CLEAN_TEMPLATE = WEB / "template.xlsx"

app = FastAPI(title="财务报表整理", docs_url=None, redoc_url=None, openapi_url=None)
_convert_lock = threading.Lock()
_hits_lock = threading.Lock()
_hits: dict[str, list[float]] = defaultdict(list)
RATE_LIMIT = 6
RATE_WINDOW = 60.0
UNZIP_LIMIT = 80 * 1024 * 1024


def lan_addresses() -> list[str]:
    found = set()
    try:
        probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        probe.connect(("192.168.1.1", 80))
        found.add(probe.getsockname()[0])
        probe.close()
    except OSError:
        pass
    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
            ip = info[4][0]
            if not ip.startswith("127."):
                found.add(ip)
    except OSError:
        pass
    return sorted(found)


def _suffix(name: str) -> str:
    return Path(name or "").suffix.lower()


def _safe_name(name: str, fallback: str) -> str:
    suffix = _suffix(name)
    if suffix not in SOURCE_SUFFIXES and suffix != ".xlsx":
        suffix = Path(fallback).suffix.lower() or ".xlsx"
    return "upload" + suffix


def _client_ip(request: Request) -> str:
    if request.client is None:
        return "unknown"
    return request.client.host or "unknown"


def _too_fast(ip: str) -> bool:
    now = time.monotonic()
    with _hits_lock:
        recent = [stamp for stamp in _hits[ip] if now - stamp < RATE_WINDOW]
        if len(recent) >= RATE_LIMIT:
            _hits[ip] = recent
            return True
        recent.append(now)
        _hits[ip] = recent
        return False


def _xlsx_explodes(data: bytes) -> bool:
    if not data.startswith(b"PK"):
        return False
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as book:
            names = book.infolist()
            if len(names) > 400:
                return True
            return sum(item.file_size for item in names) > UNZIP_LIMIT
    except zipfile.BadZipFile:
        return True


@app.get("/")
def index():
    return FileResponse(WEB / "index.html")


if (WEB / "pyodide").is_dir():
    app.mount("/pyodide", StaticFiles(directory=WEB / "pyodide"), name="pyodide")
if (WEB / "vendor").is_dir():
    app.mount("/vendor", StaticFiles(directory=WEB / "vendor"), name="vendor")


def public_url() -> str:
    path = WEB / "public_url.txt"
    if not path.exists():
        return ""
    text = path.read_text(encoding="utf-8").strip()
    return text if text.startswith("http") else ""


@app.get("/api/info")
def info():
    return {"public_url": public_url()}


@app.post("/api/convert")
async def convert(
    request: Request,
    source: UploadFile = File(...),
    template: UploadFile | None = File(None),
    unit: str = Form("yi"),
    periods: int = Form(4),
):
    if _too_fast(_client_ip(request)):
        return JSONResponse({"detail": "操作太频繁，请稍后再试。"}, status_code=429)
    if not _convert_lock.acquire(blocking=False):
        return JSONResponse({"detail": "正在处理另一份报表，请稍后再试。"}, status_code=429)
    try:
        return await _convert(source, template, unit, periods)
    finally:
        _convert_lock.release()


async def _convert(source, template, unit, periods):
    if unit not in UNIT_NAME:
        return JSONResponse({"detail": "输出单位参数错误，仅支持wan(万元) / yi(亿元)。"}, status_code=400)
    if periods < 1 or periods > 20:
        return JSONResponse({"detail": "期数请填写 1 到 20 的整数。"}, status_code=400)
    if _suffix(source.filename or "") not in SOURCE_SUFFIXES:
        return JSONResponse({"detail": "报表请上传 xls、xlsx 或 et 文件。"}, status_code=400)

    source_bytes = await source.read()
    if not source_bytes:
        return JSONResponse({"detail": "上传的报表是空文件。"}, status_code=400)
    if len(source_bytes) > MAX_BYTES:
        return JSONResponse({"detail": "文件超过 25MB，请拆分后再上传。"}, status_code=400)
    if not source_bytes.startswith(b"PK") and not source_bytes.startswith(b"\xd0\xcf\x11\xe0"):
        return JSONResponse({"detail": "无法识别报表格式。请上传系统导出的 Excel 或 WPS 表格。"}, status_code=400)
    if _xlsx_explodes(source_bytes):
        return JSONResponse({"detail": "报表解压后过大，请检查文件后再上传。"}, status_code=400)

    template_bytes = None
    if template is not None and template.filename:
        if _suffix(template.filename) != ".xlsx":
            return JSONResponse({"detail": "自定义模板只支持真正的 xlsx 文件。"}, status_code=400)
        template_bytes = await template.read()
        if not template_bytes.startswith(b"PK"):
            return JSONResponse({"detail": "自定义模板不是 xlsx 格式。"}, status_code=400)
        if _xlsx_explodes(template_bytes):
            return JSONResponse({"detail": "自定义模板解压后过大。"}, status_code=400)
    elif not (CLEAN_TEMPLATE.exists() or DEFAULT_TEMPLATE.exists()):
        return JSONResponse({"detail": "服务器上缺少默认导出模板。"}, status_code=500)

    try:
        exporter = Exporter(
            output_unit=unit,
            include_rd=False,
            ebitda_mode="std",
            period_count=int(periods),
        )
    except ValueError as exc:
        return JSONResponse({"detail": str(exc)}, status_code=400)

    with tempfile.TemporaryDirectory(prefix="fin_") as folder:
        work = Path(folder)
        source_path = work / _safe_name(source.filename or "", "source.xlsx")
        source_path.write_bytes(source_bytes)
        if template_bytes is None:
            template_path = CLEAN_TEMPLATE if CLEAN_TEMPLATE.exists() else DEFAULT_TEMPLATE
        else:
            template_path = work / "template.xlsx"
            template_path.write_bytes(template_bytes)
        output_path = work / "资产类财务数据_导出.xlsx"
        try:
            exporter.run(source_path, template_path, output_path)
        except ValueError as exc:
            return JSONResponse({"detail": str(exc)}, status_code=400)
        except Exception:
            return JSONResponse({"detail": "整理失败，请确认上传的是系统导出的财务报表。"}, status_code=400)
        payload = output_path.read_bytes()

    logs = list(exporter.warnings)
    logs.append(f"已选择输出单位：{exporter.unit_name}，全部金额已完成换算，比率未缩放。")
    return {
        "filename": "资产类财务数据_导出.xlsx",
        "logs": logs,
        "content_base64": base64.b64encode(payload).decode("ascii"),
    }


if __name__ == "__main__":
    import uvicorn

    local = f"http://127.0.0.1:{PORT}"
    opened = public_url() or "见 web/public_url.txt（公网通道启动后写入）"
    print("财务报表整理已启动")
    print("本机打开：", local)
    print("公网打开：", opened)
    print("电脑需保持开机并联网。对方不必和这台电脑在同一网络。")
    uvicorn.run(app, host=HOST, port=PORT, log_level="info")
