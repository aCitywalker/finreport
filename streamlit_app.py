# -*- coding: utf-8 -*-
"""Streamlit 云端入口。页面沿用原来的整理界面，计算在 Streamlit 上完成。"""
from __future__ import annotations

import base64
import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from generate_financial_export import Exporter

ROOT = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = ROOT / "web" / "template.xlsx"
SOURCE_SUFFIXES = {".xls", ".xlsx", ".et"}
MAX_BYTES = 25 * 1024 * 1024
UNZIP_LIMIT = 80 * 1024 * 1024
_ui = components.declare_component("fin_ui", path=str(ROOT / "web" / "component"))

st.set_page_config(page_title="财务报表整理", page_icon="📊", layout="centered")
st.markdown(
    """
<style>
#MainMenu, footer, header, [data-testid="stHeader"], [data-testid="stToolbar"] { visibility: hidden; height: 0; }
.block-container { padding: 0; max-width: 100%; }
.stApp { background: #f4f0e8; }
iframe { border: 0; }
</style>
""",
    unsafe_allow_html=True,
)


def _suffix(name: str) -> str:
    return Path(name or "").suffix.lower()


def _decode(text: str) -> bytes:
    return base64.b64decode(text or "")


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


def _check_source(name: str, data: bytes) -> str | None:
    if _suffix(name) not in SOURCE_SUFFIXES:
        return "报表请上传 xls、xlsx 或 et 文件。"
    if not data:
        return "上传的报表是空文件。"
    if len(data) > MAX_BYTES:
        return "文件超过 25MB，请拆分后再上传。"
    if not data.startswith(b"PK") and not data.startswith(b"\xd0\xcf\x11\xe0"):
        return "无法识别报表格式。请上传系统导出的 Excel 或 WPS 表格。"
    if _xlsx_explodes(data):
        return "报表解压后过大，请检查文件后再上传。"
    return None


def _convert(source_name: str, source_bytes: bytes, template_bytes: bytes | None, unit: str, periods: int):
    with tempfile.TemporaryDirectory(prefix="fin_") as folder:
        work = Path(folder)
        source_path = work / ("upload" + (_suffix(source_name) or ".xlsx"))
        source_path.write_bytes(source_bytes)
        if template_bytes is None:
            template_path = DEFAULT_TEMPLATE
        else:
            template_path = work / "template.xlsx"
            template_path.write_bytes(template_bytes)
        output_path = work / "资产类财务数据_导出.xlsx"
        exporter = Exporter(
            output_unit=unit,
            include_rd=False,
            ebitda_mode="std",
            period_count=periods,
        )
        exporter.run(source_path, template_path, output_path)
        payload = output_path.read_bytes()
    logs = list(exporter.warnings)
    logs.append(f"已选择输出单位：{exporter.unit_name}，全部金额已完成换算，比率未缩放。")
    return payload, logs


def _fail(request_id: str, detail: str) -> dict:
    return {"id": request_id, "ok": False, "detail": detail}


event = _ui(result=st.session_state.get("reply"), default=None, key="fin_ui")
if isinstance(event, dict) and event.get("kind") == "convert":
    request_id = str(event.get("id") or "")
    if request_id and request_id != st.session_state.get("done_id"):
        st.session_state["done_id"] = request_id
        try:
            source_bytes = _decode(str(event.get("source_b64") or ""))
            problem = _check_source(str(event.get("source_name") or ""), source_bytes)
            template_bytes = None
            if problem is None and event.get("template_b64"):
                template_bytes = _decode(str(event.get("template_b64")))
                if _suffix(str(event.get("template_name") or "")) != ".xlsx" or not template_bytes.startswith(b"PK"):
                    problem = "自定义模板不是 xlsx 格式。"
                elif _xlsx_explodes(template_bytes):
                    problem = "自定义模板解压后过大。"
            if problem is None and not DEFAULT_TEMPLATE.exists() and template_bytes is None:
                problem = "缺少默认导出模板。"
            periods = int(event.get("periods") or 4)
            unit = str(event.get("unit") or "yi")
            if problem:
                st.session_state["reply"] = _fail(request_id, problem)
            else:
                payload, logs = _convert(
                    str(event.get("source_name") or "source.xlsx"),
                    source_bytes,
                    template_bytes,
                    unit,
                    periods,
                )
                st.session_state["reply"] = {
                    "id": request_id,
                    "ok": True,
                    "filename": "资产类财务数据_导出.xlsx",
                    "logs": logs,
                    "content_base64": base64.b64encode(payload).decode("ascii"),
                }
        except ValueError as exc:
            st.session_state["reply"] = _fail(request_id, str(exc))
        except Exception:
            st.session_state["reply"] = _fail(request_id, "整理失败，请确认上传的是系统导出的财务报表。")
        st.rerun()
