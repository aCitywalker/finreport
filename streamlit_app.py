# -*- coding: utf-8 -*-
"""Streamlit 云端入口。计算在 Streamlit 云上完成，不经过现有服务器。"""
from __future__ import annotations

import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from generate_financial_export import UNIT_NAME, Exporter

ROOT = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = ROOT / "web" / "template.xlsx"
SOURCE_SUFFIXES = {".xls", ".xlsx", ".et"}
MAX_BYTES = 25 * 1024 * 1024
UNZIP_LIMIT = 80 * 1024 * 1024

st.set_page_config(page_title="财务报表整理", page_icon="📊", layout="centered")
st.markdown(
    """
<style>
.block-container { padding-top: 1.2rem; max-width: 42rem; }
div.stButton > button { width: 100%; height: 3rem; }
div.stDownloadButton > button { width: 100%; height: 3rem; }
</style>
""",
    unsafe_allow_html=True,
)


def _suffix(name: str) -> str:
    return Path(name or "").suffix.lower()


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


st.title("财务报表整理")
st.caption("上传系统导入版三表，生成资产、损益、现金流和指标简表。金额会换算，比率和百分比保持原口径。")

source = st.file_uploader("财务报表", type=["xls", "xlsx", "et"])
unit_label = st.radio("输出单位", ["亿元", "万元"], horizontal=True)
unit = "yi" if unit_label == "亿元" else "wan"
periods = st.number_input("取最近几期", min_value=1, max_value=20, value=4, step=1)
template = st.file_uploader("自定义模板（可选，仅 xlsx）", type=["xlsx"])

if st.button("生成并下载", type="primary"):
    if source is None:
        st.error("请先选择财务报表。")
    elif not DEFAULT_TEMPLATE.exists() and template is None:
        st.error("服务器上缺少默认导出模板。")
    else:
        source_bytes = source.getvalue()
        problem = _check_source(source.name, source_bytes)
        template_bytes = None
        if problem is None and template is not None:
            template_bytes = template.getvalue()
            if not template_bytes.startswith(b"PK"):
                problem = "自定义模板不是 xlsx 格式。"
            elif _xlsx_explodes(template_bytes):
                problem = "自定义模板解压后过大。"
        if problem:
            st.error(problem)
        else:
            try:
                with st.spinner("正在整理，请稍等…"):
                    payload, logs = _convert(source.name, source_bytes, template_bytes, unit, int(periods))
                st.session_state["payload"] = payload
                st.session_state["logs"] = logs
            except ValueError as exc:
                st.error(str(exc))
            except Exception:
                st.error("整理失败，请确认上传的是系统导出的财务报表。")

logs = st.session_state.get("logs") or []
if logs:
    warned = any(line.startswith("【") for line in logs)
    if warned:
        st.warning("已生成，请看提示")
    else:
        st.success("已生成")
    for line in logs:
        if line.startswith("【勾稽告警】") or line.startswith("【严重告警】"):
            st.warning(line)
        elif line.startswith("【未映射科目】") or line.startswith("【未识别模板行】"):
            st.error(line)
        else:
            st.caption(line)

payload = st.session_state.get("payload")
if payload:
    st.download_button(
        f"下载导出文件（{UNIT_NAME[unit]}）",
        data=payload,
        file_name="资产类财务数据_导出.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
