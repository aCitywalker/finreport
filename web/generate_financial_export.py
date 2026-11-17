# -*- coding: utf-8 -*-
"""从表1（系统导入版财务报表）生成表2（资产类财务数据导出表）。

调用参数只在代码内使用，不写入 Excel。
    OUTPUT_UNIT:          "wan" 万元 / "yi" 亿元
    INCLUDE_RD_IN_3FEE:   True 三费含研发费用；False 不含
    EBITDA_MODE:          "std" 利润总额+利息费用+折旧+摊销
                          "legacy" 用净利润代替 EBITDA
    PERIOD_COUNT:         输出最近 N 期
"""

from __future__ import annotations

import argparse
import re
from copy import copy
from datetime import datetime, date
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

import xlrd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, PatternFill

# 外部可配置入参（也可被命令行覆盖）
OUTPUT_UNIT = "yi"
INCLUDE_RD_IN_3FEE = False
EBITDA_MODE = "std"
PERIOD_COUNT = 4

ROOT = Path(__file__).resolve().parent
DEFAULT_INPUT = ROOT / "1.财务报表（系统导入版）.xlsx"
DEFAULT_TEMPLATE = ROOT / "2.资产类财务数据2023-202606（优化）.xlsx"
DEFAULT_OUTPUT = ROOT / "资产类财务数据_导出.xlsx"

UNIT_NAME = {"wan": "万元", "yi": "亿元"}
DIV_SCALE = {"wan": Decimal("10000"), "yi": Decimal("100000000")}
TOLERANCE = Decimal("0.02")

# 简表明细：按源表顺序。别名按优先使用的科目排列。
# role: detail / subtotal / total ；base: asset / liability / equity
BS_LINES = [
    ("header", "资产类", [], None),
    ("detail", "货币资金", ["货币资金"], "asset"),
    ("detail", "结算备付金", ["结算备付金"], "asset"),
    ("detail", "拆出资金", ["拆出资金"], "asset"),
    ("detail", "交易性金融资产", ["交易性金融资产"], "asset"),
    ("detail", "以公允价值计量且其变动计入当期损益的金融资产", ["以公允价值计量且其变动计入当期损益的金融资产"], "asset"),
    ("detail", "衍生金融资产", ["衍生金融资产"], "asset"),
    ("detail", "应收票据", ["应收票据"], "asset"),
    ("detail", "应收账款", ["应收账款"], "asset"),
    ("detail", "应收款项融资", ["应收款项融资"], "asset"),
    ("detail", "预付款项", ["预付款项"], "asset"),
    ("detail", "应收保费", ["应收保费"], "asset"),
    ("detail", "应收分保账款", ["应收分保账款"], "asset"),
    ("detail", "应收分保合同准备金", ["应收分保合同准备金"], "asset"),
    ("detail", "应收利息", ["应收利息"], "asset"),
    ("detail", "应收股利", ["应收股利"], "asset"),
    ("detail", "其他应收款", ["其他应收款", "其他应收款项"], "asset"),
    ("detail", "买入返售金融资产", ["买入返售金融资产"], "asset"),
    ("detail", "存货", ["存货"], "asset"),
    ("detail", "合同资产", ["合同资产"], "asset"),
    ("detail", "划分为持有待售资产", ["划分为持有待售资产"], "asset"),
    ("detail", "一年内到期的非流动资产", ["一年内到期的非流动资产"], "asset"),
    ("detail", "其他流动资产", ["其他流动资产"], "asset"),
    ("subtotal", "流动资产小计", ["流动资产合计"], "asset"),
    ("detail", "发放贷款及垫款", ["发放贷款及垫款"], "asset"),
    ("detail", "可供出售金融资产", ["可供出售金融资产"], "asset"),
    ("detail", "持有至到期投资", ["持有至到期投资"], "asset"),
    ("detail", "债权投资", ["债权投资"], "asset"),
    ("detail", "以摊余成本计量的金融资产", ["以摊余成本计量的金融资产"], "asset"),
    ("detail", "其他债权投资", ["其他债权投资"], "asset"),
    ("detail", "长期应收款", ["长期应收款"], "asset"),
    ("detail", "长期股权投资", ["长期股权投资"], "asset"),
    ("detail", "其他权益工具投资", ["其他权益工具投资"], "asset"),
    ("detail", "其他非流动金融资产", ["其他非流动金融资产"], "asset"),
    ("detail", "投资性房地产", ["投资性房地产"], "asset"),
    ("detail", "固定资产", ["固定资产"], "asset"),
    ("detail", "固定资产清理", ["固定资产清理"], "asset"),
    ("detail", "在建工程", ["在建工程", "在建工程合计"], "asset"),
    ("detail", "工程物资", ["工程物资"], "asset"),
    ("detail", "生产性生物资产", ["生产性生物资产"], "asset"),
    ("detail", "油气资产", ["油气资产"], "asset"),
    ("detail", "使用权资产", ["使用权资产"], "asset"),
    ("detail", "无形资产", ["无形资产"], "asset"),
    ("detail", "开发支出", ["开发支出"], "asset"),
    ("detail", "商誉", ["商誉"], "asset"),
    ("detail", "长期待摊费用", ["长期待摊费用"], "asset"),
    ("detail", "递延所得税资产", ["递延所得税资产"], "asset"),
    ("detail", "其他非流动资产", ["其他非流动资产"], "asset"),
    ("subtotal", "非流动资产小计", ["非流动资产合计"], "asset"),
    ("total", "总资产合计", ["资产总计"], "asset"),
    ("header", "负债类", [], None),
    ("detail", "短期借款", ["短期借款"], "liability"),
    ("detail", "向中央银行借款", ["向中央银行借款"], "liability"),
    ("detail", "拆入资金", ["拆入资金"], "liability"),
    ("detail", "交易性金融负债", ["交易性金融负债"], "liability"),
    ("detail", "以公允价值计量且其变动计入当期损益的金融负债", ["以公允价值计量且其变动计入当期损益的金融负债"], "liability"),
    ("detail", "衍生金融负债", ["衍生金融负债"], "liability"),
    ("detail", "应付票据", ["应付票据"], "liability"),
    ("detail", "应付账款", ["应付账款"], "liability"),
    ("detail", "预收账款", ["预收账款"], "liability"),
    ("detail", "合同负债", ["合同负债"], "liability"),
    ("detail", "卖出回购金融资产款", ["卖出回购金融资产款"], "liability"),
    ("detail", "吸收存款及同业存放", ["吸收存款及同业存放"], "liability"),
    ("detail", "代理买卖证券款", ["代理买卖证券款"], "liability"),
    ("detail", "代理承销证券款", ["代理承销证券款"], "liability"),
    ("detail", "应付职工薪酬", ["应付职工薪酬"], "liability"),
    ("detail", "应交税费", ["应交税费"], "liability"),
    ("detail", "应付利息", ["应付利息"], "liability"),
    ("detail", "应付股利", ["应付股利"], "liability"),
    ("detail", "其他应付款", ["其他应付款", "其他应付款项"], "liability"),
    ("detail", "应付手续费及佣金", ["应付手续费及佣金"], "liability"),
    ("detail", "应付分保账款", ["应付分保账款"], "liability"),
    ("detail", "应付短期债券", ["应付短期债券"], "liability"),
    ("detail", "持有待售负债", ["持有待售负债"], "liability"),
    ("detail", "一年内到期的非流动负债", ["一年内到期的非流动负债"], "liability"),
    ("detail", "其他流动负债", ["其他流动负债"], "liability"),
    ("subtotal", "流动负债小计", ["流动负债合计"], "liability"),
    ("detail", "保险合同准备金", ["保险合同准备金"], "liability"),
    ("detail", "长期借款", ["长期借款"], "liability"),
    ("detail", "应付债券", ["应付债券"], "liability"),
    ("detail", "租赁负债", ["租赁负债"], "liability"),
    ("detail", "长期应付款", ["长期应付款", "长期应付款项"], "liability"),
    ("detail", "专项应付款", ["专项应付款"], "liability"),
    ("detail", "预计负债", ["预计负债"], "liability"),
    ("detail", "递延收益", ["递延收益"], "liability"),
    ("detail", "递延所得税负债", ["递延所得税负债"], "liability"),
    ("detail", "其他非流动负债", ["其他非流动负债"], "liability"),
    ("subtotal", "非流动负债小计", ["非流动负债合计"], "liability"),
    ("total", "总负债小计", ["负债合计"], "liability"),
    ("header", "权益类", [], None),
    ("detail", "实收资本（股本）", ["实收资本（或股本）", "实收资本(或股本)", "实收资本"], "equity"),
    ("detail", "其他权益工具", ["其他权益工具"], "equity"),
    ("detail", "资本公积", ["资本公积"], "equity"),
    ("detail", "减：库存股", ["减:库存股", "减：库存股"], "equity"),
    ("detail", "其他综合收益", ["其他综合收益"], "equity"),
    ("detail", "专项储备", ["专项储备"], "equity"),
    ("detail", "盈余公积", ["盈余公积"], "equity"),
    ("detail", "一般风险准备", ["一般风险准备"], "equity"),
    ("detail", "未分配利润", ["未分配利润"], "equity"),
    ("detail", "外币报表折算差额", ["外币报表折算差额"], "equity"),
    ("detail", "少数股东权益", ["少数股东权益"], "equity"),
    ("total", "所有者权益小计", ["所有者权益（或股东权益）合计", "所有者权益合计"], "equity"),
]

# 已并入净值或其他展示行、不单独出简表、也不算“未映射”的结构行
ABSORBED_NAMES = {
    "固定资产及清理合计",
    "固定资产原值",
    "减:累计折旧",
    "减：累计折旧",
    "累计折旧",
    "减:固定资产减值准备",
    "减：固定资产减值准备",
    "固定资产减值准备",
    "在建工程合计",
    "其他应收款项",
    "其他应付款项",
    "长期应付款项",
    "归属于母公司股东权益合计",
    "归属于母公司所有者权益合计",
    "归属于母公司所有者的净利润",
    "少数股东损益",
    "负债和所有者权益（或股东权益）总计",
    "负债和所有者权益总计",
}


def fold_label(value) -> str:
    return norm_label(value).casefold()


def norm_label(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\u3000", " ").replace("\xa0", " ").strip()
    text = text.replace("(", "（").replace(")", "）")
    text = text.replace("：", ":")
    text = re.sub(r"\s+", "", text)
    return text.rstrip("*＊")


def core_name(value) -> str:
    text = norm_label(value)
    text = re.sub(r"^[一二三四五六七八九十]+、", "", text)
    text = re.sub(r"^加[:：]", "", text)
    text = re.sub(r"^减[:：]", "", text)
    text = re.sub(r"（损失以.*?）", "", text)
    text = re.sub(r"（亏损总额以.*?）", "", text)
    text = re.sub(r"（净亏损以.*?）", "", text)
    text = re.sub(r"（收益以.*?）", "", text)
    return text


def round_half_up(number, digits: int) -> Decimal:
    quant = Decimal("1").scaleb(-digits)
    if not isinstance(number, Decimal):
        number = Decimal(format(number, ".12f"))
    return number.quantize(quant, rounding=ROUND_HALF_UP)


def to_decimal(value) -> Decimal | None:
    if value is None:
        return None
    return Decimal(format(value, ".12f"))


def meaningful(value: Decimal | None) -> bool:
    return value is not None and abs(value) >= Decimal("0.005")


def parse_period(value, datemode=0):
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if value > 20000:
            return xlrd.xldate_as_datetime(value, datemode).date()
        return None
    text = str(value).strip()
    head = text[:10]
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(head, fmt).date()
        except ValueError:
            continue
    return None


def period_label(day: date) -> str:
    if day.month == 12:
        return f"{day.year}年"
    if day.month == 6 and day.day >= 28:
        return f"{day.year}年1—6月"
    if day.month == 3 and day.day >= 28:
        return f"{day.year}年1—3月"
    if day.month == 9 and day.day >= 28:
        return f"{day.year}年1—9月"
    return f"{day.year}年{day.month}月"


def is_annual(day: date) -> bool:
    return day.month == 12


class Line:
    def __init__(self, sheet: str, row: int, name: str, values: list[Decimal | None]):
        self.sheet = sheet
        self.row = row
        self.name = name
        self.norm = norm_label(name)
        self.core = core_name(name)
        self.values = values

    def at(self, index: int) -> Decimal | None:
        if index < 0 or index >= len(self.values):
            return None
        return self.values[index]


class Statement:
    def __init__(self, key: str, name: str, periods: list[date], lines: list[Line]):
        self.key = key
        self.name = name
        self.periods = periods
        self.lines = lines
        self.index = {day: i for i, day in enumerate(periods)}


def _cell_number_xlrd(cell) -> Decimal | None:
    if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return None
    if cell.ctype == xlrd.XL_CELL_TEXT:
        text = str(cell.value).strip()
        if text == "":
            return None
        try:
            return Decimal(text.replace(",", ""))
        except Exception:
            return None
    if cell.ctype == xlrd.XL_CELL_NUMBER:
        return to_decimal(cell.value)
    return None


def _read_sheet_xlrd(book, sheet, key: str) -> Statement:
    periods = []
    for col in range(1, sheet.ncols):
        periods.append(parse_period(sheet.cell_value(0, col), book.datemode))
    lines = []
    for row in range(1, sheet.nrows):
        name = sheet.cell_value(row, 0)
        if name is None or str(name).strip() == "":
            continue
        values = [_cell_number_xlrd(sheet.cell(row, col)) for col in range(1, sheet.ncols)]
        lines.append(Line(key, row, str(name).strip(), values))
    return Statement(key, sheet.name, periods, lines)


def _cell_number_openpyxl(value) -> Decimal | None:
    if value is None or value == "":
        return None
    if isinstance(value, str):
        text = value.strip()
        if text == "":
            return None
        try:
            return Decimal(text.replace(",", ""))
        except Exception:
            return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return to_decimal(value)
    return None


def _read_sheet_openpyxl(ws, key: str) -> Statement:
    periods = []
    for col in range(2, ws.max_column + 1):
        periods.append(parse_period(ws.cell(1, col).value))
    while periods and periods[-1] is None:
        periods.pop()
    lines = []
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row, 1).value
        if name is None or str(name).strip() == "":
            continue
        values = [_cell_number_openpyxl(ws.cell(row, col).value) for col in range(2, len(periods) + 2)]
        lines.append(Line(key, row, str(name).strip(), values))
    return Statement(key, ws.title, periods, lines)


def load_table1(path: Path) -> dict[str, Statement]:
    raw = path.read_bytes()[:8]
    books: dict[str, Statement] = {}
    if raw.startswith(b"PK"):
        wb = load_workbook(path, data_only=True)
        for ws in wb.worksheets:
            key = _sheet_key(ws.title)
            if key:
                books[key] = _read_sheet_openpyxl(ws, key)
        wb.close()
    elif raw.startswith(b"\xd0\xcf\x11\xe0"):
        book = xlrd.open_workbook(path, formatting_info=False)
        for sheet in book.sheets():
            key = _sheet_key(sheet.name)
            if key:
                books[key] = _read_sheet_xlrd(book, sheet, key)
    else:
        raise ValueError(f"无法识别表1文件格式：{path}")
    missing = [name for name in ("bs", "is", "cf") if name not in books]
    if missing:
        raise ValueError("表1缺少报表：" + "、".join(missing))
    return books


def _sheet_key(name: str) -> str | None:
    if "资产负债" in name:
        return "bs"
    if "现金" in name:
        return "cf"
    if "损益" in name or "利润" in name:
        return "is"
    return None


def find_lines(stmt: Statement, *aliases: str) -> list[Line]:
    wanted = {norm_label(a) for a in aliases}
    wanted |= {core_name(a) for a in aliases}
    found = []
    for line in stmt.lines:
        if line.norm in wanted or line.core in wanted:
            found.append(line)
    return found


def resolve_line(stmt: Statement, aliases: list[str], consumed: set, selected_idx: list[int]) -> Line | None:
    grouped: list[list[Line]] = []
    for alias in aliases:
        wanted_norm = norm_label(alias)
        wanted_core = core_name(alias)
        matches = [
            line for line in stmt.lines
            if line.norm == wanted_norm or line.core == wanted_core or line.norm == wanted_core
        ]
        grouped.append(matches)
    chosen = None
    for matches in grouped:
        for line in matches:
            if any(meaningful(line.at(i)) for i in selected_idx):
                chosen = line
                break
        if chosen is not None:
            break
    if chosen is None:
        for matches in grouped:
            if matches:
                chosen = matches[0]
                break
    for matches in grouped:
        for line in matches:
            consumed.add((line.sheet, line.row))
    return chosen


def value_of(stmt: Statement, day: date, aliases: list[str], consumed: set | None = None) -> Decimal | None:
    found = find_lines(stmt, *aliases)
    if consumed is not None:
        for line in found:
            consumed.add((line.sheet, line.row))
    if day not in stmt.index or not found:
        return None
    idx = stmt.index[day]
    for line in found:
        if line.at(idx) is not None:
            return line.at(idx)
    return None


def series_of(stmt: Statement, days: list[date], aliases: list[str], consumed: set) -> list[Decimal | None]:
    find_lines(stmt, *aliases)  # touch
    lines = find_lines(stmt, *aliases)
    for line in lines:
        consumed.add((line.sheet, line.row))
    out = []
    for day in days:
        if day not in stmt.index or not lines:
            out.append(None)
            continue
        idx = stmt.index[day]
        picked = None
        for line in lines:
            if line.at(idx) is not None:
                picked = line.at(idx)
                break
        out.append(picked)
    return out


def sum_series(*parts: list[Decimal | None]) -> list[Decimal | None]:
    length = len(parts[0]) if parts else 0
    out = []
    for i in range(length):
        nums = [part[i] for part in parts if part[i] is not None]
        out.append(sum(nums, Decimal("0")) if nums else None)
    return out


def sub_series(left: list[Decimal | None], right: list[Decimal | None]) -> list[Decimal | None]:
    out = []
    for a, b in zip(left, right):
        if a is None and b is None:
            out.append(None)
        else:
            out.append((a or Decimal("0")) - (b or Decimal("0")))
    return out


def scale_series(values: list[Decimal | None], div: Decimal) -> list[Decimal | None]:
    scaled = []
    for value in values:
        if value is None:
            scaled.append(None)
        else:
            scaled.append(round_half_up(value / div, 2))
    return scaled


def ratio_series(numer: list[Decimal | None], denom: list[Decimal | None], digits: int) -> list:
    out = []
    for a, b in zip(numer, denom):
        if a is None or b is None or b == 0:
            out.append(None)
        else:
            out.append(float(round_half_up(a / b, digits)))
    return out


def previous_day(all_days: list[date], day: date):
    earlier = [item for item in all_days if item < day]
    return earlier[-1] if earlier else None


def snapshot_style(cell) -> dict:
    return {
        "font": copy(cell.font),
        "border": copy(cell.border),
        "fill": copy(cell.fill),
        "alignment": copy(cell.alignment),
        "number_format": cell.number_format,
    }


def apply_style(cell, style: dict) -> None:
    cell.font = copy(style["font"])
    cell.border = copy(style["border"])
    cell.fill = copy(style["fill"])
    cell.alignment = copy(style["alignment"])
    cell.number_format = style["number_format"]


HEADER_FILL = PatternFill(fill_type="solid", fgColor="D7D7D7")
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)


def paint_header(cell) -> None:
    cell.fill = HEADER_FILL
    cell.alignment = HEADER_ALIGN


def clear_after(ws, row: int, last_col: int) -> None:
    """去掉数据列右侧多出来的边框和残留内容，不向右扩出新列。"""
    end = ws.max_column or last_col
    for col in range(last_col + 1, end + 1):
        cell = ws.cell(row, col)
        if isinstance(cell, MergedCell):
            continue
        cell.value = None
        cell.border = Border()


def unmerge_from_row(ws, min_row: int) -> None:
    targets = [str(item) for item in ws.merged_cells.ranges if item.min_row >= min_row]
    for item in targets:
        ws.unmerge_cells(item)


def unmerge_row(ws, row: int) -> None:
    targets = [str(item) for item in ws.merged_cells.ranges if item.min_row == row and item.max_row == row]
    for item in targets:
        ws.unmerge_cells(item)


def find_row(ws, label: str, start: int = 1) -> int | None:
    target = norm_label(label)
    for row in range(start, ws.max_row + 1):
        if norm_label(ws.cell(row, 1).value) == target:
            return row
    return None


def rewrite_unit_text(text, unit_name: str):
    if not isinstance(text, str):
        return text
    if "万元" not in text and "亿元" not in text:
        return text
    text = text.replace("万元", unit_name).replace("亿元", unit_name)
    marker = f"（单位：{unit_name}）"
    text = text.replace(marker, "@@UNIT@@")
    text = text.replace(f"（{unit_name}）", marker)
    return text.replace("@@UNIT@@", marker)


def is_block_header(label, right_value) -> bool:
    if label is None:
        return False
    text = str(label)
    key = norm_label(text)
    if key in {"年度", "科目（亿元）", "科目（万元）"}:
        return True
    if "单位" in text and "占比" not in text and not key.startswith("净利润中财政补贴"):
        return True
    if key.startswith("科目") or key.startswith("项目"):
        return True
    return False


def put_number(cell, value, kind: str) -> None:
    if value is None:
        cell.value = None
        return
    if value == "-":
        cell.value = "-"
        cell.number_format = "General"
        return
    cell.value = float(value) if isinstance(value, Decimal) else value
    current = str(cell.number_format or "")
    if kind == "percent":
        cell.number_format = "0.00%"
    elif kind == "ratio":
        cell.number_format = "0.00" if "%" in current or "0" not in current else current
    else:
        if "%" in current or "0" not in current:
            cell.number_format = "0.00"


class Exporter:
    def __init__(self, output_unit: str, include_rd: bool, ebitda_mode: str, period_count: int):
        if output_unit not in ("wan", "yi"):
            raise ValueError("输出单位参数错误，仅支持wan(万元) / yi(亿元)。")
        if ebitda_mode not in ("std", "legacy"):
            raise ValueError("EBITDA口径参数错误，仅支持std(标准口径) / legacy(兼容旧模板)。")
        if not isinstance(period_count, int) or period_count < 1:
            raise ValueError("PERIOD_COUNT必须为大于0的整数。")
        self.output_unit = output_unit
        self.include_rd = bool(include_rd)
        self.ebitda_mode = ebitda_mode
        self.period_count = period_count
        self.unit_name = UNIT_NAME[output_unit]
        self.div = DIV_SCALE[output_unit]
        self.warnings: list[str] = []

    def warn(self, text: str) -> None:
        self.warnings.append(text)
        print(text)

    def run(self, source: Path, template: Path, output: Path) -> None:
        books = load_table1(source)
        days = self._select_periods(books["bs"])
        consumed: set[tuple[str, int]] = set()
        self._absorb_structural(books, consumed)
        values = self._build_values(books, days, consumed)
        self._log_unmapped(books, consumed)
        self._validate(books, days)
        self._write(template, output, days, values)
        print(f"已选择输出单位：{self.unit_name}，全部金额已完成换算，比率未缩放。")
        print(f"输出文件：{output}")

    def _select_periods(self, bs: Statement) -> list[date]:
        days = [day for day in bs.periods if day is not None]
        if len(days) < len(bs.periods):
            self.warn("【严重告警】资产负债表存在无法识别的期间列，已跳过。")
        if not days:
            raise ValueError("表1资产负债表没有可用期间。")
        if self.period_count > len(days):
            self.warn(f"【严重告警】可用期间只有{len(days)}期，少于PERIOD_COUNT={self.period_count}，已全部输出。")
        return days[-self.period_count :]

    def _absorb_structural(self, books: dict[str, Statement], consumed: set) -> None:
        absorbed = {norm_label(name) for name in ABSORBED_NAMES}
        absorbed |= {core_name(name) for name in ABSORBED_NAMES}
        for stmt in books.values():
            in_supplement = False
            for line in stmt.lines:
                if "补充资料" in line.name or "补充信息" in line.name:
                    in_supplement = True
                if in_supplement or line.norm.startswith("其中") or line.core.startswith("其中"):
                    consumed.add((line.sheet, line.row))
                if line.norm in absorbed or line.core in absorbed or line.norm.startswith("注:"):
                    consumed.add((line.sheet, line.row))

    def _build_values(self, books, days, consumed) -> dict:
        bs, inc, cf = books["bs"], books["is"], books["cf"]
        selected_idx = [bs.index[day] for day in days if day in bs.index]

        def raw(stmt, aliases):
            return series_of(stmt, days, aliases, consumed)

        assets = raw(bs, ["资产总计"])
        current_assets = raw(bs, ["流动资产合计"])
        noncurrent_assets = raw(bs, ["非流动资产合计"])
        liabilities = raw(bs, ["负债合计"])
        current_liab = raw(bs, ["流动负债合计"])
        noncurrent_liab = raw(bs, ["非流动负债合计"])
        equity = raw(bs, ["所有者权益（或股东权益）合计", "所有者权益合计"])
        inventory = raw(bs, ["存货"])
        cash = raw(bs, ["货币资金"])
        ar = raw(bs, ["应收账款"])
        long_equity = raw(bs, ["长期股权投资"])
        short_loan = raw(bs, ["短期借款"])
        long_loan = raw(bs, ["长期借款"])
        due_in_year = raw(bs, ["一年内到期的非流动负债"])
        notes_pay = raw(bs, ["应付票据"])
        short_bond = raw(bs, ["应付短期债券"])
        bonds = raw(bs, ["应付债券"])
        lease = raw(bs, ["租赁负债"])

        revenue = raw(inc, ["其中:营业收入", "营业收入"])
        if all(v is None for v in revenue):
            revenue = raw(inc, ["营业总收入"])
        cost = raw(inc, ["其中:营业成本", "营业成本"])
        sell = raw(inc, ["销售费用"])
        admin = raw(inc, ["管理费用"])
        rnd = raw(inc, ["研发费用"])
        finance = raw(inc, ["财务费用"])
        interest = raw(inc, ["其中:利息费用", "利息费用"])
        impairment = raw(inc, ["资产减值损失（损失以“-”号填列）", "资产减值损失"])
        other_income = raw(inc, ["其他收益"])
        invest_income = raw(inc, ["投资收益（损失以“－”号填列）", "投资收益"])
        operate_profit = raw(inc, ["营业利润"])
        nonop_income = raw(inc, ["营业外收入"])
        nonop_expense = raw(inc, ["营业外支出"])
        total_profit = raw(inc, ["利润总额（亏损总额以“-”号填列）", "利润总额"])
        tax = raw(inc, ["所得税费用"])
        net_profit = raw(inc, ["净利润（净亏损以“-”号填列）", "净利润"])

        fee_parts = [sell, admin, finance]
        if self.include_rd:
            fee_parts.append(rnd)
        three_fees = sum_series(*fee_parts)
        nonop_net = sub_series(nonop_income, nonop_expense)

        bank_loan = sum_series(short_loan, long_loan, due_in_year)
        interest_debt = sum_series(bank_loan, notes_pay, short_bond, bonds, lease)
        short_interest = sum_series(short_loan, notes_pay, short_bond, due_in_year)

        op_cf = raw(cf, ["经营活动产生的现金流量净额"])
        sales_cash = raw(cf, ["销售商品、提供劳务收到的现金"])
        other_op_in = raw(cf, ["收到的其他与经营活动有关的现金"])
        buy_cash = raw(cf, ["购买商品、接受劳务支付的现金"])
        other_op_out = raw(cf, ["支付的其他与经营活动有关的现金"])
        inv_cf = raw(cf, ["投资活动产生的现金流量净额"])
        recover_inv = raw(cf, ["收回投资所收到的现金"])
        inv_proceed = raw(cf, ["取得投资收益所收到的现金"])
        pay_inv = raw(cf, ["投资所支付的现金"])
        other_inv_out = raw(cf, ["支付的其他与投资活动有关的现金"])
        fin_cf = raw(cf, ["筹资活动产生的现金流量净额"])
        absorb = raw(cf, ["吸收投资收到的现金"])
        borrow = raw(cf, ["取得借款收到的现金"])
        other_fin_in = raw(cf, ["收到其他与筹资活动有关的现金"])
        repay = raw(cf, ["偿还债务支付的现金"])
        dividend = raw(cf, ["分配股利、利润或偿付利息所支付的现金"])
        other_fin_out = raw(cf, ["支付其他与筹资活动有关的现金"])
        fx = raw(cf, ["汇率变动对现金及现金等价物的影响"])
        cash_increase = raw(cf, ["现金及现金等价物净增加额"])
        cash_end = raw(cf, ["期末现金及现金等价物余额"])
        dep_fixed = raw(cf, ["固定资产折旧、油气资产折耗、生产性物资折旧"])
        dep_invprop = raw(cf, ["投资性房地产折旧"])
        amort_ia = raw(cf, ["无形资产摊销"])
        amort_lt = raw(cf, ["长期待摊费用摊销"])

        for key, series in {
            "资产总计": assets,
            "流动资产合计": current_assets,
            "非流动资产合计": noncurrent_assets,
            "负债合计": liabilities,
            "流动负债合计": current_liab,
            "非流动负债合计": noncurrent_liab,
            "所有者权益合计": equity,
        }.items():
            if all(v is None for v in series):
                self.warn(f"【严重告警】关键合计科目缺失：{key}")

        ebitda = self._ebitda(days, total_profit, net_profit, interest, dep_fixed, dep_invprop, amort_ia, amort_lt)
        annual_flags = [is_annual(day) for day in days]

        def cover(numer, denom):
            out = []
            for flag, a, b in zip(annual_flags, numer, denom):
                if not flag:
                    out.append("-")
                elif a is None or b is None or b == 0:
                    out.append(None)
                else:
                    out.append(float(round_half_up(a / b, 2)))
            return out

        roa = self._roa(books, days, assets, total_profit, interest)
        gross = ratio_series(sub_series(revenue, cost), revenue, 4)
        net_margin = ratio_series(net_profit, revenue, 4)
        subsidy_ratio = ratio_series(other_income, net_profit, 4)
        debt_ratio = ratio_series(liabilities, assets, 4)
        current_ratio = ratio_series(current_assets, current_liab, 2)
        quick_base = sub_series(current_assets, inventory)
        quick_ratio = ratio_series(quick_base, current_liab, 2)
        cash_cover = ratio_series(cash, short_interest, 2)

        amounts = {
            "总资产": assets,
            "流动资产": current_assets,
            "非流动资产": noncurrent_assets,
            "长期股权投资": long_equity,
            "总负债": liabilities,
            "流动负债": current_liab,
            "非流动负债": noncurrent_liab,
            "有息负债（含应付票据）": interest_debt,
            "银行贷款": bank_loan,
            "应付债券": bonds,
            "所有者权益": equity,
            "营业收入": revenue,
            "投资收益": invest_income,
            "其他收益": other_income,
            "净利润": net_profit,
            "经营性现金流净额": op_cf,
            "投资性现金流净额": inv_cf,
            "筹资性现金流净额": fin_cf,
            "一、营业收入": revenue,
            "减：营业成本": cost,
            "三费": three_fees,
            "资产减值损失": impairment,
            "其中：财政补贴": other_income,
            "二、营业利润（损失以“－“号填列）": operate_profit,
            "加：营业外净收入": nonop_net,
            "三、利润总额（损失以“－“号填列）": total_profit,
            "四、净利润（损失以“－“号填列）": net_profit,
            "财务费用": finance,
            "管理费用": admin,
            "销售费用": sell,
            "经营活动净现金流量": op_cf,
            "销售商品、提供劳务收到的现金": sales_cash,
            "收到的其他与经营活动有关的现金": other_op_in,
            "购买商品、接受劳务支付的现金": buy_cash,
            "支付的其他与经营活动有关的现金": other_op_out,
            "投资活动净现金流量": inv_cf,
            "收回投资所收到的现金": recover_inv,
            "取得投资收益所收到的现金": inv_proceed,
            "投资所支付的现金": pay_inv,
            "支付的其他与投资活动有关的现金": other_inv_out,
            "筹资活动净现金流量": fin_cf,
            "吸收投资收到的现金": absorb,
            "取得借款收到的现金": borrow,
            "收到其他与筹资活动有关的现金": other_fin_in,
            "偿还债务支付的现金": repay,
            "分配股利、利润或偿付利息所支付的现金": dividend,
            "支付其他与筹资活动有关的现金": other_fin_out,
            "净现金流量": cash_increase,
        }
        percents = {
            "资产负债率": debt_ratio,
            "资产负债率%": debt_ratio,
            "净利润中财政补贴占比（%）": subsidy_ratio,
            "毛利率%": gross,
            "营业净利率%": net_margin,
            "总资产报酬率%": roa,
        }
        ratios = {
            "流动比率": current_ratio,
            "速动比率": quick_ratio,
            "有息负债/ebitda": cover(interest_debt, ebitda),
            "短期有息负债/ebitda": cover(short_interest, ebitda),
            "有息负债/营业收入": cover(interest_debt, revenue),
            "有息负债/净利润": cover(interest_debt, net_profit),
            "货币资金/短期有息负债": cash_cover,
        }
        body = self._balance_body(bs, days, selected_idx, consumed, assets, liabilities, equity)
        cf_body = self._cashflow_body(cf, days, consumed)
        turnover = self._turnover(books, days, revenue, cost, ar, inventory)
        store = {
            "cash_end": cash_end,
            "fx": fx,
            "tax": tax,
            "body": body,
        }
        return {
            "amounts": {fold_label(k): scale_series(v, self.div) for k, v in amounts.items()},
            "percents": {fold_label(k): v for k, v in percents.items()},
            "ratios": {fold_label(k): v for k, v in ratios.items()},
            "body": body,
            "cf_body": cf_body,
            "turnover": turnover,
            "raw_for_check": {
                "assets": scale_series(assets, self.div),
                "current_assets": scale_series(current_assets, self.div),
                "noncurrent_assets": scale_series(noncurrent_assets, self.div),
                "liabilities": scale_series(liabilities, self.div),
                "current_liab": scale_series(current_liab, self.div),
                "noncurrent_liab": scale_series(noncurrent_liab, self.div),
                "equity": scale_series(equity, self.div),
                "capital": scale_series(raw(bs, ["实收资本（或股本）", "实收资本(或股本)"]), self.div),
                "capital_reserve": scale_series(raw(bs, ["资本公积"]), self.div),
                "oci": scale_series(raw(bs, ["其他综合收益"]), self.div),
                "surplus": scale_series(raw(bs, ["盈余公积"]), self.div),
                "retained": scale_series(raw(bs, ["未分配利润"]), self.div),
                "minority": scale_series(raw(bs, ["少数股东权益"]), self.div),
                "operate_profit": scale_series(operate_profit, self.div),
                "nonop_net": scale_series(nonop_net, self.div),
                "total_profit": scale_series(total_profit, self.div),
                "tax": scale_series(tax, self.div),
                "net_profit": scale_series(net_profit, self.div),
                "op_cf": scale_series(op_cf, self.div),
                "inv_cf": scale_series(inv_cf, self.div),
                "fin_cf": scale_series(fin_cf, self.div),
                "fx": scale_series(fx, self.div),
                "cash_increase": scale_series(cash_increase, self.div),
                "cash": scale_series(cash, self.div),
                "cash_end": scale_series(cash_end, self.div),
                "balance_total": scale_series(raw(bs, ["负债和所有者权益（或股东权益）总计", "负债和所有者权益总计"]), self.div),
            },
            "_store": store,
        }

    def _ebitda(self, days, total_profit, net_profit, interest, dep_fixed, dep_invprop, amort_ia, amort_lt):
        if self.ebitda_mode == "legacy":
            return list(net_profit)
        out = []
        warned = False
        for i, day in enumerate(days):
            if total_profit[i] is None:
                out.append(None)
                continue
            pieces = [interest[i], dep_fixed[i], dep_invprop[i], amort_ia[i], amort_lt[i]]
            if all(piece is None for piece in pieces):
                out.append(None)
                if is_annual(day) and not warned:
                    self.warn("【严重告警】标准EBITDA缺少利息、折旧、摊销，相关比率按空值处理。")
                    warned = True
                continue
            total = total_profit[i]
            for piece in pieces:
                if piece is not None:
                    total += piece
            out.append(total)
        return out

    def _roa(self, books, days, assets, total_profit, interest):
        bs = books["bs"]
        all_days = [day for day in bs.periods if day is not None]
        asset_by_day = {}
        for day in all_days:
            asset_by_day[day] = value_of(bs, day, ["资产总计"])
        out = []
        warned = False
        for i, day in enumerate(days):
            if not is_annual(day):
                out.append("-")
                continue
            end = assets[i]
            profit = total_profit[i]
            if end is None or profit is None:
                out.append(None)
                continue
            prev = previous_day(all_days, day)
            begin = asset_by_day.get(prev) if prev else None
            average = (begin + end) / Decimal("2") if begin is not None else end
            extra = interest[i]
            if extra is None and not warned:
                self.warn("【严重告警】利息费用缺失，总资产报酬率仅以利润总额计算。")
                warned = True
            base = profit + (extra or Decimal("0"))
            if average == 0:
                out.append(None)
            else:
                out.append(float(round_half_up(base / average, 4)))
        return out

    def _turnover(self, books, days, revenue, cost, ar, inventory):
        bs = books["bs"]
        all_days = [day for day in bs.periods if day is not None]

        def history(aliases):
            data = {}
            for day in all_days:
                data[day] = value_of(bs, day, aliases)
            return data

        ar_hist = history(["应收账款"])
        inv_hist = history(["存货"])
        annual_days = [day for day in days if is_annual(day)]

        def rate(hist, flow):
            result = []
            for day in annual_days:
                idx = days.index(day)
                end = hist.get(day)
                prev = previous_day(all_days, day)
                begin = hist.get(prev) if prev else None
                numerator = flow[idx]
                if numerator is None or end is None:
                    result.append(None)
                    continue
                average = (begin + end) / Decimal("2") if begin is not None else end
                if average == 0:
                    result.append(None)
                else:
                    result.append(float(round_half_up(numerator / average, 2)))
            return annual_days, result

        ar_days, ar_rate = rate(ar_hist, revenue)
        _, inv_rate = rate(inv_hist, cost)
        return {"days": ar_days, "应收账款周转率": ar_rate, "存货周转率": inv_rate}

    def _balance_body(self, bs, days, selected_idx, consumed, assets, liabilities, equity):
        bases = {"asset": assets, "liability": liabilities, "equity": equity}
        body = []
        for role, display, aliases, base in BS_LINES:
            if role == "header":
                body.append({"role": "header", "label": display, "amounts": None, "ratios": None})
                continue
            line = resolve_line(bs, aliases, consumed, selected_idx)
            raw_values = []
            for day in days:
                if line is None or day not in bs.index:
                    raw_values.append(None)
                else:
                    raw_values.append(line.at(bs.index[day]))
            keep = role in {"subtotal", "total"} or any(meaningful(v) for v in raw_values)
            if not keep:
                continue
            scaled = scale_series(raw_values, self.div)
            ratios = ratio_series(raw_values, bases[base], 4)
            body.append({"role": role, "label": display, "amounts": scaled, "ratios": ratios})
        return body

    def _cashflow_kind(self, line: Line) -> str:
        norm = line.norm
        if "补充资料" in line.name or "补充信息" in line.name:
            return "stop"
        if "小计" in norm or norm.startswith("注:"):
            return "skip"
        if "期初现金" in norm or "期末现金" in norm:
            return "skip"
        if "净增加" in norm and "现金" in norm:
            return "end"
        if ("经营活动" in norm and "净额" in norm) or norm.endswith("经营活动产生的现金流量净额"):
            return "net_op"
        if "投资活动" in norm and "净额" in norm:
            return "net_inv"
        if "筹资活动" in norm and "净额" in norm:
            return "net_fin"
        if ("产生的现金流量" in norm or "产生现金流量" in norm) and "净额" not in norm and "净增加" not in norm:
            return "skip"
        if re.match(r"^[一二三四五六七八九十]+、", norm) and "净额" not in norm and "净增加" not in norm and "汇率" not in norm:
            return "skip"
        return "detail"

    def _cashflow_body(self, cf: Statement, days: list[date], consumed: set) -> list[dict]:
        labels = {
            "net_op": "经营活动净现金流量",
            "net_inv": "投资活动净现金流量",
            "net_fin": "筹资活动净现金流量",
        }
        seen = set()
        bucket: list[tuple[str, list]] = []
        rows: list[dict] = []

        def take(line: Line) -> list:
            consumed.add((line.sheet, line.row))
            values = []
            for day in days:
                if day not in cf.index:
                    values.append(None)
                else:
                    values.append(line.at(cf.index[day]))
            return values

        def flush_details() -> None:
            for name, values in bucket:
                if any(meaningful(v) for v in values):
                    indent = "      " if norm_label(name).startswith("其中") else "    "
                    rows.append({
                        "role": "detail",
                        "label": indent + str(name).strip(),
                        "amounts": scale_series(values, self.div),
                    })
            bucket.clear()

        for line in cf.lines:
            kind = self._cashflow_kind(line)
            if kind == "stop":
                consumed.add((line.sheet, line.row))
                break
            values = take(line)
            if kind == "skip":
                continue
            if kind == "detail":
                bucket.append((line.name, values))
                continue
            if kind in labels:
                rows.append({
                    "role": "net",
                    "label": labels[kind],
                    "amounts": scale_series(values, self.div),
                })
                flush_details()
                seen.add(kind)
                continue
            if kind == "end":
                flush_details()
                rows.append({
                    "role": "net",
                    "label": "净现金流量",
                    "amounts": scale_series(values, self.div),
                })
                seen.add("end")
        for kind, label in (
            ("net_op", "经营活动净现金流量"),
            ("net_inv", "投资活动净现金流量"),
            ("net_fin", "筹资活动净现金流量"),
            ("end", "净现金流量"),
        ):
            if kind not in seen:
                self.warn(f"【严重告警】现金流量表缺少：{label}")
        return rows

    def _log_unmapped(self, books, consumed) -> None:
        seen = set()
        for stmt in books.values():
            for line in stmt.lines:
                if (line.sheet, line.row) in consumed:
                    continue
                if not any(meaningful(v) for v in line.values):
                    continue
                if line.norm.startswith("其中") or "补充" in line.name:
                    continue
                if line.name not in seen:
                    seen.add(line.name)
                    self.warn(f"【未映射科目】{line.name}")

    def _near(self, left: Decimal | None, right: Decimal | None) -> bool:
        if left is None or right is None:
            return False
        return abs(left - right) <= TOLERANCE

    def _validate(self, books, days) -> None:
        # 使用已经四舍五入的缩放值做勾稽，容差 ±0.02
        # raw_for_check 在 _build_values 里生成，这里重新取最后一次写入前的缓存
        cached = getattr(self, "_pending_check", None)
        if cached is None:
            return
        self._validate_cached(days, cached)

    def _validate_cached(self, days, data) -> None:
        def add(items):
            nums = [item for item in items if item is not None]
            if len(nums) != len(items):
                return None
            return sum(nums, Decimal("0"))

        for i, day in enumerate(days):
            label = period_label(day)
            self._check(label, "资产总计=流动资产合计+非流动资产合计", data["assets"][i], add([data["current_assets"][i], data["noncurrent_assets"][i]]))
            equity_sum = add([
                data["capital"][i] or Decimal("0"),
                data["capital_reserve"][i] or Decimal("0"),
                data["oci"][i] or Decimal("0"),
                data["surplus"][i] or Decimal("0"),
                data["retained"][i] or Decimal("0"),
                data["minority"][i] or Decimal("0"),
            ])
            # 空值按0参与加总，但合计本身缺失则告警
            self._check(label, "负债合计=流动负债合计+非流动负债合计", data["liabilities"][i], add([data["current_liab"][i], data["noncurrent_liab"][i]]))
            right_equity = None
            if data["liabilities"][i] is not None and data["equity"][i] is not None:
                right_equity = data["liabilities"][i] + data["equity"][i]
            balance_total = data["balance_total"][i] if data["balance_total"][i] is not None else data["assets"][i]
            self._check(label, "负债及所有者权益总计=负债合计+所有者权益合计", balance_total, right_equity)
            self._check(label, "所有者权益合计=实收资本+资本公积+其他综合收益+盈余公积+未分配利润+少数股东权益", data["equity"][i], equity_sum)
            profit_bridge = None
            if data["operate_profit"][i] is not None and data["nonop_net"][i] is not None:
                profit_bridge = data["operate_profit"][i] + data["nonop_net"][i]
            self._check(label, "利润总额=营业利润+营业外净收入", data["total_profit"][i], profit_bridge)
            net_bridge = None
            if data["total_profit"][i] is not None and data["tax"][i] is not None:
                net_bridge = data["total_profit"][i] - data["tax"][i]
            elif data["total_profit"][i] is not None and data["tax"][i] is None:
                net_bridge = data["total_profit"][i]
            self._check(label, "净利润=利润总额-所得税费用", data["net_profit"][i], net_bridge)
            cash_sum = add([
                data["op_cf"][i] if data["op_cf"][i] is not None else Decimal("0"),
                data["inv_cf"][i] if data["inv_cf"][i] is not None else Decimal("0"),
                data["fin_cf"][i] if data["fin_cf"][i] is not None else Decimal("0"),
                data["fx"][i] if data["fx"][i] is not None else Decimal("0"),
            ]) if data["cash_increase"][i] is not None else None
            self._check(label, "经营+投资+筹资+汇率变动=现金净增加额", data["cash_increase"][i], cash_sum)
            if data["cash"][i] is not None and data["cash_end"][i] is not None and abs(data["cash"][i] - data["cash_end"][i]) > TOLERANCE:
                diff = data["cash"][i] - data["cash_end"][i]
                self.warn(
                    f"【勾稽告警】{label} 货币资金与现金及现金等价物期末余额存在差异（允许合理差异），差异 {float(diff):.2f}"
                )

    def _check(self, period: str, name: str, left, right) -> None:
        if left is None or right is None:
            self.warn(f"【勾稽告警】{period} {name} 缺少数据，无法校验")
            return
        if abs(left - right) > TOLERANCE:
            diff = left - right
            self.warn(f"【勾稽告警】{period} {name} 不匹配，差异 {float(diff):.2f}")

    def _write(self, template: Path, output: Path, days: list[date], values: dict) -> None:
        self._pending_check = values["raw_for_check"]
        self._validate_cached(days, values["raw_for_check"])
        wb = load_workbook(template)
        ws = wb.active
        total_row = find_row(ws, "总资产")
        if total_row is None:
            raise ValueError("表2模板中找不到“总资产”起始行。")
        anchor = total_row
        if total_row > 1 and is_block_header(ws.cell(total_row - 1, 1).value, None):
            anchor = total_row - 1
        fixed_records = []
        for row in range(anchor, ws.max_row + 1):
            label = ws.cell(row, 1).value
            slots = []
            col = 2
            while col <= max(ws.max_column, 2):
                if row == anchor and col > 8:
                    break
                slots.append(col)
                col += 1
                if col > 8:
                    break
            header_slots = []
            col = 2
            while col <= ws.max_column and ws.cell(row, col).value not in (None, ""):
                header_slots.append(col)
                col += 1
            nxt = ""
            for look in range(row + 1, min(row + 5, ws.max_row + 1)):
                if ws.cell(look, 1).value:
                    nxt = norm_label(ws.cell(look, 1).value)
                    break
            fixed_records.append({
                "label": label,
                "header": is_block_header(label, ws.cell(row, 2).value),
                "header_slots": header_slots,
                "annual_only": nxt in {norm_label("应收账款周转率"), norm_label("存货周转率")} or norm_label(label) == norm_label("年 度") and nxt.startswith("应收账款"),
            })
        # 年 度 行的下一项就是周转率，上面的判断覆盖它
        for rec, row in zip(fixed_records, range(anchor, ws.max_row + 1)):
            if norm_label(rec["label"]) == "年度" and rec["header_slots"]:
                nxt = ""
                for look in range(row + 1, min(row + 4, ws.max_row + 1)):
                    if ws.cell(look, 1).value:
                        nxt = norm_label(ws.cell(look, 1).value)
                        break
                rec["annual_only"] = "周转率" in nxt

        styles = self._capture_styles(ws)
        old_span = anchor - 3
        body = values["body"]
        new_span = len(body) + 2
        unmerge_from_row(ws, 3)
        if old_span > 0:
            ws.delete_rows(3, old_span)
        if new_span > 0:
            ws.insert_rows(3, new_span)
        self._write_body(ws, body, styles, len(days))
        self._write_period_banner(ws, days, styles)
        new_anchor = find_row(ws, str(fixed_records[0]["label"]))
        if new_anchor is None:
            raise ValueError("生成后未能定位固定表格。")
        self._assert_fixed_labels(ws, new_anchor, fixed_records)
        self._fill_fixed(ws, new_anchor, fixed_records, days, values)
        self._rewrite_cashflow(ws, days, values, styles)
        output.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output)
        wb.close()

    def _capture_styles(self, ws) -> dict:
        detail_row = find_row(ws, "货币资金") or 4
        sub_row = find_row(ws, "流动资产小计") or detail_row
        total_row = find_row(ws, "总资产合计") or sub_row
        section_row = find_row(ws, "资产类") or 3
        return {
            "detail_label": snapshot_style(ws.cell(detail_row, 1)),
            "detail_amount": snapshot_style(ws.cell(detail_row, 2)),
            "detail_ratio": snapshot_style(ws.cell(detail_row, 3)),
            "sub_label": snapshot_style(ws.cell(sub_row, 1)),
            "sub_amount": snapshot_style(ws.cell(sub_row, 2)),
            "sub_ratio": snapshot_style(ws.cell(sub_row, 3)),
            "total_label": snapshot_style(ws.cell(total_row, 1)),
            "total_amount": snapshot_style(ws.cell(total_row, 2)),
            "total_ratio": snapshot_style(ws.cell(total_row, 3)),
            "section": snapshot_style(ws.cell(section_row, 1)),
            "banner": snapshot_style(ws.cell(1, 2)),
            "money_title": snapshot_style(ws.cell(2, 2)),
            "ratio_title": snapshot_style(ws.cell(2, 3)),
            "unit_title": snapshot_style(ws.cell(2, 1)),
            "detail_height": ws.row_dimensions[detail_row].height or 15.15,
            "section_height": ws.row_dimensions[section_row].height or 15.9,
            "cf_net_label": snapshot_style(ws.cell(find_row(ws, "经营活动净现金流量") or detail_row, 1)),
            "cf_net_amount": snapshot_style(ws.cell(find_row(ws, "经营活动净现金流量") or detail_row, 2)),
            "cf_detail_label": snapshot_style(ws.cell(find_row(ws, "销售商品、提供劳务收到的现金") or detail_row, 1)),
            "cf_detail_amount": snapshot_style(ws.cell(find_row(ws, "销售商品、提供劳务收到的现金") or detail_row, 2)),
        }

    def _write_body(self, ws, body, styles, period_count: int) -> None:
        last_col = 1 + period_count * 2
        for offset, item in enumerate(body):
            row = 3 + offset
            role = item["role"]
            if role == "header":
                ws.row_dimensions[row].height = styles["section_height"]
                for col in range(1, last_col + 1):
                    cell = ws.cell(row, col)
                    apply_style(cell, styles["section"])
                    cell.value = item["label"] if col == 1 else None
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
                clear_after(ws, row, last_col)
                continue
            style_key = {"detail": "detail", "subtotal": "sub", "total": "total"}[role]
            ws.row_dimensions[row].height = styles["detail_height"]
            label_cell = ws.cell(row, 1)
            apply_style(label_cell, styles[f"{style_key}_label"])
            label_cell.value = item["label"]
            for i in range(period_count):
                amount_cell = ws.cell(row, 2 + i * 2)
                ratio_cell = ws.cell(row, 3 + i * 2)
                apply_style(amount_cell, styles[f"{style_key}_amount"])
                apply_style(ratio_cell, styles[f"{style_key}_ratio"])
                put_number(amount_cell, item["amounts"][i], "amount")
                put_number(ratio_cell, None if item["amounts"][i] is None else item["ratios"][i], "percent")
            clear_after(ws, row, last_col)

    def _write_period_banner(self, ws, days, styles) -> None:
        unmerge_row(ws, 1)
        last_col = 1 + len(days) * 2
        title = ws.cell(1, 1)
        apply_style(title, styles["unit_title"])
        title.value = "科目/年度"
        title.border = copy(styles["banner"]["border"])
        paint_header(title)
        unit_cell = ws.cell(2, 1)
        apply_style(unit_cell, styles["unit_title"])
        unit_cell.value = f"科目（单位：{self.unit_name}）"
        paint_header(unit_cell)
        end = max(ws.max_column or last_col, last_col)
        for col in range(2, end + 1):
            for header_row in (1, 2):
                cell = ws.cell(header_row, col)
                if isinstance(cell, MergedCell):
                    continue
                cell.value = None
                if col > last_col:
                    cell.border = Border()
        for i, day in enumerate(days):
            left = 2 + i * 2
            right = left + 1
            for col in (left, right):
                banner = ws.cell(1, col)
                apply_style(banner, styles["banner"])
                banner.number_format = "General"
                banner.value = period_label(day) if col == left else None
                paint_header(banner)
            money = ws.cell(2, left)
            ratio = ws.cell(2, right)
            apply_style(money, styles["money_title"])
            apply_style(ratio, styles["ratio_title"])
            money.value = "金额"
            ratio.value = "占比"
            paint_header(money)
            paint_header(ratio)
            ws.merge_cells(start_row=1, start_column=left, end_row=1, end_column=right)
        clear_after(ws, 1, last_col)
        clear_after(ws, 2, last_col)

    def _assert_fixed_labels(self, ws, anchor: int, records: list[dict]) -> None:
        for offset, rec in enumerate(records):
            actual = ws.cell(anchor + offset, 1).value
            expected = rec["label"]
            if norm_label(actual) != norm_label(expected):
                raise RuntimeError(f"固定表格行标签发生变化：第{offset + 1}行 {expected!r} -> {actual!r}")

    def _fill_fixed(self, ws, anchor: int, records, days, values) -> None:
        labels = [period_label(day) for day in days]
        annual_labels = [period_label(day) for day in values["turnover"]["days"]]
        for offset, rec in enumerate(records):
            row = anchor + offset
            label = rec["label"]
            if label is None or str(label).strip() == "":
                continue
            cell_a = ws.cell(row, 1)
            rewritten = rewrite_unit_text(label, self.unit_name)
            if rewritten != label:
                cell_a.value = rewritten
            if rec["header"]:
                use_labels = annual_labels if rec["annual_only"] else labels
                slots = rec["header_slots"] or list(range(2, 2 + len(use_labels)))
                if len(use_labels) > len(slots):
                    self.warn(f"【严重告警】模板期间列少于输出期数，{label} 只写入前{len(slots)}期。")
                    use_labels = use_labels[-len(slots):]
                for col in slots:
                    ws.cell(row, col).value = None
                for col, text in zip(slots, use_labels):
                    header_cell = ws.cell(row, col)
                    header_cell.value = text
                    header_cell.number_format = "General"
                if slots:
                    clear_after(ws, row, max(slots))
                continue
            key = fold_label(label)
            if "周转率" in key:
                series = values["turnover"].get(key.replace(" ", ""))
                # norm 已经去掉空格
                series = values["turnover"].get(key)
                kind = "ratio"
                self._write_across(ws, row, series or [], kind, rec["header_slots"] or [2, 3, 4])
                continue
            if key in values["percents"]:
                self._write_across(ws, row, values["percents"][key], "percent", list(range(2, 2 + len(days))))
            elif key in values["ratios"]:
                self._write_across(ws, row, values["ratios"][key], "ratio", list(range(2, 2 + len(days))))
            elif key in values["amounts"]:
                self._write_across(ws, row, values["amounts"][key], "amount", list(range(2, 2 + len(days))))
            else:
                self.warn(f"【未识别模板行】{label}")
                for col in range(2, 10):
                    ws.cell(row, col).value = None

    def _find_cashflow_header(self, ws) -> int | None:
        for row in range(1, ws.max_row + 1):
            if not norm_label(ws.cell(row, 1).value).startswith("年度"):
                continue
            for look in range(row + 1, min(row + 8, ws.max_row + 1)):
                if norm_label(ws.cell(look, 1).value) == "经营活动净现金流量":
                    return row
        return None

    def _rewrite_cashflow(self, ws, days, values, styles) -> None:
        header = self._find_cashflow_header(ws)
        if header is None:
            self.warn("【严重告警】模板中找不到现金流量区块，明细行未重排。")
            return
        end = None
        for row in range(header + 1, ws.max_row + 1):
            if norm_label(ws.cell(row, 1).value) == "净现金流量":
                end = row
                break
        if end is None:
            self.warn("【严重告警】模板中找不到“净现金流量”，明细行未重排。")
            return
        targets = [
            str(item) for item in ws.merged_cells.ranges
            if item.max_row >= header + 1 and item.min_row <= end
        ]
        for item in targets:
            ws.unmerge_cells(item)
        body = values.get("cf_body") or []
        ws.delete_rows(header + 1, end - header)
        if body:
            ws.insert_rows(header + 1, len(body))
        for offset, item in enumerate(body):
            row = header + 1 + offset
            net = item["role"] == "net"
            label_cell = ws.cell(row, 1)
            apply_style(label_cell, styles["cf_net_label"] if net else styles["cf_detail_label"])
            label_cell.value = item["label"]
            ws.row_dimensions[row].height = styles["detail_height"]
            last_col = 1 + len(item["amounts"])
            for col in range(2, last_col + 1):
                cell = ws.cell(row, col)
                apply_style(cell, styles["cf_net_amount"] if net else styles["cf_detail_amount"])
                cell.value = None
            for index, amount in enumerate(item["amounts"]):
                put_number(ws.cell(row, 2 + index), amount, "amount")
            clear_after(ws, row, last_col)

    def _write_across(self, ws, row: int, series, kind: str, cols: list[int]) -> None:
        end = ws.max_column or 1
        for col in range(2, end + 1):
            if col not in cols:
                cell = ws.cell(row, col)
                if isinstance(cell, MergedCell):
                    continue
                cell.value = None
                cell.border = Border()
        if series is None:
            return
        usable = series[-len(cols):] if len(series) > len(cols) else series
        for col in cols[len(usable):]:
            ws.cell(row, col).value = None
        for col, value in zip(cols, usable):
            put_number(ws.cell(row, col), value, kind)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="由系统导入版财务报表生成资产类财务数据表")
    parser.add_argument("--input", default=str(DEFAULT_INPUT))
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE))
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT))
    parser.add_argument("--unit", choices=["wan", "yi"], default=OUTPUT_UNIT)
    parser.add_argument("--rd", choices=["true", "false"], default="true" if INCLUDE_RD_IN_3FEE else "false")
    parser.add_argument("--ebitda", choices=["std", "legacy"], default=EBITDA_MODE)
    parser.add_argument("--periods", type=int, default=PERIOD_COUNT)
    return parser


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    exporter = Exporter(
        output_unit=args.unit,
        include_rd=args.rd == "true",
        ebitda_mode=args.ebitda,
        period_count=args.periods,
    )
    exporter.run(Path(args.input), Path(args.template), Path(args.output))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ValueError as exc:
        print(exc)
        raise SystemExit(1)
