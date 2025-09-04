# -*- coding: utf-8 -*-
"""
bank_fx_data.py
---------------
数据处理模块：从《银行结售汇数据时间序列.xlsx》抽取图表所需的结构化数据，并计算同比、YTD等指标。

设计要点：
1) 仅依赖 pandas；保持纯函数，方便在 Streamlit 或其它框架中复用。
2) 针对 Excel 中的“以人民币计价（月度）”表做了稳健解析：
   - 首列为分项（项目），第二列通常为空或“结汇/售汇/差额”等标签；后续列为 Excel 日期序列数值（如 40200）。
   - 解析“结汇/售汇/差额”主表；解析“经常项目/资本与金融项目”分项；解析“远期结售汇签约额”“累计未到期额”三元组（结汇/售汇/差额）。
3) 同比=与上年同月相比（pct_change(12)）；YTD=当年1月至最新月份累计，与上年同期累计同比。

返回的主要对象：
- main: DataFrame(index=月末日期, columns=[结汇, 售汇, 差额, 结汇同比, 售汇同比, 差额同比])
- comp: DataFrame(index=月末日期, columns=[结汇_经常项目, 结汇_资本项目, 售汇_经常项目, 售汇_资本项目])
- fwd_sign: DataFrame(index=月末日期, columns=[结汇, 售汇, 差额]) —— 远期结售汇“签约额”
- fwd_out:  DataFrame(index=月末日期, columns=[结汇, 售汇, 差额]) —— 远期结售汇“累计未到期额”
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Optional, Tuple

import numpy as np
import pandas as pd


DEFAULT_PATTERNS = [
    "银行结售汇数据时间序列.xlsx",
    "*银行*结*售*汇*数据*时间*序*.xlsx",
]


def _find_file(path_or_glob: Optional[str] = None) -> Optional[str]:
    """寻找 Excel 文件；允许传入确切路径或通配符。"""
    if path_or_glob and Path(path_or_glob).exists():
        return str(path_or_glob)
    if path_or_glob:
        for p in Path(".").glob(path_or_glob):
            return str(p)
    for pat in DEFAULT_PATTERNS:
        for p in Path(".").glob(pat):
            return str(p)
    return None


def _to_month_end(values) -> pd.DatetimeIndex:
    """将 Excel 列标签（可能是数值序列号/字符串日期）统一到【月末】时间戳。"""
    out = []
    for c in values:
        v = pd.NaT
        try:
            if isinstance(c, (int, float, np.integer, np.floating)) or (
                isinstance(c, str) and str(c).strip().replace(".", "", 1).isdigit()
            ):
                v = pd.to_datetime(float(c), unit="D", origin="1899-12-30", errors="coerce")
            else:
                v = pd.to_datetime(c, errors="coerce")
        except Exception:
            v = pd.NaT
        out.append(v)
    idx = pd.to_datetime(out)
    # 统一到月末（避免日差影响同比/YTD对齐）
    idx = idx.to_period("M").to_timestamp("M")
    return idx


def _normalize_series(s: pd.Series) -> pd.Series:
    """清洗数列：索引转月末、数值转数值、去重并排序。
    注意：若同一月份出现重复列（例如 Excel 中存在两个表示 2023-12 的列），
    取该月“最后一个非空值”（last），避免累计或折线出现重复点。
    """
    s2 = pd.to_numeric(s, errors="coerce")
    s2.index = _to_month_end(s.index)
    s2 = s2.dropna()
    # 月份去重：保留每月最后一个有效值
    s2 = s2.groupby(s2.index).last().sort_index()
    return s2


def _pct_yoy_12(s: pd.Series) -> pd.Series:
    """12期同比（与上年同月相比）。"""
    return s.pct_change(12)


def _block(df: pd.DataFrame, start_label: str, end_label: Optional[str]) -> pd.DataFrame:
    """按“项目”标签切片一个区块（包含 start，不含 end）。"""
    idx = df.index
    try:
        i = idx.get_loc(start_label)
    except KeyError:
        return pd.DataFrame()
    if end_label is None:
        j = len(idx)
    else:
        try:
            j = idx.get_loc(end_label)
        except KeyError:
            j = len(idx)
    return df.iloc[i:j]


def _find_row_in(block: pd.DataFrame, keyword: str) -> Optional[pd.Series]:
    """在区块内模糊查找包含关键字的第一行。忽略空格。"""
    for i, name in enumerate(block.index.astype(str)):
        if keyword in name.replace(" ", ""):
            return block.iloc[i]
    return None


@dataclass
class BankFXData:
    main: pd.DataFrame
    comp: pd.DataFrame
    fwd_sign: pd.DataFrame
    fwd_out: pd.DataFrame


def load_bank_fx(path_or_glob: Optional[str] = None, sheet_hint: str = "以人民币计价（月度）") -> Optional[BankFXData]:
    """读取并解析 Excel，返回结构化数据。"""
    path = _find_file(path_or_glob)
    if not path:
        return None

    try:
        raw = pd.read_excel(path, sheet_name=sheet_hint, header=3, index_col=0)
    except Exception:
        # 兜底：取第一个工作表
        xls = pd.ExcelFile(path)
        raw = pd.read_excel(path, sheet_name=xls.sheet_names[0], header=3, index_col=0)

    # 第二列通常为标签列（结汇/售汇/差额）
    label_col = raw.columns[0]
    date_cols = [c for c in raw.columns if not str(c).startswith("Unnamed")]

    # —— 主表：结汇/售汇/差额 ——
    try:
        settle = _normalize_series(raw.loc["一、结汇", date_cols])
        sale   = _normalize_series(raw.loc["二、售汇", date_cols])
        bal    = _normalize_series(raw.loc["三、差额", date_cols])
    except KeyError:
        return None

    main = pd.DataFrame({"结汇": settle, "售汇": sale})
    main["差额"] = main["结汇"] - main["售汇"]
    for col in ("结汇", "售汇", "差额"):
        main[f"{col}同比"] = _pct_yoy_12(main[col])

    # —— 成分：经常项目 & 资本与金融项目（映射为“资本项目”展示） ——
    blk_settle = _block(raw, "一、结汇", "二、售汇")
    blk_sale   = _block(raw, "二、售汇", "三、差额")

    settle_cur = _find_row_in(blk_settle, "经常项目")
    settle_cap = _find_row_in(blk_settle, "资本与金融项目")
    sale_cur   = _find_row_in(blk_sale,   "经常项目")
    sale_cap   = _find_row_in(blk_sale,   "资本与金融项目")

    comp = pd.DataFrame(index=_to_month_end(date_cols))
    if settle_cur is not None:
        comp["结汇_经常项目"] = _normalize_series(settle_cur[date_cols])
    if settle_cap is not None:
        comp["结汇_资本项目"] = _normalize_series(settle_cap[date_cols])
    if sale_cur is not None:
        comp["售汇_经常项目"] = _normalize_series(sale_cur[date_cols])
    if sale_cap is not None:
        comp["售汇_资本项目"] = _normalize_series(sale_cap[date_cols])
    # 去重：若存在同一月份重复索引，取最后一个有效值
    comp = comp.dropna(how="all")
    if len(comp.index) and comp.index.has_duplicates:
        comp = comp.groupby(comp.index).last()
    comp = comp.sort_index()

    # —— 远期：签约额（三元组） ——
    def _tri_df(section_label: str) -> Optional[pd.DataFrame]:
        try:
            i = raw.index.get_loc(section_label)
        except KeyError:
            return None
        tri = raw.iloc[i : i + 3].copy()
        tri.index = tri[label_col].tolist()  # 结汇/售汇/差额
        tri = tri.drop(columns=[label_col])
        tri = tri.T
        tri.index = _to_month_end(tri.index)
        tri = tri.apply(pd.to_numeric, errors="coerce")
        tri = tri.dropna(how="all")
        # 去重：若同月重复，取最后一个有效值
        if len(tri.index) and tri.index.has_duplicates:
            tri = tri.groupby(tri.index).last()
        tri = tri.sort_index()
        return tri[["结汇", "售汇", "差额"]]

    fwd_sign = _tri_df("四、远期结售汇签约额")
    fwd_out  = _tri_df("七、本期末远期结售汇累计未到期额")

    if fwd_sign is None:
        fwd_sign = pd.DataFrame(columns=["结汇", "售汇", "差额"])
    if fwd_out is None:
        fwd_out  = pd.DataFrame(columns=["结汇", "售汇", "差额"])

    return BankFXData(main=main, comp=comp, fwd_sign=fwd_sign, fwd_out=fwd_out)


# ====== 常用统计 ======
def ytd_sum_and_yoy(s: pd.Series) -> Tuple[Optional[float], Optional[float], Optional[pd.Timestamp]]:
    """返回 (YTD累计, 同比, 最新月份)。同比为当年累计 vs 上年同期累计。"""
    if s is None or s.empty:
        return None, None, None
    s = s.dropna().sort_index()
    last = s.index.max()
    year, month = int(last.year), int(last.month)
    cur = s[s.index.year == year].iloc[:month].sum()
    prev = s[s.index.year == year - 1].iloc[:month].sum() if (year - 1) in s.index.year else np.nan
    yoy = (cur / prev - 1.0) if pd.notna(prev) and prev != 0 else np.nan
    return float(cur), (float(yoy) if pd.notna(yoy) else None), last


def gross_amount(df_main: pd.DataFrame) -> pd.Series:
    """结售汇总额（结汇+售汇）。"""
    return (df_main["结汇"] + df_main["售汇"]).dropna()


def slice_last_months(df: pd.DataFrame, months: int) -> pd.DataFrame:
    """截取最近N个月（按索引时间）。"""
    if df is None or df.empty:
        return df
    return df[df.index >= (df.index.max() - pd.offsets.MonthEnd(months))]


def get_dashboard_data(path_or_glob: Optional[str] = None, months: int = 36) -> Optional[Dict[str, pd.DataFrame]]:
    """一次性加载并裁剪用于看板的所有数据。"""
    data = load_bank_fx(path_or_glob=path_or_glob)
    if data is None:
        return None
    main = slice_last_months(data.main, months)
    comp = slice_last_months(data.comp, months)
    fwd_sign = slice_last_months(data.fwd_sign, months)
    fwd_out  = slice_last_months(data.fwd_out, months)
    return {"main": main, "comp": comp, "fwd_sign": fwd_sign, "fwd_out": fwd_out}