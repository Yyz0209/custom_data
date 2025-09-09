import os
import subprocess
import sys
import time
from datetime import datetime
import json

import pandas as pd
import streamlit as st
from pyecharts import options as opts
from pyecharts.charts import Line, Bar, HeatMap
from pyecharts.commons.utils import JsCode
from streamlit_echarts import st_pyecharts

from config import OUTPUT_FILENAME, FINAL_LOCATIONS

st.set_page_config(
    page_title="跨境数据看板",
    layout="wide",
    page_icon=None,
)


PALETTE = [
    "#5C8D57",  # muted green
    "#5C7491",  # slate blue
    "#DDBB49",  # mustard
    "#CF5A5A",  # soft red
    "#A688A3",  # mauve
    "#78B0A0",  # teal
    "#E3943A",  # orange
    "#B6AAA3",  # greige
]

st.markdown(
    """
<style>
/* 明亮主题与字体（莫兰迪暖色系） */
:root {
  /* 背景与容器：柔和暖白、低饱和灰米色 */
  --bg: #F3F0EC;        /* 背景：暖米白 */
  --panel: #FAF7F3;     /* 面板：更浅的暖色 */
  --card: #FFFFFF;      /* 卡片：保持纯白，避免过度泛黄 */
  --border: #E7E2DC;    /* 边框：米灰 */

  /* 文本与辅助色：偏暖的深灰与柔和灰褐 */
  --text: #2D2A26;
  --muted: #7A756C;

  /* 强调色：莫兰迪系的柔和蓝灰与鼠尾草绿 */
  --primary: #6B7A8F;
  --accent: #9AAE8C;
}

/* 避免重置所有 st-* 组件的背景，防止 Slider 轨道被盖掉 */
html, body {
  font-family: Inter, "HarmonyOS Sans SC", "Microsoft YaHei", system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif;
}
[data-testid="stAppViewContainer"] {
  background: var(--bg);
  color: var(--text);
}

/* Slider: clean, full-width, no shadow */
.stSlider { width: 100%; padding: 0; margin: 0; }
.stSlider [data-baseweb="slider"] { width: 100%; }
.stSlider [data-baseweb="slider"] * { box-shadow: none !important; }
.stSlider [data-baseweb="slider"] > div,
[data-baseweb="slider"] > div {
  background: transparent !important; /* 取消底部灰色轨道，仅保留选中条 */
  height: 6px; border-radius: 999px; box-shadow: none !important;
}
.stSlider [data-baseweb="slider"] [aria-hidden="true"],
[data-baseweb="slider"] [aria-hidden="true"] {
  background: var(--accent) !important; /* 仅显示选中条 */
  height: 6px; border-radius: 999px;
}
.stSlider [data-baseweb="slider"] [role="slider"] {
  box-shadow: none !important; outline: none !important;
  background: #fff !important; border: 2px solid var(--accent) !important;
  width: 14px; height: 14px; border-radius: 50%;
}
.stSlider [data-baseweb="slider"] [role="slider"]:hover {
  transform: none !important;
}
.stSlider span { color: var(--muted) !important; }

/* 顶部标题条（简洁风格） */
.hero {
  background: #FFFFFF;              /* 纯色背景，更协调 */
  border: 1px solid var(--border);  /* 细边框 */
  border-radius: 12px;
  padding: 20px 20px 16px 20px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}
.hero h1 { margin: 0; font-weight: 700; letter-spacing: .3px; color: var(--text); }
.hero .sub { color: var(--muted); margin-top: 8px; }

/* 块容器（面板） */
.panel {
  background: var(--panel);
  border: 1px solid var(--border);
  border-radius: 14px;
  padding: 18px 18px 8px 18px;
  box-shadow: 0 6px 18px rgba(0,0,0,0.05);
}

/* 指标卡片容器 */
.card {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 14px;
  padding: 16px 16px 12px 16px;
  height: 100%;
  box-shadow: 0 8px 20px rgba(0,0,0,0.06);
}
.card .label { color: var(--muted); font-size: 14px; }
.card .value { font-size: 28px; font-weight: 600; margin-top: 6px; letter-spacing: .2px; color: var(--text); }
.chip {
  display:inline-block; margin-top: 10px; padding: 3px 8px; border-radius: 999px; font-size: 12px; font-weight: 600;
}
.chip.up { background: #F5D7D7; color: #8A3B3B; border: 1px solid #E7B8B8; }
.chip.down { background: #E3F0E8; color: #2F6B4F; border: 1px solid #BBD4C4; }

/* 图表圆角容器（仅做裁剪与轻边框，不是卡片） */
.chart-rounded {
  border-radius: 12px;
  overflow: hidden;
  border: 1px solid var(--border);
}

/* 侧边栏优化 */
[data-testid="stSidebar"] { background: #FFFFFF; border-right: 1px solid var(--border); }

/* 分割线更柔和 */
hr { border: none; height: 1px; background: linear-gradient(90deg, transparent, var(--border), transparent); }

</style>
""",
    unsafe_allow_html=True,
)


@st.cache_data
def load_data():
    if not os.path.exists(OUTPUT_FILENAME):
        return None
    try:
        xls = pd.ExcelFile(OUTPUT_FILENAME)
        return {sheet: xls.parse(sheet) for sheet in xls.sheet_names}
    except Exception as exc:
        st.error(f"加载 Excel 失败: {exc}")
        return None


def run_data_updater():
    try:

        result = subprocess.run(
            [sys.executable, "自动数据更新器.py"],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore',
            timeout=600
        )

        if result.returncode == 0:
            output = result.stdout or ""
            output_lines = output.strip().split('\n') if output else []

            has_updates = False
            message = "更新检查完成"

            for line in output_lines:
                if line and "发现新数据并已更新" in line:
                    has_updates = True
                    message = line.split("消息: ")[-1] if "消息: " in line else line
                    break
                elif line and "当前已是最新数据" in line:
                    message = line.split("消息: ")[-1] if "消息: " in line else line
                    break

            return {
                "success": True,
                "has_updates": has_updates,
                "message": message,
                "output": result.stdout
            }
        else:
            error_msg = result.stderr or "未知错误"
            return {
                "success": False,
                "message": f"更新失败: {error_msg}",
                "output": error_msg
            }

    except subprocess.TimeoutExpired:
        return {
            "success": False,
            "message": "更新超时（超过10分钟），请稍后重试",
            "output": ""
        }
    except Exception as e:
        return {
            "success": False,
            "message": f"更新过程中出错: {str(e)}",
            "output": ""
        }


def fmt_value(v):
    if pd.isna(v):
        return "—"
    return f"{v:,.0f}"


def fmt_delta(y):
    if pd.isna(y):
        return None, ""
    val = f"{y:+.2%}"
    cls = "up" if y > 0 else "down"
    return cls, val


def render_card(label, val, yoy):
    cls, delta = fmt_delta(yoy if pd.notna(yoy) else None)
    chip_html = f'<span class="chip {cls}">{delta}</span>' if delta else ""
    return f'<div class="card"><div class="label">{label}</div><div class="value">{fmt_value(val)}</div>{chip_html}</div>'



def nice_percent_axis(vals):
    xs = [v for v in vals if v is not None and not pd.isna(v)]
    if not xs:
        return None, None
    vmin, vmax = min(xs), max(xs)
    if vmin == vmax:
        vmin -= 5
        vmax += 5
    span = vmax - vmin
    vmin -= span * 0.1
    vmax += span * 0.1
    import math
    nice_min = math.floor(vmin / 5.0) * 5
    nice_max = math.ceil(vmax / 5.0) * 5
    nice_min = min(nice_min, 0)
    nice_max = max(nice_max, 0)
    if nice_max - nice_min < 10:
        nice_max = nice_min + 10
    return nice_min, nice_max


def show_chart(chart, height="520px"):
    st.markdown('<div class="chart-rounded">', unsafe_allow_html=True)
    st_pyecharts(chart, height=height)
    st.markdown('</div>', unsafe_allow_html=True)

# === CME FedWatch loaders and charts ===
@st.cache_data
def load_fedwatch_probabilities(path: str = os.path.join("outputs", "fedwatch_probabilities.json")):
    try:
        if not os.path.exists(path):
            return None
        with open(path, "r", encoding="utf-8") as f:
            obj = json.load(f)
        rows = obj.get("rows", [])
        if not rows:
            return None
        header_row = rows[0]
        x_labels = [str(x).strip() for x in header_row[1:]]
        data_rows = rows[1:]
        y_labels = [r[0] for r in data_rows]
        # values matrix as floats (%), missing -> None
        def _to_num(s):
            s = str(s).strip()
            if not s or s == "—":
                return None
            s = s.replace("%", "")
            try:
                return float(s)
            except Exception:
                return None
        matrix = [[_to_num(v) for v in r[1:1+len(x_labels)]] for r in data_rows]
        return {"x": x_labels, "y": y_labels, "z": matrix}
    except Exception:
        return None


def build_fedwatch_heatmap(prob_data, title=""):
    if not prob_data:
        return None
    x_labels = prob_data["x"]
    # 需求：日期早的在上面 -> 反转 y 与矩阵行
    y_labels_orig = prob_data["y"]
    z_orig = prob_data["z"]
    y_labels = list(reversed(y_labels_orig))
    z = list(reversed(z_orig))
    # HeatMap requires list of [x_index, y_index, value]
    data = []
    vmax = 0
    for yi, row in enumerate(z):
        for xi, v in enumerate(row):
            if v is None:
                continue
            vmax = max(vmax, v)
            data.append([xi, yi, float(v)])
    vmax = max(vmax, 100)
    hm = HeatMap(init_opts=opts.InitOpts(bg_color="#FFFFFF", width="100%"))
    hm.add_xaxis(x_labels)
    hm.add_yaxis(
        "概率(%)",
        y_labels,
        data,
        label_opts=opts.LabelOpts(
            is_show=True,
            color="#2D2A26",
            formatter=JsCode("function(p){return (p.value[2]).toFixed(1)+'%';}")
        ),
        itemstyle_opts=opts.ItemStyleOpts(border_color="#E6E1DA", border_width=1)
    )
    hm.set_global_opts(
        title_opts=opts.TitleOpts(title=""),
        tooltip_opts=opts.TooltipOpts(is_show=True, formatter=JsCode("function(p){return p.name + '<br/>概率: ' + (p.value[2]).toFixed(1) + '%';}")),
        xaxis_opts=opts.AxisOpts(position="top", axislabel_opts=opts.LabelOpts(rotate=0)),
        yaxis_opts=opts.AxisOpts(axislabel_opts=opts.LabelOpts(rotate=0)),
        visualmap_opts=opts.VisualMapOpts(
            max_=100,
            min_=0,
            orient="vertical",
            pos_right="10",
            pos_top="middle",
            range_color=["#ffffff", "#fff3cd", "#ffe08a", "#cdeaf7", "#8fd3f4", "#4aa3e0"]
        )
    )
    return hm


@st.cache_data
def load_fedwatch_dotplot(path: str = os.path.join("outputs", "fedwatch_dot_plot_table.json")):
    try:
        if not os.path.exists(path):
            return None
        with open(path, "r", encoding="utf-8") as f:
            obj = json.load(f)
        header = obj.get("header", [])
        rows = obj.get("rows", [])
        if not header or not rows:
            return None
        x_labels = [str(h).strip() for h in header[1:]]
        # y labels are target rates, sort descending numerically
        def _num_rate(s):
            try:
                return float(str(s).strip())
            except Exception:
                return None
        y_pairs = [(_num_rate(r[0]), str(r[0])) for r in rows]
        y_pairs = [p for p in y_pairs if p[0] is not None]
        y_pairs.sort(key=lambda p: p[0], reverse=True)
        y_labels = [p[1] for p in y_pairs]
        # build value dict
        row_by_rate = {str(r[0]): r[1:] for r in rows}
        matrix = []
        for y in y_labels:
            r = row_by_rate.get(y, [])
            vals = []
            for v in r[:len(x_labels)]:
                s = str(v).strip()
                vals.append(int(s) if s.isdigit() else 0)
            matrix.append(vals)
        return {"x": x_labels, "y": y_labels, "z": matrix}
    except Exception:
        return None





def build_fedwatch_dot_table(dp_data):
    """将点阵数据转成表格（DataFrame）展示。
    第一列为 TARGET RATE，后续为各年份/长期列；0 显示为空白。
    """
    if not dp_data:
        return None
    x_labels = dp_data.get("x", [])
    y_labels = dp_data.get("y", [])
    z = dp_data.get("z", [])
    if not x_labels or not y_labels or not z:
        return None
    rows = []
    for yi, row in enumerate(z):
        cells = []
        for v in row[: len(x_labels)]:
            try:
                iv = int(v)
            except Exception:
                iv = 0
            cells.append("" if iv == 0 else str(iv))
        if len(cells) < len(x_labels):
            cells += [""] * (len(x_labels) - len(cells))
        rows.append([str(y_labels[yi])] + cells)
    columns = ["TARGET RATE"] + [str(x) for x in x_labels]
    try:
        df = pd.DataFrame(rows, columns=columns)
        # 确保所有列都是字符串，避免 Arrow 类型冲突
        for c in df.columns:
            df[c] = df[c].astype(str)
    except Exception:
        return None
    return df




def build_line(x_list, series_dict, title):

    chart = Line(init_opts=opts.InitOpts(bg_color="#FFFFFF"))
    chart.add_xaxis(xaxis_data=x_list)

    chart.set_colors(PALETTE)
    for name, y in series_dict.items():
        chart.add_yaxis(
            series_name=name,
            y_axis=y,
            label_opts=opts.LabelOpts(is_show=False),
        )
    chart.set_global_opts(
        title_opts=opts.TitleOpts(
            title=title,
            pos_left="center",
            title_textstyle_opts=opts.TextStyleOpts(color="#111827"),
        ),
        tooltip_opts=opts.TooltipOpts(trigger="axis"),
        toolbox_opts=opts.ToolboxOpts(is_show=True),
        xaxis_opts=opts.AxisOpts(type_="category", boundary_gap=False),
        yaxis_opts=opts.AxisOpts(name="金额"),
        legend_opts=opts.LegendOpts(orient="horizontal", pos_top="40"),
    )
    return chart



CATEGORY_ORDER = [
    "农产品", "矿产品", "化学制品", "纺织服装", "木制品",
    "金属石料制品", "电子设备", "交通设备", "其他制品",
]
CATEGORY_COLOR_MAP = {
    "农产品": "#59A14F",
    "矿产品": "#4E79A7",
    "化学制品": "#EDC948",
    "纺织服装": "#E15759",
    "木制品": "#9C755F",
    "金属石料制品": "#B07AA1",
    "电子设备": "#76B7B2",
    "交通设备": "#F28E2B",
    "其他制品": "#BAB0AC",
}
CATEGORY_PALETTE = [CATEGORY_COLOR_MAP[c] for c in CATEGORY_ORDER]

# 修复：由于源码中 CATEGORY_ORDER 存在编码损坏，新增一套标准化的类别与配色
CATEGORY_LABELS = [
    "农产品",
    "矿产品",
    "化学制品",
    "纺织服装",
    "木制品",
    "金属石料制品",
    "电子设备",
    "交通设备",
    "其他制品",
]
CATEGORY_COLOR = {
    "农产品": "#59A14F",
    "矿产品": "#4E79A7",
    "化学制品": "#EDC948",
    "纺织服装": "#E15759",
    "木制品": "#9C755F",
    "金属石料制品": "#B07AA1",
    "电子设备": "#76B7B2",
    "交通设备": "#F28E2B",
    "其他制品": "#BAB0AC",
}
CATEGORY_PALETTE2 = [CATEGORY_COLOR[c] for c in CATEGORY_LABELS]


@st.cache_data
def load_category_data2():
    filename = "9大类产品分析表.xlsx"
    if not os.path.exists(filename):
        return None
    try:
        xls = pd.ExcelFile(filename)
        data = {}
        for sheet in xls.sheet_names:
            df = pd.read_excel(filename, sheet_name=sheet, index_col=0)
            # normalize labels and numbers
            def _norm_label(s: str) -> str:
                t = str(s).replace('\u3000',' ').replace('\xa0',' ').strip()
                t = t.replace('（','(').replace('）',')')
                t = t.replace('其它', '其他')
                alias = {
                    '农林牧渔产品': '农产品',
                    '矿物产品': '矿产品',
                    '矿产': '矿产品',
                    '木材及制品': '木制品',
                    '电子电气设备': '电子设备',
                    '电子及电气设备': '电子设备',
                    '金属及石料制品': '金属石料制品',
                    '交通运输设备': '交通设备',
                    '交通运输设备制造': '交通设备',
                    '交通装备': '交通设备',
                    '其它制品': '其他制品',
                    '其他': '其他制品',
                }
                if t.startswith('其他制品') or t.startswith('其它制品'):
                    t = '其他制品'
                return alias.get(t, t)

            df.index = df.index.map(lambda x: str(x).replace('\u3000',' ').replace('\xa0',' ').strip())
            df.rename(columns=lambda c: _norm_label(c), inplace=True)
            for c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce')
            # 若‘其他制品’缺失，尝试用合计减去其余八类推算
            if '其他制品' not in df.columns:
                total_candidates = [
                    '合计','总计','进出口合计','进出口总额','总额',
                    '合计(美元)','合计(人民币)','进出口_合计','进出口_总计'
                ]
                tot = next((c for c in total_candidates if c in df.columns), None)
                if tot:
                    others = [c for c in CATEGORY_LABELS if c != '其他制品' and c in df.columns]
                    if others:
                        df['其他制品'] = (pd.to_numeric(df[tot], errors='coerce') - df[others].sum(axis=1))
            data[sheet] = df
        return data
    except Exception as e:
        st.error(f"加载9大类产品 Excel 失败: {e}")
        return None


def create_horizontal_percentage_chart(data, title, regions, categories):


    categories_ordered = [c for c in CATEGORY_LABELS if c in categories]
    if not categories_ordered:
        categories_ordered = categories
    data = data.reindex(columns=categories_ordered, fill_value=0)


    row_sums = data.sum(axis=1).replace(0, 1)
    percentage_data = data.div(row_sums, axis=0) * 100.0


    reversed_regions = list(reversed(regions))


    chart_height = max(400, len(regions) * 60 + 150)


    chart = Bar(init_opts=opts.InitOpts(bg_color="#FFFFFF", width="100%", height=f"{chart_height}px"))


    chart.add_xaxis(reversed_regions)


    for category in categories_ordered:
        values = []
        for region in reversed_regions:
            if region in percentage_data.index and category in percentage_data.columns:
                val = percentage_data.loc[region, category]
                values.append(float(val) if pd.notna(val) else 0)
            else:
                values.append(0)

        chart.add_yaxis(
            series_name=category,
            y_axis=values,
            stack="stack1",
            label_opts=opts.LabelOpts(is_show=False),
            category_gap="35%",
            itemstyle_opts=opts.ItemStyleOpts(color=CATEGORY_COLOR.get(category)),
        )

    chart.set_colors(CATEGORY_PALETTE2)

    chart.set_global_opts(
        title_opts=opts.TitleOpts(
            title=title,
            pos_left="center",
            title_textstyle_opts=opts.TextStyleOpts(color="#2D2A26", font_size=16),
        ),
        tooltip_opts=opts.TooltipOpts(
            trigger="axis",
            axis_pointer_type="shadow",
            formatter="{b}<br/>{a}: {c:.1f}%",
        ),
        legend_opts=opts.LegendOpts(
            orient="horizontal",
            pos_top="40px",
            type_="scroll",
            selected_mode="multiple",
            textstyle_opts=opts.TextStyleOpts(color="#2D2A26"),
        ),
        xaxis_opts=opts.AxisOpts(
            type_="value",
            min_=0,
            max_=100,
            axislabel_opts=opts.LabelOpts(formatter="{value}%", color="#2D2A26"),
            name="占比（%）",
            name_textstyle_opts=opts.TextStyleOpts(color="#2D2A26"),
        ),
        yaxis_opts=opts.AxisOpts(
            type_="category",
            axislabel_opts=opts.LabelOpts(color="#2D2A26"),
        ),
        toolbox_opts=opts.ToolboxOpts(
            is_show=True,
            feature=opts.ToolBoxFeatureOpts(
                save_as_image=opts.ToolBoxFeatureSaveAsImageOpts(),
                data_zoom=opts.ToolBoxFeatureDataZoomOpts(),
                restore=opts.ToolBoxFeatureRestoreOpts(),
            ),
        ),
    )


    chart.reversal_axis()
    return chart




@st.cache_data
def load_fx_deposit_loan():
    """
    从 pbc_fx_credit_all.csv 读取“金融机构外汇信贷收支表”，
    提取：外币存款_境内/境外、外币贷款_境内/境外 的月度序列；
    计算合计与12期同比（%），返回列：
        日期, 外币存款_境内, 外币存款_境外, 外币存款_合计, 存款同比,
            外币贷款_境内, 外币贷款_境外, 外币贷款_合计, 贷款同比
    """
    import os
    import re
    import pandas as pd

    csv_path = os.path.join(os.getcwd(), 'pbc_fx_credit_all.csv')
    if not os.path.exists(csv_path):
        return None

    try:
        raw = pd.read_csv(csv_path, encoding='utf-8-sig')
    except Exception:
        # 尝试无 BOM
        raw = pd.read_csv(csv_path, encoding='utf-8', errors='ignore')

    # 识别月份列（yyyy-mm）
    month_cols = [c for c in raw.columns if re.fullmatch(r'20\d{2}-\d{2}', str(c).strip())]
    if not month_cols:
        return None
    month_cols = sorted(month_cols)

    def _norm_label(s: str) -> str:
        s = str(s or '')
        s = s.replace('\u3000',' ').replace('\xa0',' ').strip()
        s = s.replace('（','(').replace('）',')')
        return s

    raw['项目'] = raw.get('项目', '').map(_norm_label)
    raw['Item'] = raw.get('Item', '').map(_norm_label)

    def _find_row(keyword_pairs):
        """keyword_pairs: list of (zh_kw_regex, en_kw_regex)"""
        for _, r in raw.iterrows():
            zh = _norm_label(r.get('项目', ''))
            en = _norm_label(r.get('Item', ''))
            for zke, eke in keyword_pairs:
                z_ok = bool(re.search(zke, zh)) if zke else False
                e_ok = bool(re.search(eke, en, flags=re.IGNORECASE)) if eke else False
                if z_ok or e_ok:
                    return r
        return None

    # 关键行匹配（尽量用中文，英文兜底）
    # 存款：按要求，境内=住户存款 + 非金融企业存款；境外=表内“境外.*存款”
    row_dep_hh = _find_row([
        (r'住户.*存款', r'Household.*Deposits'),
    ])
    row_dep_nfe = _find_row([
        (r'非金融.*企业.*存款|非金融企业.*存款', r'Non.?financial.*Enterprises?.*Deposits'),
    ])
    row_dep_dom_legacy = _find_row([
        (r'境内.*存款', r'Domestic.*Deposits'),
        (r'国内.*存款', r''),
    ])
    row_dep_for = _find_row([
        (r'境外.*存款', r'Overseas.*Deposits'),
    ])
    row_loan_dom = _find_row([
        (r'境内.*贷款', r'Domestic.*Loans?'),
        (r'国内.*贷款', r''),
    ])
    row_loan_for = _find_row([
        (r'境外.*贷款', r'Overseas.*Loans?'),
    ])

    # 构造时间序列
    def _series_from_row(row):
        if row is None:
            return pd.Series(dtype=float)
        s = pd.to_numeric([row.get(c) for c in month_cols], errors='coerce')
        ts = pd.Series(s, index=pd.to_datetime(month_cols, format='%Y-%m'))
        return ts

    # 计算境内=住户 + 非金融企业；若二者缺失则退回到“境内存款”行
    s_hh = _series_from_row(row_dep_hh)
    s_nfe = _series_from_row(row_dep_nfe)
    dep_dom_candidate = s_hh.add(s_nfe, fill_value=0)
    dep_dom_legacy = _series_from_row(row_dep_dom_legacy)
    dep_dom = dep_dom_candidate if not dep_dom_candidate.empty else dep_dom_legacy
    dep_for = _series_from_row(row_dep_for)
    loan_dom = _series_from_row(row_loan_dom)
    loan_for = _series_from_row(row_loan_for)

    # 对齐索引
    idx = dep_dom.index.union(dep_for.index).union(loan_dom.index).union(loan_for.index)
    idx = idx.sort_values()
    dep_dom = dep_dom.reindex(idx)
    dep_for = dep_for.reindex(idx)
    loan_dom = loan_dom.reindex(idx)
    loan_for = loan_for.reindex(idx)

    df = pd.DataFrame({
        '日期': idx,
        '外币存款_境内': dep_dom.values,
        '外币存款_境外': dep_for.values,
        '外币贷款_境内': loan_dom.values,
        '外币贷款_境外': loan_for.values,
    })

    # 合计与同比（合计的12个月同比，单位为百分比）
    df['外币存款_合计'] = pd.to_numeric(df['外币存款_境内'], errors='coerce').fillna(0) + pd.to_numeric(df['外币存款_境外'], errors='coerce').fillna(0)
    df['外币贷款_合计'] = pd.to_numeric(df['外币贷款_境内'], errors='coerce').fillna(0) + pd.to_numeric(df['外币贷款_境外'], errors='coerce').fillna(0)

    # 可选：计算同比，但在图表中不直接显示折线
    try:
        df['存款同比'] = (df['外币存款_合计'].replace(0, pd.NA).pct_change(12) * 100).round(2)
        df['贷款同比'] = (df['外币贷款_合计'].replace(0, pd.NA).pct_change(12) * 100).round(2)
        first_yoy_date = pd.to_datetime(df['日期']).min() + pd.DateOffset(months=12)
        df.loc[pd.to_datetime(df['日期']) < first_yoy_date, ['存款同比','贷款同比']] = pd.NA
    except Exception:
        df['存款同比'] = pd.NA
        df['贷款同比'] = pd.NA

    return df

# 简洁的两序列堆积柱图（与整体风格一致）
def build_stack_bar_two_series(x, s1, s2, title, name1="境内", name2="境外", color1=PALETTE[2], color2=PALETTE[5]):
    bar = Bar(init_opts=opts.InitOpts(bg_color="#FFFFFF"))
    bar.add_xaxis(x)
    bar.add_yaxis(name1, s1, stack="sum", category_gap="30%",
                  itemstyle_opts=opts.ItemStyleOpts(color=color1, opacity=0.85),
                  label_opts=opts.LabelOpts(is_show=False), z=1)
    bar.add_yaxis(name2, s2, stack="sum",
                  itemstyle_opts=opts.ItemStyleOpts(color=color2, opacity=0.85),
                  label_opts=opts.LabelOpts(is_show=False), z=1)
    bar.set_global_opts(
        title_opts=opts.TitleOpts(title=title, pos_left="center", title_textstyle_opts=opts.TextStyleOpts(color="#111827")),
        tooltip_opts=opts.TooltipOpts(trigger="axis", axis_pointer_type="shadow"),
        xaxis_opts=opts.AxisOpts(type_="category", boundary_gap=True),
        yaxis_opts=opts.AxisOpts(name="金额"),
        legend_opts=opts.LegendOpts(orient="horizontal", pos_top="40"),
        datazoom_opts=[opts.DataZoomOpts(type_="inside"), opts.DataZoomOpts(type_="slider")],
        toolbox_opts=opts.ToolboxOpts(is_show=True, feature=opts.ToolBoxFeatureOpts(
            save_as_image=opts.ToolBoxFeatureSaveAsImageOpts(),
            data_view=opts.ToolBoxFeatureDataViewOpts(is_show=True),
            data_zoom=opts.ToolBoxFeatureDataZoomOpts(),
            restore=opts.ToolBoxFeatureRestoreOpts(),
        )),
    )
    return bar

# 堆积柱 + 同比折线（右轴）
def build_stack_bar_two_series_with_yoy(x, s1, s2, yoy_pct, title, name1="境内", name2="境外", color1=PALETTE[2], color2=PALETTE[5], line_color=PALETTE[6]):
    from math import floor, ceil
    xs = [v for v in (yoy_pct or []) if v is not None and not pd.isna(v)]
    if xs:
        vmin, vmax = min(xs), max(xs)
        if vmin == vmax:
            vmin -= 5; vmax += 5
        span = vmax - vmin
        vmin -= span * 0.1; vmax += span * 0.1
        nice_min = floor(vmin / 5.0) * 5
        nice_max = ceil(vmax / 5.0) * 5
        nice_min = min(nice_min, 0)
        nice_max = max(nice_max, 0)
        if nice_max - nice_min < 10:
            nice_max = nice_min + 10
    else:
        nice_min, nice_max = None, None

    bar = Bar(init_opts=opts.InitOpts(bg_color="#FFFFFF"))
    bar.add_xaxis(x)
    bar.add_yaxis(name1, s1, stack="sum", category_gap="30%",
                  itemstyle_opts=opts.ItemStyleOpts(color=color1, opacity=0.85),
                  label_opts=opts.LabelOpts(is_show=False), z=1)
    bar.add_yaxis(name2, s2, stack="sum",
                  itemstyle_opts=opts.ItemStyleOpts(color=color2, opacity=0.85),
                  label_opts=opts.LabelOpts(is_show=False), z=1)
    bar.extend_axis(yaxis=opts.AxisOpts(name="同比", type_="value", position="right",
                                        axislabel_opts=opts.LabelOpts(formatter="{value}%"),
                                        min_=nice_min, max_=nice_max))
    bar.set_global_opts(
        title_opts=opts.TitleOpts(title=title, pos_left="center", title_textstyle_opts=opts.TextStyleOpts(color="#111827")),
        tooltip_opts=opts.TooltipOpts(trigger="axis", axis_pointer_type="shadow"),
        xaxis_opts=opts.AxisOpts(type_="category", boundary_gap=True),
        yaxis_opts=opts.AxisOpts(name="金额"),
        legend_opts=opts.LegendOpts(orient="horizontal", pos_top="40"),
        datazoom_opts=[opts.DataZoomOpts(type_="inside"), opts.DataZoomOpts(type_="slider")],
        toolbox_opts=opts.ToolboxOpts(is_show=True, feature=opts.ToolBoxFeatureOpts(
            save_as_image=opts.ToolBoxFeatureSaveAsImageOpts(),
            data_view=opts.ToolBoxFeatureDataViewOpts(is_show=True),
            data_zoom=opts.ToolBoxFeatureDataZoomOpts(),
            restore=opts.ToolBoxFeatureRestoreOpts(),
        )),
    )
    line = Line()
    line.add_xaxis(x)
    line.add_yaxis(
        "同比",
        [None if pd.isna(v) else v for v in yoy_pct],
        yaxis_index=1,
        is_smooth=False,
        is_symbol_show=True,
        symbol="circle",
        symbol_size=7,
        label_opts=opts.LabelOpts(is_show=False),  # 隐藏数据标签，仅显示折线与节点
        linestyle_opts=opts.LineStyleOpts(width=2.2, color=line_color),
        itemstyle_opts=opts.ItemStyleOpts(color=line_color),
        z=100,
    )
    return bar.overlap(line)
def build_bar_line_dual_axis(x_list, amt_list, yoy_pct_list, amt_name, title, bar_color=None, line_color=None):
    """
    左轴金额(柱)，右轴同比%(线)。莫兰迪配色更柔和，与背景协调；折线与柱状色彩区分明显。
    - 柱：PALETTE[0]，略透明；
    - 线：PALETTE[1]；
    - 折线置顶 z=10，避免遮挡；
    - 同比轴“好看”的上下界（5%步长，含0）。
    """
    from pyecharts import options as opts
    from pyecharts.charts import Bar, Line

    # 计算同比轴“好看”的上下界
    def _nice_percent_axis(vals):
        xs = [v for v in vals if v is not None]
        if not xs:
            return None, None
        vmin, vmax = min(xs), max(xs)
        if vmin == vmax:
            vmin -= 5
            vmax += 5
        span = vmax - vmin
        vmin -= span * 0.1
        vmax += span * 0.1
        import math
        nice_min = math.floor(vmin / 5.0) * 5
        nice_max = math.ceil(vmax / 5.0) * 5
        nice_min = min(nice_min, 0)
        nice_max = max(nice_max, 0)
        if nice_max - nice_min < 10:
            nice_max = nice_min + 10
        return nice_min, nice_max

    nice_min, nice_max = _nice_percent_axis(yoy_pct_list)
    bar_color = bar_color or PALETTE[5]
    line_color = line_color or PALETTE[6]
    bar = Bar(init_opts=opts.InitOpts(bg_color="#FFFFFF", width="100%"))
    bar.add_xaxis(x_list)
    bar.add_yaxis(
        series_name=amt_name,
        y_axis=amt_list,
        label_opts=opts.LabelOpts(is_show=False),
        category_gap="35%",
        # 柱子顶部改为平顶（无圆角）
        itemstyle_opts=opts.ItemStyleOpts(color=bar_color, opacity=0.82, border_radius=[0, 0, 0, 0]),
        z=1,
    )
    # 右侧第二坐标轴：同比（%）
    bar.extend_axis(
        yaxis=opts.AxisOpts(
            name="同比",
            type_="value",
            position="right",
            axislabel_opts=opts.LabelOpts(formatter="{value}%"),
            min_=nice_min,
            max_=nice_max,
        )
    )
    bar.set_global_opts(
        title_opts=opts.TitleOpts(
            title=title,
            pos_left="center",
            title_textstyle_opts=opts.TextStyleOpts(color="#111827"),
        ),
        tooltip_opts=opts.TooltipOpts(trigger="axis"),
        toolbox_opts=opts.ToolboxOpts(is_show=True),
        xaxis_opts=opts.AxisOpts(type_="category", boundary_gap=True),
        yaxis_opts=opts.AxisOpts(name="金额"),
        legend_opts=opts.LegendOpts(orient="horizontal", pos_top="40"),
        datazoom_opts=[
            opts.DataZoomOpts(type_="inside"),
            opts.DataZoomOpts(type_="slider"),
        ],
    )
    # 折线（右轴）：莫兰迪的柔和橄榄绿，层级更高
    line = Line()
    line.add_xaxis(x_list)
    line.add_yaxis(
        series_name="同比",
        y_axis=yoy_pct_list,
        yaxis_index=1,
        label_opts=opts.LabelOpts(is_show=False),
        is_smooth=False,
        is_symbol_show=True,
        symbol_size=6,
        z=10,
        linestyle_opts=opts.LineStyleOpts(width=2.2, color=line_color),
        itemstyle_opts=opts.ItemStyleOpts(color=line_color),
    )
    chart = bar.overlap(line)
    return chart




st.title("跨境数据看板")
st.markdown("")


data = load_data()
latest_info = ""
if data and "全国" in data and not data["全国"].empty:
    try:
        latest_month = pd.to_datetime(data["全国"]["时间"]).max()
        latest_info = f"全国数据更新至 {latest_month.strftime('%Y-%m')}"
    except Exception:
        pass
if data and "杭州市" in data and not data["杭州市"].empty:
    try:
        latest_month_zhejiang = pd.to_datetime(data["杭州市"]["时间"]).max()
        latest_info_zhejiang = f"浙江省数据更新至 {latest_month_zhejiang.strftime('%Y-%m')}"
    except Exception:
        pass



with st.sidebar:
    st.header("页面选择")
    # 先选分类，再选页面
    _categories = ["海关数据", "银行数据", "宏观数据"]
    _category_labels = {
        "海关数据": "🧭 海关数据",
        "银行数据": "🏦 银行数据",
        "宏观数据": "🌐 宏观数据",
    }
    category = st.radio(
        "选择分类",
        _categories,
        index=0,
        horizontal=False,
        format_func=lambda x: _category_labels.get(x, x),
        key="page_category",
    )

    _group_pages = {
        "海关数据": ["海关综合看板", "海关产品类别看板"],
        "银行数据": ["机构外币存贷款看板", "银行结售汇"],
        "宏观数据": ["汇率数据", "利率数据", "CME FEDWATCH"],
    }
    _page_options = _group_pages.get(category, [])
    _page_labels = {
        "海关综合看板": "📊 海关综合看板",
        "海关产品类别看板": "📦 海关产品类别看板",
        "机构外币存贷款看板": "💱 机构外币存贷款看板",
        "银行结售汇": "🏦 银行结售汇",
        "汇率数据": "💹 汇率数据",
        "利率数据": "📉 利率数据",
        "CME FEDWATCH": "📈 CME FedWatch",
    }
    page = st.radio(
        "选择页面",
        _page_options,
        index=0,
        horizontal=False,
        format_func=lambda x: _page_labels.get(x, x),
        key="page_select",
    )
    st.markdown("---")

    if page == "海关综合看板":
        st.header("操作面板")
        if latest_info:
            st.caption(latest_info)

        if latest_info_zhejiang:
            st.caption(latest_info_zhejiang)


        st.header("地区筛选")
        default_index = FINAL_LOCATIONS.index("浙江省") if "浙江省" in FINAL_LOCATIONS else 0
        selected_location = st.selectbox("选择地区", options=FINAL_LOCATIONS, index=default_index, label_visibility="collapsed")

        st.header("展示设置")
        show_overview = st.checkbox("显示全国与重点地区概览", value=True)


        st.markdown("---")
        st.subheader("数据更新")
        if st.button("更新海关统计数据", type="primary", use_container_width=True, key="btn_update_customs"):
            with st.spinner("正在更新海关统计数据，请稍候..."):
                res = run_data_updater()
            if res.get("success"):
                st.success(res.get("message", "更新完成"))
                if res.get("output"):
                    st.text_area("输出", res.get("output", ""), height=160)
                try:
                    st.rerun()
                except Exception:
                    try:
                        st.experimental_rerun()
                    except Exception:
                        pass
            else:
                st.error(res.get("message", "更新失败"))

    elif page == "海关产品类别看板":  # 产品类别看板
        st.header("数据控制面板")


        if "cat_filters_visible" not in st.session_state:
            st.session_state["cat_filters_visible"] = True

        def hide_cat_filters():
            st.session_state["cat_filters_visible"] = False

        def show_cat_filters():
            st.session_state["cat_filters_visible"] = True

        category_data_sidebar = load_category_data2()
        all_regions_sidebar = []
        if category_data_sidebar:
            sample_sheet = list(category_data_sidebar.keys())[0]
            all_regions_sidebar = category_data_sidebar[sample_sheet].index.tolist()

        if st.session_state.get("cat_filters_visible", True):
            st.subheader("币种")
            st.radio(
                "选择币种：",
                options=["美元", "人民币"],
                index=0,
                horizontal=True,
                key="cat_currency",
            )

            st.subheader("地区筛选")
            default_regions = all_regions_sidebar[:6] if len(all_regions_sidebar) >= 6 else all_regions_sidebar
            st.multiselect(
                "选择要显示的地区：",
                options=all_regions_sidebar,
                default=default_regions,
                key="cat_selected_regions",
                help="最多选择8个地区以保证图表清晰度",
            )

            if st.session_state.get("cat_selected_regions") and len(st.session_state["cat_selected_regions"]) > 8:
                st.warning("为保证图表清晰度，建议最多选择8个地区")
                st.session_state["cat_selected_regions"] = st.session_state["cat_selected_regions"][:8]

            st.markdown("---")
            st.subheader("产品类别")
            st.multiselect(
                "选择要显示的产品类别:",
                options=CATEGORY_LABELS,
                default=CATEGORY_LABELS,
                key="cat_selected_categories",
                help="可以选择部分类别进行对比分析",
            )

            st.markdown("---")
            st.button("应用并隐藏筛选", type="primary", use_container_width=True, on_click=hide_cat_filters)
        else:

            cur = st.session_state.get("cat_currency", "美元")
            regs = st.session_state.get("cat_selected_regions", [])
            cats = st.session_state.get("cat_selected_categories", CATEGORY_LABELS)

            def summarize(items):
                if not items:
                    return "未选择"
                if len(items) <= 3:
                    return "、".join(items)
                return "、".join(items[:3]) + f" 等{len(items)}项"

            st.caption(f"币种：{cur}")
            st.caption(f"地区：{len(regs)} 项（{summarize(regs)}）")
            st.caption(f"产品类别：{len(cats)} 项（{summarize(cats)}）")
            st.button("编辑筛选", use_container_width=True, on_click=show_cat_filters)

        st.caption("数据来源：海关总署")

# ——— FedWatch：侧边栏更新按钮（全局，在页面选择下方） ———
with st.sidebar:
    try:
        if st.session_state.get("page_select") == "CME FEDWATCH":
            st.markdown("---")
            st.subheader("FedWatch 数据更新")
            st.caption("一键从 CME FedWatch 抓取并更新数据（需联网，可能需要代理）")
            proxy = st.text_input("代理（可选）", value=os.environ.get("PLAYWRIGHT_PROXY", ""), key="fedwatch_proxy")
            if st.button("更新 FedWatch 数据", type="primary", use_container_width=True, key="btn_update_fedwatch"):
                with st.spinner("正在抓取并更新 FedWatch 数据，请稍候..."):
                    import subprocess as _sp
                    try:
                        env = os.environ.copy()
                        if proxy:
                            env["PLAYWRIGHT_PROXY"] = proxy
                        # 直接抓取（Cloud 安装逻辑已移除，部署到 Cloud Run）
                        result = _sp.run(
                            [sys.executable, os.path.join("scripts", "cme_fedwatch_scrape.py")],
                            capture_output=True,
                            text=True,
                            encoding="utf-8",
                            errors="ignore",
                            timeout=600,
                            env=env,
                        )
                        if result.returncode == 0:
                            st.success("FedWatch 数据更新完成，已写入 outputs/ 目录！")
                            if result.stdout:
                                st.text_area("输出", result.stdout, height=160)
                        else:
                            st.error("更新失败")
                            st.text_area("错误输出", result.stderr or result.stdout, height=180)
                    except _sp.TimeoutExpired:
                        st.error("更新超时（超过10分钟）。")
                    except Exception as e:
                        st.error(f"更新过程中出错：{e}")
    except Exception:
        pass




if page == "海关综合看板":
    if not data:
        st.info("未检测到数据文件，请先运行数据处理脚本生成 Excel。")
        st.stop()


    if show_overview:
        nat = data.get("全国")
        if nat is not None and not nat.empty:
            latest_row = nat.iloc[nat["时间"].map(pd.to_datetime).idxmax()]
            st.markdown("### 全国(单位:万元)")
            cols = st.columns(3)
            metrics = [
                ("进出口 (年初至今)", latest_row.get("进出口_年初至今"), latest_row.get("进出口_年初至今同比")),
                ("进口 (年初至今)", latest_row.get("进口_年初至今"), latest_row.get("进口_年初至今同比")),
                ("出口 (年初至今)", latest_row.get("出口_年初至今"), latest_row.get("出口_年初至今同比")),
            ]
            for c, (label, val, yoy) in zip(cols, metrics):
                c.markdown(render_card(label, val, yoy), unsafe_allow_html=True)

    st.markdown("<hr/>", unsafe_allow_html=True)


    top_regions = ["北京市", "上海市", "深圳市", "南京市", "合肥市", "浙江省"]
    if show_overview:
        st.subheader("重点地区数据概览(单位:万元)")
        for loc in top_regions:
            df = data.get(loc)
            if df is None or df.empty:
                continue
            latest = df.iloc[df["时间"].map(pd.to_datetime).idxmax()]
            st.markdown(f"### {loc}")
            cards = st.columns(3)
            metrics = [
                ("进出口(年初至今)", latest.get("进出口_年初至今"), latest.get("进出口_年初至今同比")),
                ("进口(年初至今)", latest.get("进口_年初至今"), latest.get("进口_年初至今同比")),
                ("出口(年初至今)", latest.get("出口_年初至今"), latest.get("出口_年初至今同比")),
            ]
            for c, (label, val, yoy) in zip(cards, metrics):
                c.markdown(render_card(label, val, yoy), unsafe_allow_html=True)

            if loc == "浙江省":
                ZHEJIANG_CITIES = ["杭州市", "宁波市", "温州市", "湖州市", "金华市", "台州市", "嘉兴市", "丽水市", "衢州市", "绍兴市", "舟山市"]
                with st.expander("展开/收起 浙江省地市数据概览", expanded=False):
                    shown_any = False
                    for city in ZHEJIANG_CITIES:
                        city_df = data.get(city)
                        if city_df is None or city_df.empty:
                            continue
                        latest_city = city_df.iloc[city_df["时间"].map(pd.to_datetime).idxmax()]
                        month_str = pd.to_datetime(latest_city.get("时间")).strftime("%Y-%m") if pd.notna(latest_city.get("时间")) else ""
                        st.markdown(f"**{city}（{month_str}）**")
                        cc1, cc2, cc3 = st.columns(3)
                        cc1.markdown(render_card("进出口(年初至今)", latest_city.get("进出口_年初至今"), latest_city.get("进出口_年初至今同比")), unsafe_allow_html=True)
                        cc2.markdown(render_card("进口(年初至今)", latest_city.get("进口_年初至今"), latest_city.get("进口_年初至今同比")), unsafe_allow_html=True)
                        cc3.markdown(render_card("出口(年初至今)", latest_city.get("出口_年初至今"), latest_city.get("出口_年初至今同比")), unsafe_allow_html=True)
                        st.markdown("")
                        shown_any = True
                    if not shown_any:
                        st.info("未找到可展示的浙江省地市数据")


    st.markdown("<hr/>", unsafe_allow_html=True)


    st.subheader(f"{selected_location} · 详细数据与走势")
    loc_df = data.get(selected_location)
    if loc_df is None or loc_df.empty:
        st.warning(f"未找到 {selected_location} 的数据")
    else:
        loc_df_sorted = loc_df.copy()
        loc_df_sorted["时间"] = pd.to_datetime(loc_df_sorted["时间"])
        loc_df_sorted.sort_values("时间", inplace=True)
        x_axis = loc_df_sorted["时间"].dt.strftime("%Y-%m").tolist()

        tab1, tab2 = st.tabs(["当月走势", "详细数据表"])
        with tab1:
            series = {
                "进出口(当月)": loc_df_sorted["进出口_当月"].tolist(),
                "进口(当月)": loc_df_sorted["进口_当月"].tolist(),
                "出口(当月)": loc_df_sorted["出口_当月"].tolist(),
            }
            chart = build_line(x_axis, series, f"{selected_location} · 当月数据走势")
            st.markdown('<div class="chart-rounded">', unsafe_allow_html=True)
            st_pyecharts(chart, height="520px")
            st.markdown('</div>', unsafe_allow_html=True)

        with tab2:
            disp = loc_df_sorted.copy()
            # 百分比列格式化显示
            for col in list(disp.columns):
                if "同比" in col:
                    try:
                        disp[col] = disp[col].apply(lambda v: f"{float(v):.2%}" if pd.notna(v) else "—")
                    except Exception:
                        pass
            disp["时间"] = disp["时间"].dt.strftime("%Y-%m")
            st.dataframe(disp, use_container_width=True)

elif page == "海关产品类别看板":
    category_data = load_category_data2()
    if not category_data:
        st.error("未找到数据文件 '9大类产品分析表.xlsx'，请先运行数据处理脚本生成Excel文件。")
        st.stop()


    currency = st.session_state.get("cat_currency", "美元")
    all_regions_for_page = []
    try:
        sample_sheet = list(category_data.keys())[0]
        all_regions_for_page = category_data[sample_sheet].index.tolist()
    except Exception:
        pass
    default_regions_page = all_regions_for_page[:6] if len(all_regions_for_page) >= 6 else all_regions_for_page
    selected_regions = st.session_state.get("cat_selected_regions", default_regions_page)
    if selected_regions and len(selected_regions) > 8:
        selected_regions = selected_regions[:8]
    selected_categories = st.session_state.get("cat_selected_categories", CATEGORY_LABELS)

    st.markdown("## 主要产品类别贸易结构")
    chart_configs = [
        (f"进出口_{currency}", f"进出口贸易结构（{currency}）"),
        (f"出口_{currency}", f"出口贸易结构（{currency}）"),
        (f"进口_{currency}", f"进口贸易结构（{currency}）"),
    ]

    if selected_regions and selected_categories:
        for sheet_name, chart_title in chart_configs:
            if sheet_name in category_data:
                st.markdown(f"### {chart_title}")
                available_regions = [r for r in selected_regions if r in category_data[sheet_name].index]
                available_categories = [c for c in selected_categories if c in category_data[sheet_name].columns]
                if not available_regions or not available_categories:
                    st.warning("该数据表中缺少选中的地区或产品类别")
                    continue

                chart_df = category_data[sheet_name].loc[available_regions, available_categories].fillna(0)
                chart = create_horizontal_percentage_chart(
                    data=chart_df,
                    title=chart_title,
                    regions=available_regions,
                    categories=available_categories,
                )
                chart_height = max(400, len(available_regions) * 60 + 150)
                show_chart(chart, height=f"{chart_height}px")


                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**原始金额**")
                    display_data = chart_df.copy()
                    for col in display_data.columns:
                        display_data[col] = display_data[col].apply(lambda x: f"{x:,.0f}" if pd.notna(x) and x != 0 else "—")
                    st.dataframe(
                        display_data,
                        use_container_width=True,
                        height=min(len(available_regions) * 35 + 50, 300),
                    )
                with col2:
                    st.markdown("**占比（%）**")
                    row_sums = chart_df.sum(axis=1).replace(0, 1)
                    percentage_data = chart_df.div(row_sums, axis=0) * 100
                    for col in percentage_data.columns:
                        percentage_data[col] = percentage_data[col].apply(lambda x: f"{x:.1f}%" if pd.notna(x) and x != 0 else "—")
                    st.dataframe(
                        percentage_data,
                        use_container_width=True,
                        height=min(len(available_regions) * 35 + 50, 300),
                    )

elif page == "机构外币存贷款看板":
    st.subheader("金融机构外币存贷款(亿美元)")
    df = load_fx_deposit_loan()
    if df is None or df.empty:
        st.error("未找到外币存贷款数据：请将 含‘外币’‘存贷款’关键字的xlsx 放在当前目录或数据目录（支持2023+任意年份），并保证含有‘境内/境外/住户/非金融企业’等行名。")
    else:
        df["月份"] = pd.to_datetime(df["日期"]).dt.strftime("%Y-%m")
        x_axis = df["月份"].tolist()

        # 年初至今卡片（外币存款/外币贷款）
        dates = pd.to_datetime(df["日期"], errors="coerce")
        def _num_col(series_or_values):
            s = pd.Series(series_or_values)
            s = s.astype(str).str.replace(',', '', regex=False).str.replace(' ', '', regex=False)
            return pd.to_numeric(s, errors="coerce")
        dep_total_col = df.get("外币存款_合计")
        if dep_total_col is None:
            dep_total_col = _num_col(df.get("外币存款_境内", 0)).fillna(0) + _num_col(df.get("外币存款_境外", 0)).fillna(0)
        else:
            dep_total_col = _num_col(dep_total_col)
        loan_total_col = df.get("外币贷款_合计")
        if loan_total_col is None:
            loan_total_col = _num_col(df.get("外币贷款_境内", 0)).fillna(0) + _num_col(df.get("外币贷款_境外", 0)).fillna(0)
        else:
            loan_total_col = _num_col(loan_total_col)
        # 使用按年分组的方式计算 YTD 与同比（避免索引/频率问题）
        if dates.notna().any():
            year_now = int(dates.max().year)
        else:
            year_now = None
        def _ytd_sum_and_yoy(vals):
            if year_now is None:
                return None, None
            v = pd.Series(vals).astype(float)
            # 以当前年最新数据的“月份”作为YTD截止月，并按同期月份对比上一年
            mask_cur_year = dates.dt.year == year_now
            if not mask_cur_year.any():
                return None, None
            last_month = int(pd.to_datetime(dates[mask_cur_year]).max().month)
            mask_cur = mask_cur_year & (dates.dt.month <= last_month)
            cur_sum = float(v[mask_cur].dropna().sum()) if mask_cur.any() else None
            # 上年同期（截止到同一个月份）
            mask_prev_year = dates.dt.year == (year_now - 1)
            mask_prev = mask_prev_year & (dates.dt.month <= last_month)
            prev_vals = v[mask_prev].dropna()
            if len(prev_vals):
                prev_sum = float(prev_vals.sum())
                yoy = (cur_sum / prev_sum - 1.0) if prev_sum != 0 and cur_sum is not None else None
            else:
                yoy = None
            return cur_sum, yoy
        ytd_dep, ytd_dep_yoy = _ytd_sum_and_yoy(dep_total_col)
        ytd_loan, ytd_loan_yoy = _ytd_sum_and_yoy(loan_total_col)
        c1, c2 = st.columns(2)
        c1.markdown(render_card("外币存款（年初至今）", ytd_dep, ytd_dep_yoy), unsafe_allow_html=True)
        c2.markdown(render_card("外币贷款（年初至今）", ytd_loan, ytd_loan_yoy), unsafe_allow_html=True)

        # 外币存款：境内/境外 堆叠 + 合计同比折线（隐藏数据标签）
        dep_dom = df.get("外币存款_境内", pd.Series([0]*len(df))).round(2).tolist()
        dep_for = df.get("外币存款_境外", pd.Series([0]*len(df))).round(2).tolist()
        dep_yoy = df.get("存款同比", pd.Series([None]*len(df)))
        dep_yoy = dep_yoy.where(dep_yoy.notna(), None).tolist()
        dep_chart = build_stack_bar_two_series_with_yoy(
            x_axis, dep_dom, dep_for, dep_yoy, "外币存款（境内/境外）"
        )
        show_chart(dep_chart, height="520px")

        # 外币贷款：境内/境外 堆叠 + 合计同比折线（隐藏数据标签）
        loan_dom = df.get("外币贷款_境内", pd.Series([0]*len(df))).round(2).tolist()
        loan_for = df.get("外币贷款_境外", pd.Series([0]*len(df))).round(2).tolist()
        loan_yoy = df.get("贷款同比", pd.Series([None]*len(df)))
        loan_yoy = loan_yoy.where(loan_yoy.notna(), None).tolist()
        loan_chart = build_stack_bar_two_series_with_yoy(
            x_axis, loan_dom, loan_for, loan_yoy, "外币贷款（境内/境外）"
        )
        show_chart(loan_chart, height="520px")

        st.markdown("<hr/>", unsafe_allow_html=True)
        st.subheader("明细数据")
        cols = [
            "月份",
            "外币存款_境内","外币存款_境外","外币存款_合计",
            "外币贷款_境内","外币贷款_境外","外币贷款_合计",
        ]
        table = df[cols].copy()
        st.dataframe(table.sort_values("月份", ascending=False), use_container_width=True, hide_index=True)

elif page == "银行结售汇":
    # 顶部标题与月份窗口（全宽，无额外阴影样式）
    st.subheader("银行结售汇(亿人民币)")
    months = st.slider("显示窗口（月）", min_value=12, max_value=120, value=36, step=6, key="bank_fx_months")

    # 即时导入（不再延迟），并启用工具栏以下载图表
    from pyecharts.commons.utils import JsCode
    from bank_fx_data import get_dashboard_data, ytd_sum_and_yoy, gross_amount, load_bank_fx
    import math

    data_fx = get_dashboard_data(months=months)
    if data_fx is None or data_fx["main"].empty:
        st.error("未找到或无法解析 ‘银行结售汇数据时间序列.xlsx’ 的人民币月度表，请将文件放在当前运行目录")
        st.stop()

    main = data_fx["main"].copy()
    comp = data_fx["comp"].copy()
    fwd_sign = data_fx["fwd_sign"].copy()
    fwd_out = data_fx["fwd_out"].copy()

    last_month = main.index.max().strftime("%Y-%m")
    st.caption(f"更新至：{last_month}")

    # 辅助函数（仅本页使用）
    def _nice_percent_axis(vals):
        xs = [v for v in vals if (v is not None and not pd.isna(v))]
        if not xs:
            return None, None
        vmin, vmax = min(xs), max(xs)
        if vmin == vmax:
            vmin -= 5
            vmax += 5
        span = vmax - vmin
        vmin -= span * 0.1
        vmax += span * 0.1
        nice_min = math.floor(vmin / 5.0) * 5
        nice_max = math.ceil(vmax / 5.0) * 5
        nice_min = min(nice_min, 0)
        nice_max = max(nice_max, 0)
        if nice_max - nice_min < 10:
            nice_max = nice_min + 10
        return nice_min, nice_max

    def build_stack_bar_plus_yoy(x, s1, s2, total_yoy_pct, title, name1, name2):
        yoy_axis_min, yoy_axis_max = _nice_percent_axis(total_yoy_pct)
        bar = Bar(init_opts=opts.InitOpts(bg_color="#FFFFFF"))
        bar.add_xaxis(x)
        bar.add_yaxis(
            name1,
            s1,
            stack="sum",
            category_gap="30%",
            itemstyle_opts=opts.ItemStyleOpts(color=PALETTE[2], opacity=0.85),
            label_opts=opts.LabelOpts(is_show=False),
            z=1,
        )
        bar.add_yaxis(
            name2,
            s2,
            stack="sum",
            itemstyle_opts=opts.ItemStyleOpts(color=PALETTE[5], opacity=0.85),
            label_opts=opts.LabelOpts(is_show=False),
            z=1,
        )
        bar.extend_axis(
            yaxis=opts.AxisOpts(
                name="同比",
                type_="value",
                position="right",
                axislabel_opts=opts.LabelOpts(
                    formatter=JsCode("function(v){return v.toFixed(2)+'%';}")
                ),
                min_=yoy_axis_min,
                max_=yoy_axis_max,
            )
        )
        bar.set_global_opts(
            title_opts=opts.TitleOpts(
                title=title,
                pos_left="center",
                title_textstyle_opts=opts.TextStyleOpts(color="#111827"),
            ),
            tooltip_opts=opts.TooltipOpts(trigger="axis"),
            xaxis_opts=opts.AxisOpts(type_="category", boundary_gap=True),
            yaxis_opts=opts.AxisOpts(name="金额（亿元）"),
            legend_opts=opts.LegendOpts(orient="horizontal", pos_top="40"),
            datazoom_opts=[opts.DataZoomOpts(type_="inside"), opts.DataZoomOpts(type_="slider")],
            toolbox_opts=opts.ToolboxOpts(is_show=True, feature=opts.ToolBoxFeatureOpts(
                save_as_image=opts.ToolBoxFeatureSaveAsImageOpts(),
                data_view=opts.ToolBoxFeatureDataViewOpts(is_show=True),
                data_zoom=opts.ToolBoxFeatureDataZoomOpts(),
                restore=opts.ToolBoxFeatureRestoreOpts(),
            )),
        )
        line = Line()
        line.add_xaxis(x)
        line.add_yaxis(
            "同比",
            [None if pd.isna(v) else v for v in total_yoy_pct],
            yaxis_index=1,
            is_smooth=False,
            label_opts=opts.LabelOpts(is_show=False),
            is_symbol_show=True,
            symbol="circle",
            symbol_size=7,
            is_connect_nones=True,
            linestyle_opts=opts.LineStyleOpts(width=2.2, color=PALETTE[6]),
            itemstyle_opts=opts.ItemStyleOpts(color=PALETTE[6]),
            z=100,
        )
        return bar.overlap(line)

    def build_bar_plus_yoy(x, amt, yoy_pct, title, bar_color=PALETTE[5], line_color=PALETTE[6]):
        yoy_axis_min, yoy_axis_max = _nice_percent_axis(yoy_pct)
        bar = Bar(init_opts=opts.InitOpts(bg_color="#FFFFFF"))
        bar.add_xaxis(x)
        bar.add_yaxis(
            "金额",
            amt,
            category_gap="30%",
            label_opts=opts.LabelOpts(is_show=False),
            itemstyle_opts=opts.ItemStyleOpts(color=bar_color, opacity=0.85),
            z=1,
        )
        bar.extend_axis(
            yaxis=opts.AxisOpts(
                name="同比",
                type_="value",
                position="right",
                axislabel_opts=opts.LabelOpts(
                    formatter=JsCode("function(v){return v.toFixed(2)+'%';}")
                ),
                min_=yoy_axis_min,
                max_=yoy_axis_max,
            )
        )
        bar.set_global_opts(
            title_opts=opts.TitleOpts(
                title=title,
                pos_left="center",
                title_textstyle_opts=opts.TextStyleOpts(color="#111827"),
            ),
            tooltip_opts=opts.TooltipOpts(trigger="axis"),
            xaxis_opts=opts.AxisOpts(type_="category", boundary_gap=True),
            yaxis_opts=opts.AxisOpts(name="金额（亿元）"),
            legend_opts=opts.LegendOpts(orient="horizontal", pos_top="40"),
            datazoom_opts=[opts.DataZoomOpts(type_="inside"), opts.DataZoomOpts(type_="slider")],
            toolbox_opts=opts.ToolboxOpts(is_show=True, feature=opts.ToolBoxFeatureOpts(
                save_as_image=opts.ToolBoxFeatureSaveAsImageOpts(),
                data_view=opts.ToolBoxFeatureDataViewOpts(is_show=True),
                data_zoom=opts.ToolBoxFeatureDataZoomOpts(),
                restore=opts.ToolBoxFeatureRestoreOpts(),
            )),
        )
        line = Line()
        line.add_xaxis(x)
        line.add_yaxis(
            "同比",
            [None if pd.isna(v) else v for v in yoy_pct],
            yaxis_index=1,
            is_smooth=False,
            label_opts=opts.LabelOpts(is_show=False),
            is_symbol_show=True,
            symbol="circle",
            symbol_size=7,
            is_connect_nones=True,
            linestyle_opts=opts.LineStyleOpts(width=2.2, color=line_color),
            itemstyle_opts=opts.ItemStyleOpts(color=line_color),
            z=100,
        )
        return bar.overlap(line)

    def build_multi_line(x, series_map: dict, title: str):
        line = Line(init_opts=opts.InitOpts(bg_color="#FFFFFF"))
        line.add_xaxis(x)
        line.set_colors(PALETTE)
        for name, y in series_map.items():
            line.add_yaxis(
                name,
                y,
                is_smooth=False,
                is_connect_nones=True,
                label_opts=opts.LabelOpts(is_show=False),
                is_symbol_show=True,
                symbol="circle",
                symbol_size=7,
                linestyle_opts=opts.LineStyleOpts(width=2),
                z=50,
            )
        line.set_global_opts(
            title_opts=opts.TitleOpts(
                title=title,
                pos_left="center",
                title_textstyle_opts=opts.TextStyleOpts(color="#111827"),
            ),
            tooltip_opts=opts.TooltipOpts(trigger="axis"),
            xaxis_opts=opts.AxisOpts(type_="category", boundary_gap=False),
            yaxis_opts=opts.AxisOpts(name="金额（亿元）"),
            legend_opts=opts.LegendOpts(orient="horizontal", pos_top="40"),
            datazoom_opts=[opts.DataZoomOpts(type_="inside"), opts.DataZoomOpts(type_="slider")],
            toolbox_opts=opts.ToolboxOpts(is_show=True, feature=opts.ToolBoxFeatureOpts(
                save_as_image=opts.ToolBoxFeatureSaveAsImageOpts(),
                data_view=opts.ToolBoxFeatureDataViewOpts(is_show=True),
                data_zoom=opts.ToolBoxFeatureDataZoomOpts(),
                restore=opts.ToolBoxFeatureRestoreOpts(),
            )),
        )
        return line

    # YTD 指标卡：使用全量数据计算，避免窗口切换导致YTD被截断
    _full = load_bank_fx()
    _src_for_ytd = _full.main if _full and _full.main is not None and not _full.main.empty else main
    ytd_settle, ytd_settle_yoy, _ = ytd_sum_and_yoy(_src_for_ytd["结汇"])
    ytd_sale, ytd_sale_yoy, _ = ytd_sum_and_yoy(_src_for_ytd["售汇"])
    gross_series = gross_amount(_src_for_ytd)
    ytd_gross, ytd_gross_yoy, _ = ytd_sum_and_yoy(gross_series)

    c1, c2, c3 = st.columns(3)
    c1.markdown(render_card("结汇（年初至今）", ytd_settle, ytd_settle_yoy), unsafe_allow_html=True)
    c2.markdown(render_card("售汇（年初至今）", ytd_sale, ytd_sale_yoy), unsafe_allow_html=True)
    c3.markdown(render_card("结售汇金额（年初至今）", ytd_gross, ytd_gross_yoy), unsafe_allow_html=True)


    # 图1：结汇 经常/资本堆叠 + 合计同比
    xs = main.index.strftime("%Y-%m").tolist()
    settle_yoy_pct = (main["结汇同比"] * 100).where(main["结汇同比"].notna(), None).tolist()
    settle_cur = comp.get("结汇_经常项目", pd.Series(index=main.index)).reindex(main.index).fillna(0).round(2).tolist()
    settle_cap = comp.get("结汇_资本项目", pd.Series(index=main.index)).reindex(main.index).fillna(0).round(2).tolist()
    chart1 = build_stack_bar_plus_yoy(
        xs,
        settle_cur,
        settle_cap,
        settle_yoy_pct,
        "结汇：经常项目 vs 资本项目(当月)",
        "经常项目",
        "资本项目",
    )
    show_chart(chart1, height="520px")

    # 图2：售汇 经常/资本堆叠 + 合计同比
    sale_yoy_pct = (main["售汇同比"] * 100).where(main["售汇同比"].notna(), None).tolist()
    sale_cur = comp.get("售汇_经常项目", pd.Series(index=main.index)).reindex(main.index).fillna(0).round(2).tolist()
    sale_cap = comp.get("售汇_资本项目", pd.Series(index=main.index)).reindex(main.index).fillna(0).round(2).tolist()
    chart2 = build_stack_bar_plus_yoy(
        xs,
        sale_cur,
        sale_cap,
        sale_yoy_pct,
        "售汇：经常项目 vs 资本项目(当月)",
        "经常项目",
        "资本项目",
    )
    show_chart(chart2, height="520px")

    # 图3：结售汇差额（单柱 + 同比）
    bal_yoy_pct = (main["差额同比"] * 100).where(main["差额同比"].notna(), None).tolist()
    chart3 = build_bar_plus_yoy(xs, main["差额"].round(2).tolist(), bal_yoy_pct, "结售汇差额(当月)")
    show_chart(chart3, height="520px")

    # 图4：远期结售汇签约额（多折线）
    if not fwd_sign.empty:
        series_map4 = {
            "结汇": fwd_sign["结汇"].reindex(main.index).round(2).tolist(),
            "售汇": fwd_sign["售汇"].reindex(main.index).round(2).tolist(),
            "差额": fwd_sign["差额"].reindex(main.index).round(2).tolist(),
        }
        chart4 = build_multi_line(xs, series_map4, "远期结售汇签约额")
        show_chart(chart4, height="520px")

    # 图5：远期结售汇累计未到期额（多折线）
    if not fwd_out.empty:
        series_map5 = {
            "结汇": fwd_out["结汇"].reindex(main.index).round(2).tolist(),
            "售汇": fwd_out["售汇"].reindex(main.index).round(2).tolist(),
            "差额": fwd_out["差额"].reindex(main.index).round(2).tolist(),
        }
        chart5 = build_multi_line(xs, series_map5, "远期结售汇累计未到期额")
        show_chart(chart5, height="520px")

    # 明细数据（直接展示，不折叠）
    table = main.copy()
    table.index = table.index.strftime("%Y-%m")
    st.dataframe(
        table[["结汇", "售汇", "差额", "结汇同比", "售汇同比", "差额同比"]].sort_index(ascending=False),
        use_container_width=True,
        hide_index=False,
    )

elif page == "利率数据":
    st.subheader("利率数据")
    st.caption("数据来源：FRED")

    @st.cache_data
    def _load_fred_rate_csv(path: str):
        import pandas as _pd
        try:
            df = _pd.read_csv(path, na_values=['.'])
        except Exception:
            return None
        if df is None or df.empty:
            return None
        cols = list(df.columns)
        date_col = 'DATE' if 'DATE' in cols else ('observation_date' if 'observation_date' in cols else None)
        if not date_col:
            return None
        val_cols = [c for c in cols if c != date_col]
        if not val_cols:
            return None
        val_col = val_cols[0]
        df = df[[date_col, val_col]].rename(columns={date_col: 'DATE', val_col: 'value'})
        df['DATE'] = _pd.to_datetime(df['DATE'], errors='coerce')
        df['value'] = _pd.to_numeric(df['value'], errors='coerce')
        df = df.dropna(subset=['DATE']).sort_values('DATE')
        df = df[df['DATE'] >= _pd.Timestamp('2020-01-01')]
        return df

    def _build_rate_line(x_list, y_list, title):
        chart = Line(init_opts=opts.InitOpts(bg_color="#FFFFFF", width="100%"))
        chart.add_xaxis(xaxis_data=x_list)
        chart.set_colors(PALETTE)
        chart.add_yaxis(
            series_name=title,
            y_axis=y_list,
            label_opts=opts.LabelOpts(is_show=False),
            is_smooth=False,
        )
        chart.set_global_opts(
            title_opts=opts.TitleOpts(title=title, pos_left="center", title_textstyle_opts=opts.TextStyleOpts(color="#111827")),
            tooltip_opts=opts.TooltipOpts(trigger="axis"),
            toolbox_opts=opts.ToolboxOpts(
                is_show=True,
                feature=opts.ToolBoxFeatureOpts(
                    save_as_image=opts.ToolBoxFeatureSaveAsImageOpts(),
                    data_view=opts.ToolBoxFeatureDataViewOpts(is_show=True, is_read_only=False),
                    data_zoom=opts.ToolBoxFeatureDataZoomOpts(),
                    magic_type=opts.ToolBoxFeatureMagicTypeOpts(type_=["line", "bar"]),
                    restore=opts.ToolBoxFeatureRestoreOpts(),
                ),
            ),
            xaxis_opts=opts.AxisOpts(type_="category", boundary_gap=False),
            yaxis_opts=opts.AxisOpts(name="利率(%)", min_="dataMin", max_="dataMax", is_scale=True),
            legend_opts=opts.LegendOpts(orient="horizontal", pos_top="40"),
            datazoom_opts=[
                opts.DataZoomOpts(type_="inside", filter_mode="filter"),
                opts.DataZoomOpts(type_="slider", pos_bottom="8", range_start=0, range_end=100, filter_mode="filter"),
            ],
        )
        return chart

    sof = _load_fred_rate_csv("SOFR.csv")
    son = _load_fred_rate_csv("IUDSOIA.csv")
    estr = _load_fred_rate_csv("ECBESTRVOLWGTTRMDMNRT.csv")

    if sof is not None and not sof.empty:
        st.subheader("SOFR")
        x = sof['DATE'].dt.strftime('%Y-%m-%d').tolist()
        y = sof['value'].round(4).tolist()
        show_chart(_build_rate_line(x, y, "SOFR"), height="420px")

    if son is not None and not son.empty:
        st.subheader("SONIA (IUDSOIA)")
        x = son['DATE'].dt.strftime('%Y-%m-%d').tolist()
        y = son['value'].round(4).tolist()
        show_chart(_build_rate_line(x, y, "SONIA"), height="420px")

    if estr is not None and not estr.empty:
        st.subheader("ESTR (ECBESTRVOLWGTTRMDMNRT)")
        x = estr['DATE'].dt.strftime('%Y-%m-%d').tolist()
        y = estr['value'].round(4).tolist()
        show_chart(_build_rate_line(x, y, "ESTR"), height="420px")

    st.markdown("<hr/>", unsafe_allow_html=True)
    st.subheader("明细数据")
    _tables = []
    if sof is not None and not sof.empty:
        _tables.append(sof.assign(Series='SOFR'))
    if son is not None and not son.empty:
        _tables.append(son.assign(Series='SONIA'))
    if estr is not None and not estr.empty:
        _tables.append(estr.assign(Series='ESTR'))
    if _tables:
        detail = pd.concat(_tables, ignore_index=True)
        detail.sort_values(['Series','DATE'], inplace=True)
        disp = detail.copy()
        disp['DATE'] = pd.to_datetime(disp['DATE']).dt.strftime('%Y-%m-%d')
        st.dataframe(disp[['Series','DATE','value']], use_container_width=True, hide_index=True)
    else:
        st.info("暂无可显示的明细数据")


elif page == "汇率数据":
    st.subheader("汇率数据")
    st.info("本板块建设中，敬请期待。")


elif page == "CME FEDWATCH":
    st.subheader("CME FedWatch")
    st.caption("数据来源：CME FedWatch")
    colA, colB = st.columns(2)
    with colA:
        st.markdown("**降息概率矩阵**")
    with colB:
        st.markdown("**点阵表格**")

    tab1, tab2 = st.tabs(["降息概率矩阵图", "点阵表格"])
    with tab1:
        prob = load_fedwatch_probabilities()
        if not prob:
            st.warning("未找到或无法解析 outputs/fedwatch_probabilities.json")
        else:
            hm = build_fedwatch_heatmap(prob)
            if hm:
                show_chart(hm, height="640px")
    with tab2:
        dp = load_fedwatch_dotplot()
        if not dp:
            st.warning("未找到或无法解析 outputs/fedwatch_dot_plot_table.json")
        else:
            df = build_fedwatch_dot_table(dp)
            if df is not None:
                st.dataframe(df, use_container_width=True, hide_index=True)
            else:
                st.info("点阵表格暂无可展示数据。")
