import os
import subprocess
import sys
import time
from datetime import datetime

import pandas as pd
import streamlit as st
from pyecharts import options as opts
from pyecharts.charts import Line, Bar
from streamlit_echarts import st_pyecharts

from config import OUTPUT_FILENAME, FINAL_LOCATIONS


# =============================
# 页面与主题配置
# =============================
st.set_page_config(
    page_title="海关进出口数据看板（明亮美化版）",
    layout="wide",
    page_icon=None,
)

# 与莫兰迪暖色背景协调的系列配色（低饱和、柔和）
PALETTE = [
    "#6B7A8F",  # 蓝灰（进出口）
    "#9AAE8C",  # 鼠尾草绿（出口/进口其一）
    "#C9A27E",  # 驼色（出口/进口其一）
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

html, body, [class*="st-"] {
  font-family: Inter, "HarmonyOS Sans SC", "Microsoft YaHei", system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif;
  background: var(--bg);
  color: var(--text);
}

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


# =============================
# 数据加载与更新功能
# =============================
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
    """运行数据更新器"""
    try:
        # 运行自动数据更新器脚本
        result = subprocess.run(
            [sys.executable, "自动数据更新器.py"],
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='ignore',  # 忽略编码错误
            timeout=600  # 10分钟超时
        )
        
        if result.returncode == 0:
            # 解析输出结果
            output = result.stdout or ""
            output_lines = output.strip().split('\n') if output else []
            
            # 查找结果信息
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


def build_line(x_list, series_dict, title):
    """按原可视化脚本的稳定实现，确保与 pyecharts 2.0.8 兼容。"""
    # 设置白色背景，避免与页面背景叠加导致透明
    chart = Line(init_opts=opts.InitOpts(bg_color="#FFFFFF"))
    chart.add_xaxis(xaxis_data=x_list)
    # 设置系列颜色，确保与页面背景协调
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


# =============================
# 9大类产品类别看板 - 常量与函数
# =============================
# 类别固定顺序与颜色映射（与产品类别看板一致）
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


@st.cache_data
def load_category_data():
    """加载 9 大类产品数据 Excel（各 Sheet 为进出口/出口/进口，币种区分）。"""
    filename = "9大类产品分析表.xlsx"
    if not os.path.exists(filename):
        return None
    try:
        xls = pd.ExcelFile(filename)
        data = {}
        for sheet in xls.sheet_names:
            df = pd.read_excel(filename, sheet_name=sheet, index_col=0)
            data[sheet] = df
        return data
    except Exception as e:
        st.error(f"加载9大类产品 Excel 失败: {e}")
        return None


def create_horizontal_percentage_chart(data, title, regions, categories):
    """创建横向 100% 堆叠条形图（与产品类别看板一致）。"""
    # 固定类别顺序，并保证缺失列补 0
    categories_ordered = [c for c in CATEGORY_ORDER if c in categories]
    if not categories_ordered:
        categories_ordered = categories
    data = data.reindex(columns=categories_ordered, fill_value=0)

    # 计算百分比数据（按行归一化）
    row_sums = data.sum(axis=1).replace(0, 1)
    percentage_data = data.div(row_sums, axis=0) * 100.0

    # 反转地区顺序，使第一个地区显示在顶部
    reversed_regions = list(reversed(regions))

    # 根据地区数量动态设置图表高度
    chart_height = max(400, len(regions) * 60 + 150)

    # 设置白色背景
    chart = Bar(init_opts=opts.InitOpts(bg_color="#FFFFFF", width="100%", height=f"{chart_height}px"))

    # 添加 x 轴（横向图中地区在 x 轴，后续会 reversal）
    chart.add_xaxis(reversed_regions)

    # 每个类别添加数据系列
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
            category_gap="60%",
            itemstyle_opts=opts.ItemStyleOpts(color=CATEGORY_COLOR_MAP.get(category)),
        )

    chart.set_colors(CATEGORY_PALETTE)

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

    # 转换为横向
    chart.reversal_axis()
    return chart


# =============================
# 页面结构
# =============================
# 顶部标题
st.title("海关进出口数据看板")
st.markdown("")

# 预加载综合数据（用于“综合看板”页）
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


# 侧边栏：页面选择 + 分页筛选
with st.sidebar:
    st.header("页面选择")
    _page_options = ["综合看板", "产品类别看板"]
    _page_labels = {"综合看板": "📊 综合看板", "产品类别看板": "📦 产品类别看板"}
    page = st.radio(
        "选择页面",
        _page_options,
        index=0,
        horizontal=False,
        format_func=lambda x: _page_labels.get(x, x),
    )
    st.markdown("---")

    if page == "综合看板":
        st.header("操作面板")
        if latest_info:
            st.caption(latest_info)

        if latest_info_zhejiang:
            st.caption(latest_info_zhejiang)    


        st.header("地区筛选")
        default_index = FINAL_LOCATIONS.index("浙江省") if "浙江省" in FINAL_LOCATIONS else 0
        selected_location = st.selectbox("", options=FINAL_LOCATIONS, index=default_index)

        st.header("展示设置")
        show_overview = st.checkbox("显示全国与重点地区概览", value=True)

    else:  # 产品类别看板
        st.header("数据控制面板")

        # 可隐藏筛选：使用 session_state 控制
        if "cat_filters_visible" not in st.session_state:
            st.session_state["cat_filters_visible"] = True

        def hide_cat_filters():
            st.session_state["cat_filters_visible"] = False

        def show_cat_filters():
            st.session_state["cat_filters_visible"] = True

        category_data_sidebar = load_category_data()
        all_regions_sidebar = []
        if category_data_sidebar:
            sample_sheet = list(category_data_sidebar.keys())[0]
            all_regions_sidebar = category_data_sidebar[sample_sheet].index.tolist()

        if st.session_state.get("cat_filters_visible", True):
            st.subheader("币种")
            st.radio(
                "选择币种:",
                options=["美元", "人民币"],
                index=0,
                horizontal=True,
                key="cat_currency",
            )

            st.subheader("地区筛选")
            default_regions = all_regions_sidebar[:6] if len(all_regions_sidebar) >= 6 else all_regions_sidebar
            st.multiselect(
                "选择要显示的地区:",
                options=all_regions_sidebar,
                default=default_regions,
                key="cat_selected_regions",
                help="最多选择8个地区以保证图表清晰度",
            )
            # 限制最多 8 个地区
            if st.session_state.get("cat_selected_regions") and len(st.session_state["cat_selected_regions"]) > 8:
                st.warning("为保证图表清晰度，建议最多选择8个地区")
                st.session_state["cat_selected_regions"] = st.session_state["cat_selected_regions"][:8]

            st.markdown("---")
            st.subheader("产品类别")
            st.multiselect(
                "选择要显示的产品类别:",
                options=CATEGORY_ORDER,
                default=CATEGORY_ORDER,
                key="cat_selected_categories",
                help="可以选择部分类别进行对比分析",
            )

            st.markdown("---")
            st.button("应用并隐藏筛选", type="primary", use_container_width=True, on_click=hide_cat_filters)
        else:
            # 隐藏状态下展示简洁摘要
            cur = st.session_state.get("cat_currency", "美元")
            regs = st.session_state.get("cat_selected_regions", [])
            cats = st.session_state.get("cat_selected_categories", CATEGORY_ORDER)

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


# 页面一：综合看板
if page == "综合看板":
    if not data:
        st.info("未检测到数据文件，请先运行数据处理脚本生成 Excel。")
        st.stop()

    # 全国概览卡片
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
                cls, delta = fmt_delta(yoy)
                c.markdown(
                    f"""
                    <div class=\"card\">\n                      <div class=\"label\">{label}</div>\n                      <div class=\"value\">{fmt_value(val)}</div>\n                      {f'<span class=\"chip {cls}\">{delta}</span>' if delta else ''}\n                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    st.markdown("<hr/>", unsafe_allow_html=True)

    # 重点地区概览（顶部几大地区）
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
                cls, delta = fmt_delta(yoy)
                c.markdown(
                    f"""
                    <div class=\"card\">\n                      <div class=\"label\">{label}</div>\n                      <div class=\"value\">{fmt_value(val)}</div>\n                      {f'<span class=\"chip {cls}\">{delta}</span>' if delta else ''}\n                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            if loc == "浙江省":
                ZHEJIANG_CITIES = ["杭州市", "宁波市", "温州市", "湖州市", "金华市", "台州市", "嘉兴市", "丽水市", "衢州市", "绍兴市", "舟山市"]
                with st.expander("展开/收起 浙江省地市数据概览"):
                    for city in ZHEJIANG_CITIES:
                        city_df = data.get(city)
                        if city_df is None or city_df.empty:
                            continue
                        latest_city = city_df.iloc[city_df["时间"].map(pd.to_datetime).idxmax()]
                        st.markdown(f"#### {city}")
                        city_cols = st.columns(3)
                        city_metrics = [
                            ("进出口(年初至今)", latest_city.get("进出口_年初至今"), latest_city.get("进出口_年初至今同比")),
                            ("进口(年初至今)", latest_city.get("进口_年初至今"), latest_city.get("进口_年初至今同比")),
                            ("出口(年初至今)", latest_city.get("出口_年初至今"), latest_city.get("出口_年初至今同比")),
                        ]
                        for c, (label, val, yoy) in zip(city_cols, city_metrics):
                            cls, delta = fmt_delta(yoy)
                            c.markdown(
                                f"""
                                <div class=\"card\">\n                                  <div class=\"label\">{label}</div>\n                                  <div class=\"value\">{fmt_value(val)}</div>\n                                  {f'<span class=\"chip {cls}\">{delta}</span>' if delta else ''}\n                                </div>
                                """,
                                unsafe_allow_html=True,
                            )

    st.markdown("<hr/>", unsafe_allow_html=True)

    # 选中地区详情
    st.subheader(f"{selected_location} · 详细数据与走势")
    loc_df = data.get(selected_location)
    if loc_df is None or loc_df.empty:
        st.warning(f"未找到 {selected_location} 的数据")
    else:
        loc_df_sorted = loc_df.copy()
        loc_df_sorted["时间"] = pd.to_datetime(loc_df_sorted["时间"])  # 保障类型
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
            for col in disp.columns:
                if "同比" in col:
                    disp[col] = disp[col].apply(lambda v: f"{v:.2%}" if pd.notna(v) else "—")
            disp["时间"] = disp["时间"].dt.strftime("%Y-%m")
            st.dataframe(
                disp.sort_values("时间", ascending=False),
                use_container_width=True,
                hide_index=True,
            )

# 页面二：产品类别看板
else:
    category_data = load_category_data()
    if not category_data:
        st.error("未找到数据文件 '9大类产品分析表.xlsx'，请先运行数据处理脚本生成Excel文件。")
        st.stop()

    # 从会话状态读取筛选（若无则给出合理默认）
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
    selected_categories = st.session_state.get("cat_selected_categories", CATEGORY_ORDER)

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
                st.markdown('<div class="chart-rounded">', unsafe_allow_html=True)
                chart_height = max(400, len(available_regions) * 60 + 150)
                st_pyecharts(chart, height=f"{chart_height}px")
                st.markdown('</div>', unsafe_allow_html=True)

                # 直接展示两张表（并排）
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



