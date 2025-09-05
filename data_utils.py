import os
import glob
from datetime import datetime
from typing import List, Dict

import pandas as pd

from config import RAW_DATA_PATH, OUTPUT_FILENAME, TARGET_LOCATIONS, FINAL_LOCATIONS


def ensure_raw_data_dir() -> str:
    os.makedirs(RAW_DATA_PATH, exist_ok=True)
    return RAW_DATA_PATH


def month_file_name(year: int, month: int) -> str:
    return f"{year}-{month:02d}.csv"


def list_existing_month_files() -> List[str]:
    ensure_raw_data_dir()
    return sorted(os.listdir(RAW_DATA_PATH))


def parse_table_html_to_df(table_html: str) -> pd.DataFrame:
    # 多级表头：header=[0, 1]
    dataframes = pd.read_html(table_html, header=[0, 1])
    if not dataframes:
        raise ValueError("页面表格解析失败：未检测到可用表格")
    return dataframes[0]


def save_raw_df(df: pd.DataFrame, filepath: str) -> None:
    df.to_csv(filepath, index=False, encoding="utf-8-sig")


def read_all_raw_csv_files() -> List[str]:
    """读取所有原始CSV文件，包括全国数据和浙江省数据"""
    pattern = os.path.join(RAW_DATA_PATH, "*.csv")
    return glob.glob(pattern)


def read_zhejiang_csv_files() -> List[str]:
    """专门读取浙江省数据CSV文件"""
    pattern = os.path.join(RAW_DATA_PATH, "浙江省-*.csv")
    return glob.glob(pattern)


def read_national_csv_files() -> List[str]:
    """读取全国数据CSV文件（排除浙江省数据）"""
    pattern = os.path.join(RAW_DATA_PATH, "[0-9]*.csv")  # 只匹配以数字开头的文件
    return glob.glob(pattern)


def tidy_one_month_csv(file_path: str) -> pd.DataFrame:
    # 针对原始CSV：第一行为第二级表头名，需要 header=1 再 drop 首行
    df = pd.read_csv(file_path, header=1)
    df.drop(index=0, inplace=True)

    # 从文件名提取年月
    filename = os.path.basename(file_path)
    year, month = map(int, filename.replace(".csv", "").split("-"))
    df["时间"] = f"{year}-{month:02d}"

    # 统一提取前7列以及时间列
    required_cols_df = df.iloc[:, [0, 1, 2, 3, 4, 5, 6, -1]].copy()

    # 根据原始列名准确判断“出口/进口”先后顺序
    # 读取 header=1 后，列名一般为：收发货人所在地, 进出口, 进出口, 出口/进口, 出口/进口, 进口/出口, 进口/出口
    # 我们识别第3列与第5列（索引2与4）是否为“出口”，据此选择映射，避免把出口/进口写反
    col2 = str(df.columns[2]) if len(df.columns) > 2 else ""
    col4 = str(df.columns[4]) if len(df.columns) > 4 else ""

    if ("出口" in col2 and "进口" in col4):
        # 正常顺序：进出口、进出口、出口、出口、进口、进口
        mapping = [
            "地区",
            "进出口_当月",
            "进出口_年初至今",
            "出口_当月",
            "出口_年初至今",
            "进口_当月",
            "进口_年初至今",
            "时间",
        ]
    elif ("进口" in col2 and "出口" in col4):
        # 颠倒顺序：进出口、进出口、进口、进口、出口、出口
        mapping = [
            "地区",
            "进出口_当月",
            "进出口_年初至今",
            "进口_当月",
            "进口_年初至今",
            "出口_当月",
            "出口_年初至今",
            "时间",
        ]
    else:
        # 兜底：按正常顺序映射（大多数月份）
        mapping = [
            "地区",
            "进出口_当月",
            "进出口_年初至今",
            "出口_当月",
            "出口_年初至今",
            "进口_当月",
            "进口_年初至今",
            "时间",
        ]

    required_cols_df.columns = mapping

    filtered_df = required_cols_df[required_cols_df["地区"].isin(TARGET_LOCATIONS)]
    return filtered_df


def tidy_zhejiang_csv(file_path: str) -> pd.DataFrame:
    """处理浙江省数据CSV文件"""
    df = pd.read_csv(file_path)
    
    # 从文件名提取年月：浙江省-YYYY-MM.csv
    filename = os.path.basename(file_path)
    parts = filename.replace("浙江省-", "").replace(".csv", "").split("-")
    year, month = int(parts[0]), int(parts[1])
    
    # 添加时间列
    df["时间"] = f"{year}-{month:02d}"
    
    # 浙江省数据的地市列表 
    zhejiang_cities = [
        "合计", "杭州地区", "湖州地区", "嘉兴地区", "金华地区", 
        "丽水地区", "宁波地区", "衢州地区", "绍兴地区", 
        "台州地区", "温州地区", "舟山地区"
    ]
    
    # 过滤只保留浙江省地市数据，排除"合计"
    filtered_df = df[df["收发货人所在地"].isin(zhejiang_cities[1:])].copy()  # 排除"合计"
    
    # 重命名列以匹配全国数据格式
    column_mapping = {
        "收发货人所在地": "地区",
        "当期进出口": "进出口_年初至今", 
        "当期进口": "进口_年初至今",
        "当期出口": "出口_年初至今",
        "进出口同比": "进出口_年初至今同比",
        "进口同比": "进口_年初至今同比", 
        "出口同比": "出口_年初至今同比"
    }
    
    filtered_df = filtered_df.rename(columns=column_mapping)
    
    # 统一地区名称：移除"地区"后缀
    filtered_df["地区"] = filtered_df["地区"].str.replace("地区", "市")
    
    # 单位转换：2024年7月开始数据单位从元改为万元，需要乘以10000统一为元
    if year == 2024 or year >= 2025:
        value_columns = ["进出口_年初至今", "进口_年初至今", "出口_年初至今"]
        for col in value_columns:
            if col in filtered_df.columns:
                filtered_df[col] = filtered_df[col] * (1/10000)
    
    return filtered_df


def calculate_monthly_from_cumulative(df: pd.DataFrame) -> pd.DataFrame:
    """从累计数据计算单月数据"""
    df = df.sort_values(['地区', '时间']).reset_index(drop=True)
    
    # 为每个地区计算单月数据
    result_rows = []
    
    for region in df['地区'].unique():
        region_data = df[df['地区'] == region].copy()
        region_data = region_data.sort_values('时间').reset_index(drop=True)
        
        for i, row in region_data.iterrows():
            new_row = row.copy()
            
            # 解析年月
            year_month = row['时间'].split('-')
            year, month = int(year_month[0]), int(year_month[1])
            
            if month == 1:
                # 1月数据就是累计数据
                new_row['进出口_当月'] = row.get('进出口_年初至今', 0)
                new_row['进口_当月'] = row.get('进口_年初至今', 0) 
                new_row['出口_当月'] = row.get('出口_年初至今', 0)
            else:
                # 找到上一个月的数据
                prev_month = f"{year}-{month-1:02d}"
                prev_data = region_data[region_data['时间'] == prev_month]
                
                if not prev_data.empty:
                    prev_row = prev_data.iloc[0]
                    # 当月累计 - 上月累计 = 当月单月
                    new_row['进出口_当月'] = row.get('进出口_年初至今', 0) - prev_row.get('进出口_年初至今', 0)
                    new_row['进口_当月'] = row.get('进口_年初至今', 0) - prev_row.get('进口_年初至今', 0)
                    new_row['出口_当月'] = row.get('出口_年初至今', 0) - prev_row.get('出口_年初至今', 0)
                else:
                    # 找不到上月数据，设为0
                    new_row['进出口_当月'] = 0
                    new_row['进口_当月'] = 0
                    new_row['出口_当月'] = 0
            
            result_rows.append(new_row)
    
    return pd.DataFrame(result_rows)


def consolidate_with_yoy(all_csv_files: List[str] = None) -> pd.DataFrame:
    """整合所有数据，包括全国数据和浙江省数据"""
    
    # 如果没有传入文件列表，自动读取所有文件
    if all_csv_files is None:
        all_csv_files = read_all_raw_csv_files()
    
    # 分离全国数据和浙江省数据
    national_files = [f for f in all_csv_files if not os.path.basename(f).startswith("浙江省-")]
    zhejiang_files = [f for f in all_csv_files if os.path.basename(f).startswith("浙江省-")]
    
    print(f"📊 处理数据文件:")
    print(f"  📈 全国数据文件: {len(national_files)} 个")
    print(f"  🏙️  浙江省数据文件: {len(zhejiang_files)} 个")
    
    # 处理全国数据
    national_rows = []
    for file_path in national_files:
        try:
            national_df = tidy_one_month_csv(file_path)
            national_rows.append(national_df)
        except Exception as exc:
            print(f"处理全国数据文件失败，已跳过: {file_path} -> {exc}")
    
    # 处理浙江省数据
    zhejiang_rows = []
    zhejiang_cities = set()  # 记录浙江省地市名称
    
    # 先收集所有浙江省累计数据
    for file_path in zhejiang_files:
        try:
            zj_df = tidy_zhejiang_csv(file_path)
            zhejiang_rows.append(zj_df)
            # 记录浙江省地市名称，用于后续去重
            zhejiang_cities.update(zj_df['地区'].unique())
        except Exception as exc:
            print(f"处理浙江省数据文件失败，已跳过: {file_path} -> {exc}")
    
    # 合并所有浙江省数据后再计算单月数据
    if zhejiang_rows:
        zhejiang_combined = pd.concat(zhejiang_rows, ignore_index=True)
        zhejiang_with_monthly = calculate_monthly_from_cumulative(zhejiang_combined)
        zhejiang_rows = [zhejiang_with_monthly]  # 重新包装为列表形式
    
    print(f"🏙️  浙江省地市列表: {sorted(list(zhejiang_cities))}")
    
    # 从全国数据中移除浙江省地市数据，避免重复
    filtered_national_rows = []
    for national_df in national_rows:
        # 过滤掉浙江省地市，保留其他地区
        filtered_df = national_df[~national_df['地区'].isin(zhejiang_cities)].copy()
        if not filtered_df.empty:
            filtered_national_rows.append(filtered_df)
    
    print(f"📈 全国数据去除浙江省地市后，保留地区数: {len(set().union(*[df['地区'].unique() for df in filtered_national_rows]) if filtered_national_rows else set())}")
    
    # 合并所有数据
    all_rows = filtered_national_rows + zhejiang_rows
    
    if not all_rows:
        raise RuntimeError("没有可整合的数据")

    master_df = pd.concat(all_rows, ignore_index=True)

    # 统一将"总值"改为"全国"
    master_df["地区"] = master_df["地区"].replace({"总值": "全国"})

    # 数值化与排序
    value_cols = [
        "进出口_当月",
        "进出口_年初至今",
        "出口_当月",
        "出口_年初至今",
        "进口_当月",
        "进口_年初至今",
    ]
    for col in value_cols:
        master_df[col] = pd.to_numeric(master_df[col], errors="coerce")

    master_df["时间"] = pd.to_datetime(master_df["时间"])  # yyyy-mm
    master_df.sort_values(by=["地区", "时间"], inplace=True)

    # 同比 pct_change(12)
    for col in value_cols:
        yoy_col = f"{col}同比"
        master_df[yoy_col] = master_df.groupby("地区")[col].pct_change(12)

    return master_df


def export_to_excel_by_location(master_df: pd.DataFrame, output_path: str = OUTPUT_FILENAME) -> None:
    """导出整合数据到Excel，包括全国数据和浙江省地市数据"""
    
    # 获取所有地区并去重，优先保留浙江省数据
    all_regions = master_df["地区"].unique()
    
    # 浙江省地市名单（从浙江省数据中获得，优先级高）
    zhejiang_cities = ["杭州市", "湖州市", "嘉兴市", "金华市", "丽水市", 
                      "宁波市", "衢州市", "绍兴市", "台州市", "温州市", "舟山市"]
    
    # 需要导出的地区（去重逻辑：如果有浙江省地市数据，就不导出全国数据中的对应地市）
    export_regions = []
    
    # 首先添加全国等非地市数据
    for region in FINAL_LOCATIONS:
        if region not in zhejiang_cities:  # 不是浙江省地市
            export_regions.append(region)
    
    # 然后添加浙江省地市数据（如果存在）
    for city in zhejiang_cities:
        if city in all_regions:
            export_regions.append(city)
        else:
            # 如果浙江省数据中没有，尝试从全国数据中获取
            national_city_name = city  # 已经是"市"结尾了
            if national_city_name in all_regions:
                export_regions.append(national_city_name)
    
    print(f"📤 准备导出 {len(export_regions)} 个地区的数据")
    print(f"  包含浙江省地市: {[city for city in export_regions if city in zhejiang_cities]}")
    
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for location in export_regions:
            location_df = master_df[master_df["地区"] == location].copy()
            if location_df.empty:
                continue

            location_df["时间"] = location_df["时间"].dt.strftime("%Y-%m")

            final_columns_order = [
                "时间",
                "进出口_当月",
                "进出口_当月同比",
                "进出口_年初至今",
                "进出口_年初至今同比",
                "出口_当月",
                "出口_当月同比",
                "出口_年初至今",
                "出口_年初至今同比",
                "进口_当月",
                "进口_当月同比",
                "进口_年初至今",
                "进口_年初至今同比",
            ]
            for col in final_columns_order:
                if col not in location_df.columns:
                    location_df[col] = None

            location_df = location_df[final_columns_order]
            location_df.to_excel(writer, sheet_name=location, index=False)
    
    print(f"✅ Excel文件已导出: {output_path}")




