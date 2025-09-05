from data_utils import read_all_raw_csv_files, consolidate_with_yoy, export_to_excel_by_location
from config import RAW_DATA_PATH


def process_and_consolidate_data() -> None:
    """
    读取 `raw_csv_data/` 下所有CSV，清洗、整合并生成带多Sheet的 `海关统计数据汇总.xlsx`。
    包括全国数据和浙江省数据的整合处理。
    """
    all_csv_files = read_all_raw_csv_files()
    if not all_csv_files:
        print(f"错误：在 '{RAW_DATA_PATH}' 文件夹中未找到任何CSV文件。请先运行爬取脚本。")
        return

    print(f"🚀 海关数据整合处理器")
    print("=" * 50)
    print(f"📁 找到 {len(all_csv_files)} 个原始数据文件")
    
    # 分析文件类型
    national_files = [f for f in all_csv_files if not "浙江省-" in f]
    zhejiang_files = [f for f in all_csv_files if "浙江省-" in f]
    
    print(f"  📈 全国数据文件: {len(national_files)} 个")
    print(f"  🏙️  浙江省数据文件: {len(zhejiang_files)} 个")
    print("\n🔄 开始整合处理...")
    
    # 整合所有数据（自动处理全国和浙江省数据）
    master_df = consolidate_with_yoy()
    
    print(f"📊 数据整合完成，共处理 {len(master_df)} 行记录")
    print(f"📍 包含地区: {len(master_df['地区'].unique())} 个")
    
    # 导出到Excel
    export_to_excel_by_location(master_df)
    print("\n🎉 任务成功完成！整合后的Excel文件已生成。")


if __name__ == "__main__":
    process_and_consolidate_data()
