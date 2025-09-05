# 全局配置项，供各脚本复用，避免硬编码与重复

RAW_DATA_PATH = "raw_csv_data"
OUTPUT_FILENAME = "海关统计数据汇总.xlsx"

# 海关官网数据入口
BASE_URL = (
    "http://www.customs.gov.cn/customs/302249/zfxxgk/2799825/302274/302277/6348926/index.html"
)

# 浏览器与网络
import os
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/99.0.4844.84 Safari/537.36"
)
# 在云端（无 DISPLAY）必须 headless 运行；允许用环境变量 PLAYWRIGHT_HEADLESS=0 在本地开发时打开可见窗口
HEADLESS = (os.environ.get("PLAYWRIGHT_HEADLESS", "1") != "0")

# 网络/超时配置（适配慢网与代理场景）
# 页面导航最大等待（毫秒）
NAV_TIMEOUT_MS = 180_000
# 选择器等待（毫秒）
SELECTOR_TIMEOUT_MS = 120_000
# 导航重试
NAV_MAX_RETRIES = 5
NAV_RETRY_DELAY_SEC = 6
# 慢速模式（毫秒）：遇到被动防爬/代理导致抖动时可开启
SLOW_MO_MS = 0

# 抗封策略：请求节流/抖动时间（秒）
REQUEST_JITTER_MIN_SEC = 2.0
REQUEST_JITTER_MAX_SEC = 5.0

# 详情页打开与解析的最大重试次数
DETAIL_MAX_RETRIES = 3

# 业务地区配置
# 注意：包含"总值"，在处理后会重命名为"全国"
TARGET_LOCATIONS = [
    "总值",
    "北京市",
    "上海市",
    "深圳市",
    "南京市",
    "合肥市",
    "浙江省",
    "杭州市",
    "宁波市",
    "温州市",
    "湖州市",
    "金华市",
    "台州市",
    # 新增浙江省地市
    "嘉兴市",
    "丽水市", 
    "衢州市",
    "绍兴市",
    "舟山市",
]

FINAL_LOCATIONS = ["全国"] + [loc for loc in TARGET_LOCATIONS if loc != "总值"]


