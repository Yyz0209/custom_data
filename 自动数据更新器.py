#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动数据更新器
检查海关总署和杭州海关是否有新的月份数据，如果有则自动下载、处理并更新Excel文件
"""

import os
import re
import sys
import time
import glob
from datetime import datetime, timedelta
from typing import List, Tuple, Optional
import logging

import random

from playwright.sync_api import sync_playwright
import pandas as pd

from config import (
    BASE_URL, RAW_DATA_PATH, OUTPUT_FILENAME,
    HEADLESS, NAV_TIMEOUT_MS, SELECTOR_TIMEOUT_MS,
    NAV_MAX_RETRIES, NAV_RETRY_DELAY_SEC, SLOW_MO_MS
)
from data_utils import (
    ensure_raw_data_dir, month_file_name, parse_table_html_to_df,
    save_raw_df, consolidate_with_yoy, export_to_excel_by_location
)

# 云端强制启用无头模式（无 DISPLAY 时）
try:
    if not os.environ.get("DISPLAY"):
        HEADLESS = True  # 覆盖配置中的 HEADLESS，避免 X server 错误
except Exception:
    pass

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('数据更新日志.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# 确保 Playwright 浏览器已安装（云端环境需要）
try:
    import subprocess as _sp
    os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", "0")
    print("[Playwright] Installing browsers to local cache (PLAYWRIGHT_BROWSERS_PATH=0)…", flush=True)
    _sp.run([sys.executable, "-m", "playwright", "install", "chromium", "chromium-headless-shell"], check=False, capture_output=False)
except Exception as _e:
    print(f"[Playwright] Install step skipped/failed: {_e}", flush=True)

logger = logging.getLogger(__name__)


class DataUpdateChecker:
    """数据更新检查器"""

    def __init__(self):
        self.current_year = datetime.now().year
        self.current_month = datetime.now().month

        # 浙江省配置
        self.zhejiang_config = {
            "base_url": "http://hangzhou.customs.gov.cn",
            "years": {
                "2024": {
                    "year_id": "5776888",
                    "months": {
                        "一月": {"month_id": "5784442"}, "二月": {"month_id": "5778212"},
                        "三月": {"month_id": "5782661"}, "四月": {"month_id": "5870816"},
                        "五月": {"month_id": "5870859"}, "六月": {"month_id": "5955471"},
                        "七月": {"month_id": "6053528"}, "八月": {"month_id": "6089867"},
                        "九月": {"month_id": "6163407"}, "十月": {"month_id": "6222827"},
                        "十一月": {"month_id": "6275194"}, "十二月": {"month_id": "6332730"}
                    }
                },
                "2025": {
                    "year_id": "6430241",
                    "months": {
                        "一月": {"month_id": "6430315"}, "二月": {"month_id": "6430326"},
                        "三月": {"month_id": "6478170"}, "四月": {"month_id": "6531356"},
                        "五月": {"month_id": "6585867"}, "六月": {"month_id": "6633632"},
                        "七月": {"month_id": "6691201"},
                        # 注意：7月及以后的数据需要根据实际发布情况动态添加
                        # 如果发现新的月份数据，需要手动更新这个配置
                    }
                }
            }
        }

        self.month_mapping = {
            "一月": 1, "二月": 2, "三月": 3, "四月": 4, "五月": 5, "六月": 6,
            "七月": 7, "八月": 8, "九月": 9, "十月": 10, "十一月": 11, "十二月": 12
        }

    def get_local_data_status(self) -> Tuple[List[Tuple[int, int]], List[Tuple[int, int]]]:
        """获取本地已有数据的状态"""
        ensure_raw_data_dir()

        # 检查全国数据
        national_files = glob.glob(os.path.join(RAW_DATA_PATH, "[0-9]*.csv"))
        national_months = []
        for file_path in national_files:
            filename = os.path.basename(file_path)
            match = re.match(r'(\d{4})-(\d{2})\.csv', filename)
            if match:
                year, month = int(match.group(1)), int(match.group(2))
                national_months.append((year, month))

        # 检查浙江省数据
        zj_files = glob.glob(os.path.join(RAW_DATA_PATH, "浙江省-*.csv"))
        zj_months = []
        for file_path in zj_files:
            filename = os.path.basename(file_path)
            match = re.match(r'浙江省-(\d{4})-(\d{2})\.csv', filename)
            if match:
                year, month = int(match.group(1)), int(match.group(2))
                zj_months.append((year, month))

        logger.info(f"本地全国数据: {len(national_months)} 个月份")
        logger.info(f"本地浙江省数据: {len(zj_months)} 个月份")

        return sorted(national_months), sorted(zj_months)

    def get_next_expected_months(self, existing_months: List[Tuple[int, int]]) -> List[Tuple[int, int]]:
        """根据现有数据，计算预期的下一个月份"""
        if not existing_months:
            return [(2024, 1)]  # 如果没有数据，从2024年1月开始

        # 找到最新的月份
        latest_year, latest_month = max(existing_months)
        next_months = []

        # 计算下一个月份
        current_date = datetime.now()
        check_date = datetime(latest_year, latest_month, 1)

        # 计算到上个月为止的所有缺失月份
        # 海关数据通常有1个月的延迟
        target_year = current_date.year
        target_month = current_date.month - 1
        if target_month == 0:
            target_year -= 1
            target_month = 12

        # 从下一个月开始，一直到目标月份
        temp_year, temp_month = latest_year, latest_month
        while (temp_year, temp_month) < (target_year, target_month):
            # 计算下一个月
            if temp_month == 12:
                temp_year += 1
                temp_month = 1
            else:
                temp_month += 1

            next_months.append((temp_year, temp_month))

        return next_months

    def check_national_data_available(self, year: int, month: int) -> bool:
        """检查全国数据是否可用"""
        logger.info(f"检查全国 {year}年{month}月 数据...")

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=HEADLESS,
                    slow_mo=SLOW_MO_MS,
                    args=[
                        '--no-sandbox',
                        '--disable-dev-shm-usage',
                        '--disable-blink-features=AutomationControlled'
                    ]
                )
                context = browser.new_context(
                    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
                    viewport={'width': 1920, 'height': 1080},
                    locale='zh-CN'
                )
                page = context.new_page()

                # 导航到主页面
                page.goto(BASE_URL, timeout=NAV_TIMEOUT_MS)
                page.wait_for_load_state("domcontentloaded")

                # 点击年份
                year_button_selector = f"//a[contains(text(), '{year}')]"
                page.wait_for_selector(year_button_selector, timeout=20000).click()
                time.sleep(2)

                # 检查是否有对应月份的链接（兼容“7月/7月份/2025年7月”等形式）
                table_row_selector = "//tr[contains(., '进出口商品收发货人所在地总值表')]"
                row = page.wait_for_selector(table_row_selector, timeout=20000)

                month_texts = []
                for link in row.query_selector_all("a"):
                    if link.get_attribute("href"):
                        try:
                            text = (link.inner_text() or "").strip()
                            if text:
                                month_texts.append(text)
                        except Exception:
                            continue

                target_month = f"{month}月"
                patterns = [
                    target_month,
                    f"{month}月份",
                    f"{year}年{month}月",
                    f"{year}年{month}月份",
                ]
                available = any(any(p in t for p in patterns) for t in month_texts)
                logger.info(f"全国月份链接文本: {month_texts}")

                browser.close()
                return available

        except Exception as e:
            logger.error(f"检查全国数据时出错: {e}")
            return False

    def discover_zhejiang_data_structure(self, year: int) -> dict:
        """动态发现杭州海关网站的年份和月份结构"""
        logger.info(f"动态发现杭州海关 {year}年 的数据结构...")

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=HEADLESS,
                    slow_mo=SLOW_MO_MS,
                    args=[
                        '--no-sandbox',
                        '--disable-dev-shm-usage',
                        '--disable-blink-features=AutomationControlled'
                    ]
                )
                context = browser.new_context(
                    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
                    viewport={'width': 1920, 'height': 1080},
                    locale='zh-CN'
                )
                page = context.new_page()

                # 访问杭州海关主页面
                main_url = f"{self.zhejiang_config['base_url']}/hangzhou_customs/575609/zlbd/575612/575612/index.html"
                page.goto(main_url, timeout=NAV_TIMEOUT_MS)
                page.wait_for_load_state('domcontentloaded')
                page.wait_for_timeout(5000)  # 增加等待时间，确保页面完全加载

                discovered_months = {}

                # 查找年份链接（兼容“2025”“2025年”“2025 年” 等文本）
                year_links = page.locator(f'a:has-text("{year}")').all()
                if not year_links:
                    year_links = page.locator(f'a:has-text("{year}年")').all()
                if not year_links:
                    # 兜底：遍历所有 a 标签，查找包含年份的
                    try:
                        for a in page.locator('a').all():
                            text = (a.inner_text() or '').strip()
                            if str(year) in text:
                                year_links = [a]
                                break
                    except Exception:
                        pass
                if not year_links:
                    logger.info(f"未找到{year}年的链接")
                    browser.close()
                    return {}

                # 点击年份链接
                year_links[0].click()
                page.wait_for_load_state('domcontentloaded')
                page.wait_for_timeout(4000)  # 增加等待时间，确保年份页面完全加载

                # 查找所有月份链接
                month_links = page.locator('a').all()
                for link in month_links:
                    try:
                        link_text = (link.inner_text() or '').strip()
                        link_href = link.get_attribute('href')

                        if not link_text or not link_href:
                            continue

                        # 检查是否是月份链接：兼容“七月/七月份/7月/7月份”
                        month_num = None
                        # 1) 阿拉伯数字
                        num_match = re.search(r'(\d{1,2})\s*月', link_text)
                        if not num_match:
                            num_match = re.search(r'(\d{1,2})\s*月份', link_text)
                        if num_match:
                            try:
                                month_num = int(num_match.group(1))
                            except Exception:
                                month_num = None
                        # 2) 中文数字
                        if month_num is None:
                            chinese_map = {k.replace('月',''): v for k, v in self.month_mapping.items()}
                            text_norm = link_text.replace('月份', '月')
                            for cname, cnum in chinese_map.items():
                                if f"{cname}月" in text_norm:
                                    month_num = cnum
                                    break

                        if month_num is None or '月' not in link_text:
                            continue

                        # 提取 month_id
                        match = re.search(r'/(\d+)/index\.html', link_href)
                        if match:
                            month_id = match.group(1)
                            discovered_months[month_num] = {
                                "month_name": link_text,
                                "month_id": month_id,
                                "link_text": link_text
                            }
                    except Exception as e:
                        continue

                browser.close()
                logger.info(f"发现{year}年的月份数据: {list(discovered_months.keys())}")
                return discovered_months

        except Exception as e:
            logger.error(f"动态发现浙江省数据结构时出错: {e}")
            return {}

    def check_zhejiang_data_available(self, year: int, month: int) -> bool:
        """检查浙江省数据是否可用（动态检查）"""
        logger.info(f"检查浙江省 {year}年{month}月 数据...")

        year_str = str(year)
        target_month_name = None
        for name, num in self.month_mapping.items():
            if num == month:
                target_month_name = name
                break

        if not target_month_name:
            logger.error(f"无法找到{month}月对应的中文名称")
            return False

        # 优先使用配置中的数据
        if year_str in self.zhejiang_config["years"] and target_month_name in self.zhejiang_config["years"][year_str]["months"]:
            logger.info(f"使用配置中的{year}年{target_month_name}数据")
            return self._check_zhejiang_with_config(year, month)

        # 配置中没有，尝试动态发现
        logger.info(f"配置中没有{year}年{target_month_name}数据，尝试动态发现...")
        discovered_months = self.discover_zhejiang_data_structure(year)

        if month not in discovered_months:
            logger.info(f"动态检查未发现{year}年{month}月的浙江省数据")
            return False

        # 发现了数据，检查是否可访问
        return self._check_zhejiang_with_discovered_data(year, month, discovered_months[month])

    def _check_zhejiang_with_config(self, year: int, month: int) -> bool:
        """使用配置检查浙江省数据"""
        year_str = str(year)
        target_month_name = None
        for name, num in self.month_mapping.items():
            if num == month:
                target_month_name = name
                break

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=HEADLESS,
                    slow_mo=SLOW_MO_MS,
                    args=[
                        '--no-sandbox',
                        '--disable-dev-shm-usage',
                        '--disable-blink-features=AutomationControlled'
                    ]
                )
                context = browser.new_context(
                    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
                    viewport={'width': 1920, 'height': 1080},
                    locale='zh-CN'
                )
                page = context.new_page()

                year_id = self.zhejiang_config["years"][year_str]["year_id"]
                month_id = self.zhejiang_config["years"][year_str]["months"][target_month_name]["month_id"]

                url = f"{self.zhejiang_config['base_url']}/hangzhou_customs/575609/zlbd/575612/575612/{year_id}/{month_id}/index.html"

                page.goto(url, timeout=NAV_TIMEOUT_MS)
                page.wait_for_load_state('domcontentloaded')
                page.wait_for_timeout(5000)  # 增加等待时间，确保月份页面完全加载

                # 查找文章链接
                articles = page.locator('a:has-text("进出口情况")').all()
                if not articles:
                    articles = page.locator('a:has-text("进出口")').all()

                browser.close()
                return len(articles) > 0

        except Exception as e:
            logger.error(f"检查浙江省数据时出错: {e}")
            return False

    def _check_zhejiang_with_discovered_data(self, year: int, month: int, month_data: dict) -> bool:
        """使用动态发现的数据检查浙江省数据"""
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=HEADLESS,
                    slow_mo=SLOW_MO_MS,
                    args=[
                        '--no-sandbox',
                        '--disable-dev-shm-usage',
                        '--disable-blink-features=AutomationControlled'
                    ]
                )
                context = browser.new_context(
                    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
                    viewport={'width': 1920, 'height': 1080},
                    locale='zh-CN'
                )
                page = context.new_page()

                # 使用动态发现的year_id（从配置中获取，如果没有则尝试推断）
                year_str = str(year)
                if year_str in self.zhejiang_config["years"]:
                    year_id = self.zhejiang_config["years"][year_str]["year_id"]
                else:
                    # 如果配置中没有year_id，可能需要更复杂的动态发现逻辑
                    logger.error(f"无法找到{year}年的year_id")
                    browser.close()
                    return False

                month_id = month_data["month_id"]
                url = f"{self.zhejiang_config['base_url']}/hangzhou_customs/575609/zlbd/575612/575612/{year_id}/{month_id}/index.html"

                logger.info(f"动态检查URL: {url}")

                page.goto(url, timeout=NAV_TIMEOUT_MS)
                page.wait_for_load_state('domcontentloaded')
                page.wait_for_timeout(5000)  # 增加等待时间，确保动态检查页面完全加载

                # 查找文章链接
                articles = page.locator('a:has-text("进出口情况")').all()
                if not articles:
                    articles = page.locator('a:has-text("进出口")').all()

                result = len(articles) > 0
                logger.info(f"动态检查结果: 发现{len(articles)}个相关文章")

                browser.close()
                return result

        except Exception as e:
            logger.error(f"使用动态数据检查浙江省数据时出错: {e}")
            return False

    def download_national_data(self, year: int, month: int) -> bool:
        """下载全国数据"""
        logger.info(f"下载全国 {year}年{month}月 数据...")

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=HEADLESS,
                    slow_mo=SLOW_MO_MS,
                    args=[
                        '--no-sandbox',
                        '--disable-dev-shm-usage',
                        '--disable-blink-features=AutomationControlled'
                    ]
                )
                context = browser.new_context(
                    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
                    viewport={'width': 1920, 'height': 1080},
                    locale='zh-CN'
                )
                page = context.new_page()

                # 导航到主页面
                page.goto(BASE_URL, timeout=NAV_TIMEOUT_MS)
                page.wait_for_load_state("domcontentloaded")

                # 点击年份
                year_button_selector = f"//a[contains(text(), '{year}')]"
                page.wait_for_selector(year_button_selector, timeout=20000).click()
                time.sleep(3)

                # 点击月份链接
                month_text = f"{month}月"
                month_link_selector = (
                    "//tr[contains(., '进出口商品收发货人所在地总值表')]//a[text()='{}']".format(
                        month_text
                    )
                )
                month_link_element = page.wait_for_selector(month_link_selector, timeout=SELECTOR_TIMEOUT_MS)

                # 点击打开新页签
                with context.expect_page() as new_page_info:
                    month_link_element.click()
                detail_page = new_page_info.value if new_page_info.value else page
                detail_page.wait_for_load_state("domcontentloaded", timeout=NAV_TIMEOUT_MS)

                # 提取表格数据
                table_container_selector = "div.easysite-news-text"
                try:
                    detail_page.wait_for_selector(table_container_selector, timeout=SELECTOR_TIMEOUT_MS)
                except Exception:
                    table_container_selector = "body"
                    detail_page.wait_for_selector(table_container_selector, timeout=SELECTOR_TIMEOUT_MS)

                table_html = detail_page.locator(table_container_selector).inner_html()

                # 解析表格
                try:
                    df = parse_table_html_to_df(table_html)
                except Exception:
                    df_list = pd.read_html(table_html, header=0)
                    if not df_list:
                        raise ValueError("未解析到表格")
                    df = df_list[0]

                # 保存数据
                file_path = os.path.join(RAW_DATA_PATH, month_file_name(year, month))
                save_raw_df(df, file_path)
                logger.info(f"全国数据已保存至: {file_path}")

                try:
                    detail_page.close()
                except Exception:
                    pass

                browser.close()
                return True

        except Exception as e:
            logger.error(f"下载全国数据失败: {e}")
            return False

    def download_zhejiang_data(self, year: int, month: int) -> bool:
        """下载浙江省数据（支持动态发现）"""
        logger.info(f"下载浙江省 {year}年{month}月 数据...")

        year_str = str(year)
        target_month_name = None
        for name, num in self.month_mapping.items():
            if num == month:
                target_month_name = name
                break

        if not target_month_name:
            logger.error(f"无法找到{month}月对应的中文名称")
            return False

        # 优先使用配置中的数据
        if year_str in self.zhejiang_config["years"] and target_month_name in self.zhejiang_config["years"][year_str]["months"]:
            logger.info(f"使用配置下载{year}年{target_month_name}数据")
            return self._download_zhejiang_with_config(year, month)

        # 配置中没有，尝试动态发现
        logger.info(f"配置中没有{year}年{target_month_name}数据，尝试动态下载...")
        discovered_months = self.discover_zhejiang_data_structure(year)

        if month not in discovered_months:
            logger.error(f"动态检查未发现{year}年{month}月的浙江省数据")
            return False

        # 使用动态发现的数据下载
        return self._download_zhejiang_with_discovered_data(year, month, discovered_months[month])

    def _download_zhejiang_with_config(self, year: int, month: int) -> bool:
        """使用配置下载浙江省数据"""
        year_str = str(year)
        target_month_name = None
        for name, num in self.month_mapping.items():
            if num == month:
                target_month_name = name
                break

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=HEADLESS,
                    slow_mo=SLOW_MO_MS,
                    args=[
                        '--no-sandbox',
                        '--disable-dev-shm-usage',
                        '--disable-blink-features=AutomationControlled'
                    ]
                )
                context = browser.new_context(
                    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
                    viewport={'width': 1920, 'height': 1080},
                    locale='zh-CN'
                )
                page = context.new_page()

                year_id = self.zhejiang_config["years"][year_str]["year_id"]
                month_id = self.zhejiang_config["years"][year_str]["months"][target_month_name]["month_id"]

                url = f"{self.zhejiang_config['base_url']}/hangzhou_customs/575609/zlbd/575612/575612/{year_id}/{month_id}/index.html"

                page.goto(url, timeout=NAV_TIMEOUT_MS)
                page.wait_for_load_state('domcontentloaded')
                page.wait_for_timeout(6000)  # 下载时需要更长等待时间，确保页面完全加载

                # 查找文章链接
                articles = page.locator('a:has-text("进出口情况")').all()
                if not articles:
                    articles = page.locator('a:has-text("进出口")').all()

                if not articles:
                    logger.error(f"{year}年{month}月未找到相关文章")
                    browser.close()
                    return False

                # 处理第一篇文章
                article = articles[0]
                with page.expect_popup() as popup_info:
                    article.click()
                article_page = popup_info.value

                article_page.wait_for_load_state('domcontentloaded')
                article_page.wait_for_timeout(6000)  # 增加文章页面等待时间

                # 查找Excel链接
                excel_links = []

                download_selectors = [
                    'a[href*=".xlsx"]:has-text("下载")',
                    'a[href*=".xlsx"]:has-text("附件")',
                    'a[href*=".xlsx"]:has-text("表格")',
                    'a[href*=".xlsx"]:has-text("数据")',
                    'a[href*=".xls"]:has-text("下载")',
                    'a[href*=".xls"]:has-text("附件")',
                ]

                for selector in download_selectors:
                    try:
                        links = article_page.locator(selector).all()
                        if links:
                            excel_links = links
                            break
                    except:
                        continue

                if not excel_links:
                    all_excel_links = article_page.locator('a[href*=".xlsx"], a[href*=".xls"]').all()
                    excel_links = [
                        link for link in all_excel_links
                        if not any(keyword in (link.get_attribute('href') or '').lower()
                                  for keyword in ['share', 'qzone', 'weibo', 'wechat', '分享'])
                        and not any(keyword in (link.inner_text() or '').lower()
                                   for keyword in ['分享', 'share', '转发'])
                    ]

                if excel_links:
                    excel_link = excel_links[0]

                    with article_page.expect_download() as download_info:
                        excel_link.click()
                    download = download_info.value

                    # 临时保存Excel文件
                    temp_excel = f"temp_{year}_{month}.xlsx"
                    download.save_as(temp_excel)

                    try:
                        # 读取十一地市sheet并转换为CSV格式
                        df = pd.read_excel(temp_excel, sheet_name="十一地市")

                        if not df.empty:
                            # 清理数据
                            df_clean = df.dropna(subset=['项目'])
                            df_clean = df_clean.rename(columns={'项目': '收发货人所在地'})

                            # 单位转换处理
                            if (year == 2024 and month >= 7) or year >= 2025:
                                value_cols = ['当期进出口', '当期进口', '当期出口']
                                for col in value_cols:
                                    if col in df_clean.columns:
                                        df_clean[col] = df_clean[col] * 10000

                            # 保存为CSV
                            zj_file_name = f"浙江省-{year}-{month:02d}.csv"
                            csv_file = os.path.join(RAW_DATA_PATH, zj_file_name)
                            save_raw_df(df_clean, csv_file)
                            logger.info(f"浙江省数据已保存至: {csv_file}")

                            return True

                    except Exception as e:
                        logger.error(f"处理浙江省数据失败: {e}")
                        return False
                    finally:
                        if os.path.exists(temp_excel):
                            os.remove(temp_excel)

                article_page.close()
                browser.close()
                return False

        except Exception as e:
            logger.error(f"下载浙江省数据失败: {e}")
            return False

    def _download_zhejiang_with_discovered_data(self, year: int, month: int, month_data: dict) -> bool:
        """使用动态发现的数据下载浙江省数据"""
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(
                    headless=HEADLESS,
                    slow_mo=SLOW_MO_MS,
                    args=[
                        '--no-sandbox',
                        '--disable-dev-shm-usage',
                        '--disable-blink-features=AutomationControlled'
                    ]
                )
                context = browser.new_context(
                    user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
                    viewport={'width': 1920, 'height': 1080},
                    locale='zh-CN'
                )
                page = context.new_page()

                # 使用动态发现的数据
                year_str = str(year)
                if year_str in self.zhejiang_config["years"]:
                    year_id = self.zhejiang_config["years"][year_str]["year_id"]
                else:
                    logger.error(f"无法找到{year}年的year_id，无法下载")
                    browser.close()
                    return False

                month_id = month_data["month_id"]
                url = f"{self.zhejiang_config['base_url']}/hangzhou_customs/575609/zlbd/575612/575612/{year_id}/{month_id}/index.html"

                logger.info(f"使用动态发现的数据下载，URL: {url}")

                page.goto(url, timeout=NAV_TIMEOUT_MS)
                page.wait_for_load_state('domcontentloaded')
                page.wait_for_timeout(6000)  # 动态下载时需要更长等待时间，确保页面完全加载

                # 查找文章链接
                articles = page.locator('a:has-text("进出口情况")').all()
                if not articles:
                    articles = page.locator('a:has-text("进出口")').all()

                if not articles:
                    logger.error(f"{year}年{month}月未找到相关文章")
                    browser.close()
                    return False

                # 处理第一篇文章
                article = articles[0]
                with page.expect_popup() as popup_info:
                    article.click()
                article_page = popup_info.value

                article_page.wait_for_load_state('domcontentloaded')
                article_page.wait_for_timeout(6000)  # 增加文章页面等待时间

                # 查找Excel链接（与配置版本相同的逻辑）
                excel_links = []

                download_selectors = [
                    'a[href*=".xlsx"]:has-text("下载")',
                    'a[href*=".xlsx"]:has-text("附件")',
                    'a[href*=".xlsx"]:has-text("表格")',
                    'a[href*=".xlsx"]:has-text("数据")',
                    'a[href*=".xls"]:has-text("下载")',
                    'a[href*=".xls"]:has-text("附件")',
                ]

                for selector in download_selectors:
                    try:
                        links = article_page.locator(selector).all()
                        if links:
                            excel_links = links
                            break
                    except:
                        continue

                if not excel_links:
                    all_excel_links = article_page.locator('a[href*=".xlsx"], a[href*=".xls"]').all()
                    excel_links = [
                        link for link in all_excel_links
                        if not any(keyword in (link.get_attribute('href') or '').lower()
                                  for keyword in ['share', 'qzone', 'weibo', 'wechat', '分享'])
                        and not any(keyword in (link.inner_text() or '').lower()
                                   for keyword in ['分享', 'share', '转发'])
                    ]

                if excel_links:
                    excel_link = excel_links[0]

                    with article_page.expect_download() as download_info:
                        excel_link.click()
                    download = download_info.value

                    # 临时保存Excel文件
                    temp_excel = f"temp_dynamic_{year}_{month}.xlsx"
                    download.save_as(temp_excel)

                    try:
                        # 读取十一地市sheet并转换为CSV格式
                        df = pd.read_excel(temp_excel, sheet_name="十一地市")

                        if not df.empty:
                            # 清理数据
                            df_clean = df.dropna(subset=['项目'])
                            df_clean = df_clean.rename(columns={'项目': '收发货人所在地'})

                            # 单位转换处理
                            if (year == 2024 and month >= 7) or year >= 2025:
                                value_cols = ['当期进出口', '当期进口', '当期出口']
                                for col in value_cols:
                                    if col in df_clean.columns:
                                        df_clean[col] = df_clean[col] * 10000

                            # 保存为CSV
                            zj_file_name = f"浙江省-{year}-{month:02d}.csv"
                            csv_file = os.path.join(RAW_DATA_PATH, zj_file_name)
                            save_raw_df(df_clean, csv_file)
                            logger.info(f"浙江省数据（动态发现）已保存至: {csv_file}")

                            return True

                    except Exception as e:
                        logger.error(f"处理动态发现的浙江省数据失败: {e}")
                        return False
                    finally:
                        if os.path.exists(temp_excel):
                            os.remove(temp_excel)

                article_page.close()
                browser.close()
                return False

        except Exception as e:
            logger.error(f"使用动态数据下载浙江省数据失败: {e}")
    def download_zhejiang_auto_newest(self) -> List[Tuple[int, int]]:
        """
        自动识别杭州海关“统计数据”栏目下最新年份与月份，下载浙江省缺失月份数据。
        返回成功下载的 (year, month) 列表。
        """
        ensure_raw_data_dir()

        def abs_url(base: str, href: str) -> str:
            href = href or ""
            return href if href.startswith("http") else f"{base}{href}"

        base = "http://hangzhou.customs.gov.cn"
        entry_url = f"{base}/hangzhou_customs/575609/zlbd/575612/575612/6430241/6430315/index.html"

        downloaded: List[Tuple[int, int]] = []

        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=HEADLESS,
                slow_mo=SLOW_MO_MS,
                args=[
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-blink-features=AutomationControlled",
                ],
            )
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
                viewport={"width": 1920, "height": 1080},
                locale="zh-CN",
            )
            page = context.new_page()

            # 进入入口页并等待“统计数据”导航
            for attempt in range(1, NAV_MAX_RETRIES + 1):
                try:
                    page.goto(entry_url, timeout=NAV_TIMEOUT_MS)
                    page.wait_for_load_state("domcontentloaded")
                    page.get_by_role("navigation", name="统计数据") \
                        .get_by_role("link").first.wait_for(state="visible", timeout=SELECTOR_TIMEOUT_MS)
                    break
                except Exception as nav_err:
                    if attempt == NAV_MAX_RETRIES:
                        raise nav_err
                    logger.info(f"导航失败，重试 {attempt}/{NAV_MAX_RETRIES - 1} ...")
                    time.sleep(NAV_RETRY_DELAY_SEC)

            years_locator = page.get_by_role("navigation", name="统计数据").get_by_role("link")
            latest_year_link = years_locator.first
            latest_year_text = latest_year_link.inner_text().strip()
            m = re.search(r"(\d{4})年", latest_year_text)
            if not m:
                browser.close()
                raise RuntimeError("未能识别到年份文本")
            latest_year = int(m.group(1))
            latest_year_url = abs_url(base, latest_year_link.get_attribute("href") or "")

            page.goto(latest_year_url, timeout=NAV_TIMEOUT_MS)
            page.wait_for_load_state("domcontentloaded")
            page.wait_for_timeout(1500)

            months_locator = page.get_by_role("navigation", name=f"{latest_year}年").get_by_role("link")
            count = months_locator.count()
            months: List[Tuple[int, str, str]] = []
            for i in range(count):
                link = months_locator.nth(i)
                name = link.inner_text().strip()
                if name not in self.month_mapping:
                    continue
                url = abs_url(base, link.get_attribute("href") or "")
                months.append((self.month_mapping[name], name, url))

            months.sort(key=lambda x: x[0])
            logger.info("杭州海关 {0} 年可用月份：{1}".format(latest_year, ", ".join([f"{m}{n:02d}" for n, m, _ in months])))

            for month_num, _month_cn, month_url in months:
                csv_name = f"浙江省-{latest_year}-{month_num:02d}.csv"
                csv_path = os.path.join(RAW_DATA_PATH, csv_name)
                if os.path.exists(csv_path):
                    continue

                try:
                    page.goto(month_url, timeout=NAV_TIMEOUT_MS)
                    page.wait_for_load_state("domcontentloaded")
                    page.wait_for_timeout(1500)

                    # 月份页首条“进出口情况”文章
                    candidates = page.locator("a:has-text('浙江省进出口情况')").all()
                    if not candidates:
                        candidates = page.locator("a:has-text('进出口情况')").all()
                    if not candidates:
                        candidates = page.locator("a:has-text('进出口')").all()
                    if not candidates:
                        logger.info(f"{latest_year}年{month_num}月未找到文章")
                        continue

                    article_url = abs_url(base, candidates[0].get_attribute("href") or "")
                    article_page = context.new_page()
                    article_page.goto(article_url, timeout=NAV_TIMEOUT_MS)
                    article_page.wait_for_load_state("domcontentloaded")
                    article_page.wait_for_timeout(1500)

                    # Excel 链接
                    links = article_page.locator('a[href$=".xlsx"], a[href$=".xls"]').all()
                    cleaned = []
                    for lk in links:
                        href = (lk.get_attribute("href") or "").lower()
                        txt = (lk.inner_text() or "").lower()
                        if any(x in href for x in ["share", "qzone", "weibo", "wechat"]) or \
                           any(x in txt for x in ["分享", "share", "转发"]):
                            continue
                        cleaned.append(lk)

                    if not cleaned:
                        article_page.close()
                        continue

                    with article_page.expect_download() as dl_info:
                        cleaned[0].click()
                    dl = dl_info.value

                    temp_excel = f"temp_{latest_year}_{month_num}.xlsx"
                    dl.save_as(temp_excel)

                    try:
                        df = pd.read_excel(temp_excel, sheet_name="十一地市")
                        if not df.empty:
                            df_clean = df.dropna(subset=["项目"]).rename(columns={"项目": "收发货人所在地"})
                            if (latest_year == 2024 and month_num >= 7) or latest_year >= 2025:
                                for col in ["当期进出口", "当期进口", "当期出口"]:
                                    if col in df_clean.columns:
                                        df_clean[col] = df_clean[col] * 10000
                            save_raw_df(df_clean, csv_path)
                            downloaded.append((latest_year, month_num))
                            logger.info(f"已保存浙江省数据：{csv_path}")
                    finally:
                        if os.path.exists(temp_excel):
                            os.remove(temp_excel)
                        article_page.close()

                    time.sleep(random.uniform(2.0, 4.0))
                except Exception as e:
                    logger.error(f"{latest_year}年{month_num}月下载失败：{e}")
                    continue

            browser.close()
        return downloaded

    def update_excel_file(self) -> bool:
        """更新Excel文件"""
        try:
            logger.info("开始更新Excel文件...")
            master_df = consolidate_with_yoy()
            export_to_excel_by_location(master_df)
            logger.info(f"Excel文件已更新: {OUTPUT_FILENAME}")
            return True
        except Exception as e:
            logger.error(f"更新Excel文件失败: {e}")
            return False

    def check_and_update(self) -> dict:
        """检查并更新数据的主函数"""
        logger.info("开始检查数据更新...")

        # 获取本地数据状态
        national_months, zj_months = self.get_local_data_status()

        # 计算需要检查的月份
        national_next = self.get_next_expected_months(national_months)
        zj_next = self.get_next_expected_months(zj_months)

        logger.info(f"需要检查的全国数据月份: {national_next}")
        logger.info(f"需要检查的浙江省数据月份: {zj_next}")

        new_national = []
        new_zj = []

        # 检查全国数据
        for year, month in national_next:
            if self.check_national_data_available(year, month):
                if self.download_national_data(year, month):
                    new_national.append((year, month))
                    logger.info(f"✅ 成功下载全国 {year}年{month}月 数据")
                else:
                    logger.error(f"❌ 下载全国 {year}年{month}月 数据失败")
            else:
                logger.info(f"全国 {year}年{month}月 数据暂未发布")

        # 检查浙江省数据
        # 使用自动识别最新月份的逻辑自动补齐缺失月份
        auto_new = self.download_zhejiang_auto_newest()
        if auto_new:
            for y, m in auto_new:
                if (y, m) not in new_zj:
                    new_zj.append((y, m))
                    logger.info(f"✅ 成功下载浙江省 {y}年{m}月 数据")
        else:
            logger.info("浙江省本年度暂无新增可下载月份或均已存在")

        # 如果有新数据，更新Excel文件
        excel_updated = False
        if new_national or new_zj:
            excel_updated = self.update_excel_file()

        # 返回更新结果
        result = {
            "has_updates": bool(new_national or new_zj),
            "new_national_months": new_national,
            "new_zj_months": new_zj,
            "excel_updated": excel_updated,
            "message": ""
        }

        if result["has_updates"]:
            msg_parts = []
            if new_national:
                national_str = ", ".join([f"{y}年{m}月" for y, m in new_national])
                msg_parts.append(f"全国数据: {national_str}")
            if new_zj:
                zj_str = ", ".join([f"{y}年{m}月" for y, m in new_zj])
                msg_parts.append(f"浙江省数据: {zj_str}")

            result["message"] = f"发现新数据并已更新: {'; '.join(msg_parts)}"
        else:
            result["message"] = "当前已是最新数据，无可更新数据"

        logger.info(result["message"])
        return result


def main():
    """命令行运行入口"""
    try:
        updater = DataUpdateChecker()
        result = updater.check_and_update()

        print("\n" + "="*50)
        print("数据更新结果")
        print("="*50)
        print(f"状态: {'有更新' if result['has_updates'] else '无更新'}")
        print(f"消息: {result['message']}")

        if result['new_national_months']:
            print(f"新增全国数据: {result['new_national_months']}")
        if result['new_zj_months']:
            print(f"新增浙江省数据: {result['new_zj_months']}")

        print(f"Excel文件更新: {'是' if result['excel_updated'] else '否'}")
        print("="*50)

        # 确保输出被刷新
        sys.stdout.flush()

    except Exception as e:
        print(f"更新器运行出错: {e}")
        logger.error(f"更新器运行出错: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
