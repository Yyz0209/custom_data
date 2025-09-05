import os
import time
import json
import re
from datetime import datetime

from playwright.sync_api import sync_playwright

from config import (
    BASE_URL,
    RAW_DATA_PATH,
    USER_AGENT,
    HEADLESS,
    NAV_TIMEOUT_MS,
    SELECTOR_TIMEOUT_MS,
    NAV_MAX_RETRIES,
    NAV_RETRY_DELAY_SEC,
    SLOW_MO_MS,
)
import random
import pandas as pd
from data_utils import ensure_raw_data_dir, month_file_name, parse_table_html_to_df, save_raw_df


def scrape_raw_data_to_csv() -> None:
    """
    下载自 2024 年起所有可见月份的原始表格，并原样保存为 CSV 到 `raw_csv_data/`。
    """
    ensure_raw_data_dir()
    print(f"原始数据将保存在 '{RAW_DATA_PATH}' 文件夹中。")

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

        try:
            current_year = datetime.now().year
            for year in range(2024, current_year + 1):
                try:
                    print(f"导航到主页面，准备处理年份: {year}")
                    # 导航加入重试，适配慢网/代理
                    for attempt in range(1, NAV_MAX_RETRIES + 1):
                        try:
                            page.goto(BASE_URL, timeout=NAV_TIMEOUT_MS)
                            page.wait_for_load_state("domcontentloaded")
                            page.wait_for_selector("//div[@class='customs-foot']", timeout=SELECTOR_TIMEOUT_MS)
                            break
                        except Exception as nav_err:
                            if attempt == NAV_MAX_RETRIES:
                                raise nav_err
                            print(f"    导航失败，重试 {attempt}/{NAV_MAX_RETRIES - 1} ...")
                            time.sleep(NAV_RETRY_DELAY_SEC)
                    print("    主页面加载成功！")

                    year_button_selector = f"//a[contains(text(), '{year}')]"
                    page.wait_for_selector(year_button_selector, timeout=20000).click()
                    time.sleep(3)

                    table_row_selector = "//tr[contains(., '进出口商品收发货人所在地总值表')]"
                    row = page.wait_for_selector(table_row_selector, timeout=20000)

                    month_texts = [
                        link.inner_text()
                        for link in row.query_selector_all("a")
                        if link.get_attribute("href")
                    ]
                    print(f"    找到 {year} 年可用的月份: {month_texts}")

                    for month_text in month_texts:
                        if "月" not in month_text:
                            continue
                        if not month_text.replace("月", "").strip().isdigit():
                            continue

                        month_num = int(month_text.replace("月", "").strip())
                        
                        # 检查文件是否已存在，如果存在则跳过
                        file_path = os.path.join(RAW_DATA_PATH, month_file_name(year, month_num))
                        if os.path.exists(file_path):
                            print(f"  ✅ {year}年{month_num}月全国数据已存在，跳过")
                            continue
                        
                        print(f"  📥 下载 {year}年{month_num}月 全国数据...")

                        month_link_selector = (
                            "//tr[contains(., '进出口商品收发货人所在地总值表')]//a[text()='{}']".format(
                                month_text
                            )
                        )
                        month_link_element = page.wait_for_selector(month_link_selector, timeout=SELECTOR_TIMEOUT_MS)

                        # 点击打开新页签；个别情况下目标会在同页打开，因此做两种分支处理
                        with context.expect_page() as new_page_info:
                            month_link_element.click()
                        detail_page = new_page_info.value if new_page_info.value else page
                        detail_page.wait_for_load_state("domcontentloaded", timeout=NAV_TIMEOUT_MS)

                        table_container_selector = "div.easysite-news-text"
                        try:
                            detail_page.wait_for_selector(table_container_selector, timeout=SELECTOR_TIMEOUT_MS)
                        except Exception:
                            # 有时候新闻正文容器 class 变化或被包裹，退而求其次直接抓取 body
                            table_container_selector = "body"
                            detail_page.wait_for_selector(table_container_selector, timeout=SELECTOR_TIMEOUT_MS)
                        table_html = detail_page.locator(table_container_selector).inner_html()

                        # 解析表格：优先多级表头；若失败，降级 header=0
                        try:
                            df = parse_table_html_to_df(table_html)
                        except Exception:
                            import pandas as pd
                            df_list = pd.read_html(table_html, header=0)
                            if not df_list:
                                raise ValueError("未解析到表格")
                            df = df_list[0]
                        file_path = os.path.join(RAW_DATA_PATH, month_file_name(year, month_num))
                        save_raw_df(df, file_path)
                        print(f"    原始数据已保存至: {file_path}")

                        try:
                            detail_page.close()
                        except Exception:
                            pass
                        # 随机抖动，降低触发频率
                        time.sleep(random.uniform(2.0, 5.0))

                except Exception as exc:
                    print(f"处理年份 {year} 时出错: {exc}")
                    continue

            print("\n所有原始数据爬取完成！")
        finally:
            print("任务结束，正在关闭浏览器...")
            browser.close()


def download_zhejiang_data():
    """动态识别杭州海关‘统计数据’栏目下最新年份/月份，下载浙江省数据为CSV到 raw_csv_data。
    逻辑：进入统计数据 → 取最新年份 → 枚举该年份下所有出现的月份（最后一个即最新）→ 进入月份页首条文章 → 下载 .xlsx → 读取“十一地市”表并保存。
    已下载的“浙江省-YYYY-MM.csv”将自动跳过。
    """

    month_mapping = {
        "一月": 1, "二月": 2, "三月": 3, "四月": 4, "五月": 5, "六月": 6,
        "七月": 7, "八月": 8, "九月": 9, "十月": 10, "十一月": 11, "十二月": 12
    }

    def abs_url(base: str, href: str) -> str:
        href = href or ""
        return href if href.startswith("http") else f"{base}{href}"

    base = "http://hangzhou.customs.gov.cn"
    # 入口页：任意一个年份页面即可展示左侧“统计数据”年份导航；使用当前可用的 2025 一月页
    entry_url = f"{base}/hangzhou_customs/575609/zlbd/575612/575612/6430241/6430315/index.html"

    print("\n🚀 开始下载浙江省数据（自动识别最新月份）...")
    ensure_raw_data_dir()

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

        try:
            # 进入入口页
            for attempt in range(1, NAV_MAX_RETRIES + 1):
                try:
                    page.goto(entry_url, timeout=NAV_TIMEOUT_MS)
                    page.wait_for_load_state("domcontentloaded")
                    # 等待左侧“统计数据”导航出现
                    page.get_by_role("navigation", name="统计数据").get_by_role("link").first.wait_for(state="visible", timeout=SELECTOR_TIMEOUT_MS)
                    break
                except Exception as nav_err:
                    if attempt == NAV_MAX_RETRIES:
                        raise nav_err
                    print(f"    导航失败，重试 {attempt}/{NAV_MAX_RETRIES - 1} ...")
                    time.sleep(NAV_RETRY_DELAY_SEC)
            print("    已打开统计数据栏目页面")

            # 统计数据 → 最新年份（第一个链接）
            years_locator = page.get_by_role("navigation", name="统计数据").get_by_role("link")
            latest_year_link = years_locator.first
            latest_year_text = latest_year_link.inner_text().strip()
            m = re.search(r"(\d{4})年", latest_year_text)
            if not m:
                raise RuntimeError("未能识别到年份文本")
            latest_year = int(m.group(1))

            latest_year_href = latest_year_link.get_attribute("href")
            latest_year_url = abs_url(base, latest_year_href)
            print(f"    最新年份：{latest_year} -> {latest_year_url}")

            # 直接导航到该年份页面（而非点击，避免新开页/同页不一致问题）
            page.goto(latest_year_url, timeout=NAV_TIMEOUT_MS)
            page.wait_for_load_state("domcontentloaded")
            page.wait_for_timeout(1500)

            # 在“YYYY年”导航中读取所有月份链接（已发布的月份才会出现）
            months_locator = page.get_by_role("navigation", name=f"{latest_year}年").get_by_role("link")
            month_count = months_locator.count()
            if month_count == 0:
                raise RuntimeError("未找到月份导航链接")

            months = []  # [(month_num, month_name_cn, month_url)]
            for i in range(month_count):
                link = months_locator.nth(i)
                name = link.inner_text().strip()
                if name not in month_mapping:
                    continue
                href = link.get_attribute("href") or ""
                url = abs_url(base, href)
                months.append((month_mapping[name], name, url))

            # 按月份升序，便于补齐缺失月份；最后一个即“最新”
            months.sort(key=lambda x: x[0])
            if not months:
                raise RuntimeError("月份解析为空")

            print("    发现月份：" + ", ".join([f"{mcn}({mnum:02d})" for mnum, mcn, _ in months]))

            for month_num, month_cn, month_url in months:
                csv_name = f"浙江省-{latest_year}-{month_num:02d}.csv"
                csv_path = os.path.join(RAW_DATA_PATH, csv_name)
                if os.path.exists(csv_path):
                    print(f"✅ {latest_year}年{month_num}月已存在，跳过")
                    continue

                print(f"📥 处理 {latest_year}年{month_num}月 → {month_url}")
                try:
                    page.goto(month_url, timeout=NAV_TIMEOUT_MS)
                    page.wait_for_load_state("domcontentloaded")
                    page.wait_for_timeout(1500)

                    # 月份页首条“进出口情况”文章链接
                    article_link = None
                    # 优先更精确
                    candidates = page.locator("a:has-text('浙江省进出口情况')").all()
                    if not candidates:
                        candidates = page.locator("a:has-text('进出口情况')").all()
                    if not candidates:
                        candidates = page.locator("a:has-text('进出口')").all()
                    if candidates:
                        article_link = candidates[0]

                    if not article_link:
                        print(f"    ❌ {latest_year}年{month_num}月未找到文章链接")
                        continue

                    article_href = article_link.get_attribute("href") or ""
                    article_url = abs_url(base, article_href)
                    print(f"    打开文章：{article_url}")

                    article_page = context.new_page()
                    article_page.goto(article_url, timeout=NAV_TIMEOUT_MS)
                    article_page.wait_for_load_state("domcontentloaded")
                    article_page.wait_for_timeout(1500)

                    # 查找 Excel 下载链接，优先 .xlsx
                    def find_excel_links() -> list:
                        # 排除分享/社交链接
                        links = article_page.locator('a[href$=".xlsx"], a[href$=".xls"]').all()
                        cleaned = []
                        for lk in links:
                            href = (lk.get_attribute("href") or "").lower()
                            txt = (lk.inner_text() or "").lower()
                            if any(x in href for x in ["share", "qzone", "weibo", "wechat"]) or \
                               any(x in txt for x in ["分享", "share", "转发"]):
                                continue
                            cleaned.append(lk)
                        return cleaned

                    excel_links = find_excel_links()
                    if not excel_links:
                        print("    ❌ 未发现Excel下载链接")
                        article_page.close()
                        continue

                    # 取第一个 Excel 链接
                    excel_link = excel_links[0]
                    link_text = excel_link.inner_text() or "(无标题)"
                    link_href = excel_link.get_attribute("href") or ""
                    print(f"    � 准备下载：{link_text} -> {link_href}")

                    with article_page.expect_download() as dl_info:
                        excel_link.click()
                    dl = dl_info.value

                    temp_excel = f"temp_{latest_year}_{month_num}.xlsx"
                    dl.save_as(temp_excel)

                    try:
                        df = pd.read_excel(temp_excel, sheet_name="十一地市")
                        if not df.empty:
                            df_clean = df.dropna(subset=["项目"]).rename(columns={"项目": "收发货人所在地"})
                            # 单位转换规则保持与原逻辑一致
                            if (latest_year == 2024 and month_num >= 7) or latest_year >= 2025:
                                for col in ["当期进出口", "当期进口", "当期出口"]:
                                    if col in df_clean.columns:
                                        df_clean[col] = df_clean[col] * 10000
                            save_raw_df(df_clean, csv_path)
                            print(f"✅ 已保存：{csv_path}")
                    except Exception as e:
                        print(f"    ❌ 数据处理失败：{e}")
                    finally:
                        if os.path.exists(temp_excel):
                            os.remove(temp_excel)
                        article_page.close()

                    time.sleep(random.uniform(2.0, 4.0))
                except Exception as e:
                    print(f"❌ {latest_year}年{month_num}月下载失败：{e}")
                    continue

            print("✅ 浙江省数据下载完成")
        finally:
            browser.close()


def main():
    """主函数"""
    print("🚀 海关数据下载器")
    print("=" * 50)
    print("💡 提示：脚本会自动跳过已下载的文件")
    print("选择下载选项:")
    print("1. 下载全国数据")
    print("2. 下载浙江省数据") 
    print("3. 下载所有数据")
    print("0. 退出")
    
    try:
        choice = input("\n请选择 (0-3，默认3): ").strip()
        
        if choice == "1":
            scrape_raw_data_to_csv()
        elif choice == "2":
            download_zhejiang_data()
        elif choice == "3" or choice == "":
            scrape_raw_data_to_csv()
            download_zhejiang_data()
        elif choice == "0":
            print("👋 退出")
            return
        else:
            print("❌ 无效选择")
            
    except KeyboardInterrupt:
        print("\n⚠️  用户中断")
    except Exception as e:
        print(f"\n❌ 运行错误: {e}")


if __name__ == "__main__":
    main()
