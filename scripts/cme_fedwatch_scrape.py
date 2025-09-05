import json
import os
import sys
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

URL = "https://www.cmegroup.com/markets/interest-rates/cme-fedwatch-tool.html"

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"


def open_with_fallback(p):
    proxy = _get_proxy_from_env()
    candidates = [
        ("chromium", {"headless": True}),
        ("chromium", {"headless": True, "args": ["--disable-http2", "--disable-quic"]}),
        ("firefox", {"headless": True}),
        ("webkit", {"headless": True}),
    ]
    errs = []
    for name, opts in candidates:
        browser = None
        try:
            launch_opts = opts.copy()
            if proxy:
                launch_opts["proxy"] = proxy
            browser = getattr(p, name).launch(**launch_opts)
            ctx = browser.new_context(ignore_https_errors=True, user_agent=UA, locale="en-US")
            page = ctx.new_page()
            page.goto(URL, wait_until="domcontentloaded", timeout=120000)
            return browser, ctx, page
        except Exception as e:
            errs.append((name, str(e)))
            try:
                if browser:
                    browser.close()
            except Exception:
                pass
    raise RuntimeError("All browser fallbacks failed: " + "; ".join([f"{n}: {msg}" for n, msg in errs]))




def _get_proxy_from_env():
    # Priority: PLAYWRIGHT_PROXY > HTTPS_PROXY > HTTP_PROXY
    for k in ("PLAYWRIGHT_PROXY", "HTTPS_PROXY", "HTTP_PROXY", "https_proxy", "http_proxy"):
        v = os.environ.get(k)
        if v:
            return {"server": v}
    return None

EVAL_PARSE_TABLE = """
(el) => {
  const t = el;
  const txt = n => (n.innerText || n.textContent || '').trim().replace(/\s+/g, ' ');
  const rows = Array.from(t.querySelectorAll('tbody tr')).map(tr => Array.from(tr.children).map(td => txt(td)));
  let header = [];
  if (rows.length) {
    const first = rows[0];
    const looksHeader = first.some(c => /meeting|target rate|longer run|bps|prob/i.test(c));
    if (looksHeader) header = first; else header = Array.from(t.querySelectorAll('thead th')).map(txt);
    if (!header.length) header = first.map((_, i) => `col_${i}`);
  }
  const body = rows.length ? (header.length && rows[0] === header ? rows.slice(1) : rows.slice(1)) : [];
  const data = body.map(r => {
    const o = {};
    r.forEach((v, i) => { o[header[i] || `col_${i}`] = v; });
    return o;
  });
  return { header, rows: body, data };
}
"""


def accept_cookies(page):
    # Try common accept buttons (Chinese + English variants)
    for text in ["接受所有 Cookie", "Accept All Cookies", "Accept All", "同意", "同意全部"]:
        try:
            page.get_by_role("button", name=text, exact=False).first.click(timeout=1500)
            return
        except PWTimeout:
            pass
        except Exception:
            pass


def get_fedwatch_frame(page, timeout_ms: int = 60000):
    """Return the FedWatch iframe Frame. Robust to extra invisible iframes."""
    page.wait_for_load_state("domcontentloaded")
    # 1) Prefer the main tool iframe whose id starts with cmeIframe
    try:
        el = page.locator('iframe[id^="cmeIframe"]').first
        el.wait_for(state="attached", timeout=20000)
        fr = el.content_frame()
        if fr:
            return fr
    except Exception:
        pass

    # 2) Fallback: poll frames for one that contains the text "FedWatch Tool"
    import time
    deadline = time.time() + (timeout_ms / 1000.0)
    last_err = None
    while time.time() < deadline:
        try:
            for fr in page.frames:
                try:
                    if fr.locator("text=FedWatch Tool").count() > 0:
                        return fr
                except Exception:
                    continue
        except Exception as e:
            last_err = e
        time.sleep(0.5)
    raise TimeoutError(f"FedWatch iframe not found within {timeout_ms} ms; last error: {last_err}")


def parse_probabilities(frame):
    frame.get_by_role("link", name="Probabilities").click()
    frame.get_by_text("Conditional Meeting Probabilities", exact=False).wait_for(timeout=30000)
    table = frame.locator("table:has-text('Conditional Meeting Probabilities')").first
    return table.evaluate(EVAL_PARSE_TABLE)


def parse_dotplot_table(frame):
    import re
    # Ensure we are on the Dot Plot tab first
    try:
        frame.get_by_role("link", name=re.compile("Dot Plot", re.I)).click(timeout=5000)
    except Exception:
        pass
    # Switch to Table sub-tab (fallback across role types)
    clicked = False
    for role in ("tab", "link", "button"):
        try:
            frame.get_by_role(role, name=re.compile("Table", re.I)).click(timeout=4000)
            clicked = True
            break
        except Exception:
            continue
    if not clicked:
        # best-effort: click by text
        try:
            frame.get_by_text("Table", exact=False).click(timeout=3000)
        except Exception:
            pass

    # Try to force render by toggling Chart -> Table if needed
    try:
        # Scroll table tab into view and click
        try:
            tlink = frame.get_by_role("link", name=re.compile("Table", re.I)).first
            tlink.scroll_into_view_if_needed(timeout=2000)
            tlink.click(timeout=4000)
        except Exception:
            pass
        # Toggle a few times if not visible yet
        for _ in range(3):
            try:
                frame.wait_for_selector("text=Longer run", timeout=5000)
                break
            except Exception:
                try:
                    frame.get_by_role("link", name=re.compile("Chart", re.I)).click(timeout=2000)
                except Exception:
                    pass
                try:
                    frame.get_by_role("link", name=re.compile("Table", re.I)).click(timeout=2000)
                except Exception:
                    pass
    except Exception:
        pass

    # If the text shows up, grab the last table on the page (Dot Plot table renders at the end)
    try:
        frame.wait_for_selector("text=Longer run", timeout=15000)
        tbl = frame.locator("table").last
        return tbl.evaluate(EVAL_PARSE_TABLE)
    except Exception:
        pass

    # Fallback: scan all tables and pick the one containing Target Rate & Longer run
    return frame.evaluate("""
    () => {
      const txt = n => (n.innerText || n.textContent || '').trim().replace(/\s+/g, ' ');
      const tables = Array.from(document.querySelectorAll('table'));
      for (const t of tables) {
        const all = (t.textContent || '').toLowerCase();
        if (all.includes('target rate') && all.includes('longer run')) {
          const thead = Array.from(t.querySelectorAll('thead th')).map(txt);
          const rows = Array.from(t.querySelectorAll('tbody tr')).map(r => Array.from(r.children).map(txt));
          const header = thead.length ? thead : (rows.length ? rows[0] : []);
          const body = thead.length ? rows : rows.slice(1);
          const data = body.map(r => { const o = {}; r.forEach((v,i)=> o[header[i] || `col_${i}`] = v ); return o; });
          return { header, rows: body, data };
        }
      }
      return { header: [], rows: [], data: [] };
    }
    """)


def main(out_dir="outputs"):
    os.makedirs(out_dir, exist_ok=True)
    with sync_playwright() as p:
        browser = None
        ctx = None
        page = None
        try:
            browser, ctx, page = open_with_fallback(p)
            accept_cookies(page)
            fw = get_fedwatch_frame(page)

            probs = parse_probabilities(fw)
            dot = parse_dotplot_table(fw)
        finally:
            # We'll close at the end after writing files
            pass

        prob_path = os.path.join(out_dir, "fedwatch_probabilities.json")
        dot_path = os.path.join(out_dir, "fedwatch_dot_plot_table.json")
        with open(prob_path, "w", encoding="utf-8") as f:
            json.dump(probs, f, ensure_ascii=False, indent=2)
        with open(dot_path, "w", encoding="utf-8") as f:
            json.dump(dot, f, ensure_ascii=False, indent=2)

        print("Saved:", prob_path, dot_path)
        print("Probabilities rows:", len(probs.get("data", [])))
        print("Dot plot rows:", len(dot.get("data", [])))
        # Cleanup browser resources
        try:
            if ctx:
                ctx.close()
            if browser:
                browser.close()
        except Exception:
            pass


if __name__ == "__main__":
    for arg in sys.argv[1:]:
        if arg.startswith("--proxy="):
            os.environ["PLAYWRIGHT_PROXY"] = arg.split("=", 1)[1]
    main()

