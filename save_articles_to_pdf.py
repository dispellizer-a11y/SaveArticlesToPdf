#!/usr/bin/env python3
"""
엑셀에 정리된 기사 링크를 열어, 각 기사 페이지의 '인쇄' 버튼을 클릭해
인쇄용 레이아웃으로 전환한 뒤 그 결과를 PDF로 저장하는 스크립트.

사용 예:
    python save_articles_to_pdf.py --excel 보도자료.xlsx --limit 5
    python save_articles_to_pdf.py --excel 보도자료.xlsx --sheets 2025,2026
    python save_articles_to_pdf.py --excel 보도자료.xlsx
"""

import argparse
import csv
import re
import sys
import time
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

# 사이트마다 마크업이 달라서, 자주 쓰이는 "인쇄" 버튼 패턴을 최대한 폭넓게 시도한다.
PRINT_SELECTORS = [
    'a[onclick*="print" i]',
    'button[onclick*="print" i]',
    'a[href*="javascript:print" i]',
    'a[href*="print" i]',
    'area[href*="print" i]',
    '[class*="print" i] a',
    '[class*="btn_print" i]',
    '[class*="print_btn" i]',
    'a[class*="print" i]',
    'button[class*="print" i]',
    'span[class*="print" i]',
    'li[class*="print" i] a',
    'li[class*="print" i]',
    'div[class*="print" i]',
    'i[class*="print" i]',
    'a[id*="print" i]',
    'button[id*="print" i]',
    'img[alt*="인쇄"]',
    'img[alt*="프린트"]',
    'img[alt*="print" i]',
    '[aria-label*="인쇄"]',
    '[aria-label*="print" i]',
    '[title*="인쇄"]',
    '[title*="print" i]',
]

PRINT_TEXT_RE = re.compile(r"^\s*(인쇄|프린트|print)\s*$", re.IGNORECASE)

INVALID_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\n\r\t]')


def sanitize_filename(text: str, max_len: int = 80) -> str:
    text = INVALID_FILENAME_CHARS.sub("", text).strip()
    text = re.sub(r"\s+", " ", text)
    if len(text) > max_len:
        text = text[:max_len].rstrip()
    return text or "제목없음"


def read_articles(excel_path: Path, sheet_names):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheets = sheet_names or wb.sheetnames
    articles = []
    for sheet in sheets:
        if sheet not in wb.sheetnames:
            print(f"[경고] 시트 '{sheet}' 를 찾을 수 없어 건너뜁니다.", file=sys.stderr)
            continue
        ws = wb[sheet]
        for row in ws.iter_rows(min_row=1, values_only=True):
            # 열 구성: B=NO, C=게시일, D=언론사, E=제목, F=주소, G=비고
            if len(row) < 6:
                continue
            no, date, press, title, url = row[1], row[2], row[3], row[4], row[5]
            if not url or not isinstance(url, str) or not url.strip().startswith("http"):
                continue
            articles.append(
                {
                    "sheet": sheet,
                    "no": no,
                    "date": str(date) if date is not None else "",
                    "press": str(press) if press is not None else "",
                    "title": str(title) if title is not None else "",
                    "url": url.strip(),
                }
            )
    return articles


def find_print_element(page):
    """일치한 요소와 함께, 진단용으로 '어떤 규칙으로 찾았는지'도 같이 반환한다."""
    for selector in PRINT_SELECTORS:
        try:
            locator = page.locator(selector).first
            if locator.count() > 0 and locator.is_visible():
                return locator, f"selector:{selector}"
        except Exception:
            continue

    # 셀렉터로 못 찾으면 화면에 보이는 텍스트가 "인쇄"/"프린트"/"print" 인
    # 클릭 가능한 요소를 찾는다.
    try:
        candidates = page.locator("a, button, span, li, div").filter(has_text=PRINT_TEXT_RE)
        count = min(candidates.count(), 20)
        for i in range(count):
            el = candidates.nth(i)
            try:
                if el.is_visible():
                    return el, "text-match"
            except Exception:
                continue
    except Exception:
        pass
    return None, None


def describe_element(locator) -> str:
    try:
        html = locator.evaluate("el => el.outerHTML")
        return html[:300] if html else ""
    except Exception:
        return ""


def render_to_pdf(page, out_path: Path):
    page.emulate_media(media="print")
    page.wait_for_timeout(500)
    page.pdf(path=str(out_path), format="A4", print_background=True,
             margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"})


POPUP_WAIT_MS = 8000


def process_article(context, page, article, out_path: Path, nav_timeout: int,
                     console_log=None, debug_dir: Path = None):
    """
    1) 기사 페이지 접속
    2) '인쇄' 버튼을 찾아 클릭 (새 탭으로 인쇄용 페이지가 열리는 사이트도 대응)
    3) 인쇄(print) 레이아웃으로 전환된 화면을 PDF로 저장
    반환값: dict(status, method, matched, before_url, final_url, notes)

    주의: window.print() 무력화 스크립트는 이 컨텍스트에서 새로 열리는 모든
    페이지(팝업 포함)에 적용되어야 한다. page 단위로만 걸면 인쇄 버튼이
    새 팝업창을 여는 사이트에서 그 팝업에는 적용되지 않아, 팝업이 뜨자마자
    실제 OS 인쇄창이 열려 자동화가 멈춰버린다. main()에서 context 생성 직후
    context.add_init_script(...)로 걸어야 한다.
    """
    if console_log is not None:
        console_log.clear()

    try:
        page.goto(article["url"], wait_until="load", timeout=nav_timeout)
    except PWTimeout:
        page.goto(article["url"], wait_until="domcontentloaded", timeout=nav_timeout)

    # 'load' 이벤트가 떠도, 광고/추적 스크립트가 그 뒤에 비동기로 인쇄
    # 버튼의 클릭 핸들러를 붙이는 사이트가 있다 (마이데일리 등). 그 핸들러가
    # 붙을 시간을 조금 준다.
    page.wait_for_timeout(1500)

    before_url = page.url
    print_el, matched = find_print_element(page)

    if print_el is None:
        # 인쇄 버튼을 못 찾은 경우: 원문 페이지에 인쇄 스타일만 적용해서 저장 (대체 경로)
        render_to_pdf(page, out_path)
        return {"status": "fallback_no_button", "method": "direct_print_css",
                "matched": "", "before_url": before_url, "final_url": page.url, "notes": ""}

    matched_html = describe_element(print_el)

    target_page = page
    new_page = None
    try:
        # 클릭은 여기서 딱 한 번만 한다. 같은 탭에서 바로 페이지가 바뀌는
        # 사이트의 경우, 클릭이 이미 성공했는데도 여기서 또 클릭하면
        # (이미 사라진) 이전 문서 기준 버튼을 찾다 실패해서 "인쇄 버튼을
        # 거치지 않은 것"처럼 오작동했었다.
        with context.expect_page(timeout=POPUP_WAIT_MS) as new_page_info:
            print_el.click(timeout=5000)
        new_page = new_page_info.value
        new_page.wait_for_load_state("load", timeout=nav_timeout)
        target_page = new_page
        method = "print_button_popup"
    except PWTimeout:
        # 새 탭이 안 열림 = 같은 페이지에서 처리된 경우 (클릭은 이미 위에서 실행됨)
        try:
            page.wait_for_load_state("load", timeout=3000)
        except PWTimeout:
            pass
        page.wait_for_timeout(500)
        method = "print_button_same_tab"

    if debug_dir is not None:
        no_str = str(article.get("no", "x"))
        try:
            target_page.screenshot(path=str(debug_dir / f"{no_str}_after_click.png"))
        except Exception:
            pass

    render_to_pdf(target_page, out_path)
    final_url = target_page.url

    if new_page is not None:
        new_page.close()

    notes = " | ".join(console_log) if console_log else ""

    return {"status": "success", "method": method, "matched": f"{matched} | {matched_html}",
            "before_url": before_url, "final_url": final_url, "notes": notes}


def main():
    parser = argparse.ArgumentParser(description="엑셀의 기사 링크를 인쇄 버튼 경유로 PDF 저장")
    parser.add_argument("--excel", required=True, help="엑셀 파일 경로")
    parser.add_argument("--output", default="output", help="PDF 저장 폴더 (기본: output)")
    parser.add_argument("--sheets", default="", help="처리할 시트명 콤마 구분 (기본: 전체 시트)")
    parser.add_argument("--limit", type=int, default=0, help="테스트용 최대 처리 건수 (0=전체)")
    parser.add_argument("--skip", type=int, default=0, help="앞에서부터 건너뛸 건수 (테스트용)")
    parser.add_argument("--delay", type=float, default=1.5, help="요청 간 대기 시간(초)")
    parser.add_argument("--timeout", type=int, default=30000, help="페이지 로딩 타임아웃(ms)")
    parser.add_argument("--headed", action="store_true", help="브라우저 창을 보이게 실행 (디버깅용)")
    parser.add_argument("--overwrite", action="store_true", help="이미 저장된 PDF도 다시 저장")
    parser.add_argument("--debug", action="store_true",
                         help="클릭 직후 스크린샷과 콘솔 경고/에러 메시지를 결과 CSV/파일로 남김")
    args = parser.parse_args()

    excel_path = Path(args.excel)
    if not excel_path.exists():
        print(f"엑셀 파일을 찾을 수 없습니다: {excel_path}", file=sys.stderr)
        sys.exit(1)

    sheet_names = [s.strip() for s in args.sheets.split(",") if s.strip()] or None
    articles = read_articles(excel_path, sheet_names)
    if args.skip:
        articles = articles[args.skip:]
    if args.limit:
        articles = articles[: args.limit]

    print(f"총 {len(articles)}건의 기사를 처리합니다.")

    out_root = Path(args.output)
    out_root.mkdir(parents=True, exist_ok=True)
    log_path = out_root / "results.csv"
    write_header = not log_path.exists()

    with sync_playwright() as p, open(log_path, "a", newline="", encoding="utf-8-sig") as log_file:
        writer = csv.writer(log_file)
        if write_header:
            writer.writerow(["sheet", "no", "date", "press", "title", "url",
                              "status", "method", "output_file", "before_url", "final_url",
                              "matched_element", "notes", "error"])

        debug_dir = None
        if args.debug:
            debug_dir = out_root / "_debug"
            debug_dir.mkdir(parents=True, exist_ok=True)

        # 자동화로 띄운 브라우저는 매번 새 프로필로 시작해서, 실제 사용해온
        # 크롬과 달리 사이트별 팝업 허용 이력/신뢰도가 없다. 그래서 인쇄
        # 버튼이 여는 새 창(window.open)이 조용히 차단되는 사이트가 있어
        # 팝업 차단 자체를 꺼서 실행한다.
        browser = p.chromium.launch(headless=not args.headed, args=["--disable-popup-blocking"])
        context = browser.new_context(user_agent=USER_AGENT, locale="ko-KR")

        # 광고/추적 스크립트가 자동화(봇)를 감지해서 window.open 같은 기능을
        # 스스로 무력화하는 경우가 있다. navigator.webdriver는 Playwright가
        # 기본으로 노출하는 가장 흔한 자동화 신호라 이것부터 감춘다.
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', { get: () => undefined });"
        )

        # 컨텍스트 단위로 걸어야 인쇄 버튼이 새로 여는 팝업창에도 적용된다.
        # (page 단위로 걸면 팝업에는 적용 안 되어 실제 인쇄창이 뜬다)
        # window.print()를 완전히 빈 함수로 막으면, 실제 인쇄창은 안 뜨지만
        # 그 안에서 발생해야 할 beforeprint 이벤트도 같이 사라진다. 많은
        # 사이트가 광고/메뉴를 숨기는 '인쇄 모드 전환' 로직을 이 이벤트에
        # 걸어두기 때문에, 이벤트만 대신 발생시켜서 그 로직은 그대로 타게 한다.
        context.add_init_script(
            "window.print = function(){ "
            "window.__printTriggered = true; "
            "try { window.dispatchEvent(new Event('beforeprint')); } catch(e) {} "
            "};"
        )

        # 콘솔 경고/에러(예: "팝업이 차단되었습니다")를 잡아서 진단에 쓴다.
        # 새로 열리는 페이지(팝업)에도 자동으로 걸리도록 컨텍스트 단위로 등록한다.
        console_log = []

        def _capture_console(msg):
            if msg.type in ("warning", "error"):
                console_log.append(f"[{msg.type}] {msg.text}")

        context.on("page", lambda p_: p_.on("console", _capture_console))
        page = context.new_page()
        page.on("console", _capture_console)

        for i, article in enumerate(articles, start=1):
            sheet_dir = out_root / article["sheet"]
            sheet_dir.mkdir(parents=True, exist_ok=True)
            no_str = f"{int(article['no']):03d}" if str(article["no"]).isdigit() else str(article["no"])
            filename = sanitize_filename(
                f"{no_str}_{article['date']}_{article['press']}_{article['title']}"
            ) + ".pdf"
            out_path = sheet_dir / filename

            if out_path.exists() and not args.overwrite:
                print(f"[{i}/{len(articles)}] 이미 존재함, 건너뜀: {filename}")
                continue

            print(f"[{i}/{len(articles)}] {article['press']} - {article['title'][:40]} ...")
            try:
                result = process_article(context, page, article, out_path, args.timeout,
                                          console_log=console_log, debug_dir=debug_dir)
                writer.writerow([article["sheet"], article["no"], article["date"], article["press"],
                                  article["title"], article["url"], result["status"], result["method"],
                                  str(out_path), result["before_url"], result["final_url"],
                                  result["matched"], result["notes"], ""])
                same_url = result["before_url"] == result["final_url"]
                print(f"    -> {result['status']} ({result['method']}) "
                      f"url_changed={not same_url}")
                if result["notes"]:
                    print(f"    콘솔 경고/에러: {result['notes']}")
            except Exception as e:
                writer.writerow([article["sheet"], article["no"], article["date"], article["press"],
                                  article["title"], article["url"], "failed", "", "", "", "", "", "", str(e)])
                print(f"    -> 실패: {e}")
            log_file.flush()

            time.sleep(args.delay)

        context.close()
        browser.close()

    print(f"완료. 로그: {log_path}")


if __name__ == "__main__":
    main()
