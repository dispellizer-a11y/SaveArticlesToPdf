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
    for selector in PRINT_SELECTORS:
        try:
            locator = page.locator(selector).first
            if locator.count() > 0 and locator.is_visible():
                return locator
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
                    return el
            except Exception:
                continue
    except Exception:
        pass
    return None


def render_to_pdf(page, out_path: Path):
    page.emulate_media(media="print")
    page.wait_for_timeout(500)
    page.pdf(path=str(out_path), format="A4", print_background=True,
             margin={"top": "10mm", "bottom": "10mm", "left": "10mm", "right": "10mm"})


def process_article(context, page, article, out_path: Path, nav_timeout: int):
    """
    1) 기사 페이지 접속
    2) '인쇄' 버튼을 찾아 클릭 (새 탭으로 인쇄용 페이지가 열리는 사이트도 대응)
    3) 인쇄(print) 레이아웃으로 전환된 화면을 PDF로 저장
    반환값: (status, method)
    """
    page.add_init_script("window.print = function(){ window.__printTriggered = true; };")
    try:
        page.goto(article["url"], wait_until="load", timeout=nav_timeout)
    except PWTimeout:
        page.goto(article["url"], wait_until="domcontentloaded", timeout=nav_timeout)

    print_el = find_print_element(page)

    if print_el is None:
        # 인쇄 버튼을 못 찾은 경우: 원문 페이지에 인쇄 스타일만 적용해서 저장 (대체 경로)
        render_to_pdf(page, out_path)
        return "fallback_no_button", "direct_print_css"

    target_page = page
    new_page = None
    try:
        with context.expect_page(timeout=4000) as new_page_info:
            print_el.click(timeout=5000)
        new_page = new_page_info.value
        new_page.wait_for_load_state("load", timeout=nav_timeout)
        target_page = new_page
    except PWTimeout:
        # 새 탭이 안 열림 = 같은 페이지에서 처리되는 경우
        try:
            print_el.click(timeout=5000)
        except Exception:
            print_el.evaluate("el => el.click()")
        page.wait_for_timeout(800)

    render_to_pdf(target_page, out_path)

    if new_page is not None:
        new_page.close()

    return "success", "print_button"


def main():
    parser = argparse.ArgumentParser(description="엑셀의 기사 링크를 인쇄 버튼 경유로 PDF 저장")
    parser.add_argument("--excel", required=True, help="엑셀 파일 경로")
    parser.add_argument("--output", default="output", help="PDF 저장 폴더 (기본: output)")
    parser.add_argument("--sheets", default="", help="처리할 시트명 콤마 구분 (기본: 전체 시트)")
    parser.add_argument("--limit", type=int, default=0, help="테스트용 최대 처리 건수 (0=전체)")
    parser.add_argument("--delay", type=float, default=1.5, help="요청 간 대기 시간(초)")
    parser.add_argument("--timeout", type=int, default=30000, help="페이지 로딩 타임아웃(ms)")
    parser.add_argument("--headed", action="store_true", help="브라우저 창을 보이게 실행 (디버깅용)")
    parser.add_argument("--overwrite", action="store_true", help="이미 저장된 PDF도 다시 저장")
    args = parser.parse_args()

    excel_path = Path(args.excel)
    if not excel_path.exists():
        print(f"엑셀 파일을 찾을 수 없습니다: {excel_path}", file=sys.stderr)
        sys.exit(1)

    sheet_names = [s.strip() for s in args.sheets.split(",") if s.strip()] or None
    articles = read_articles(excel_path, sheet_names)
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
                              "status", "method", "output_file", "error"])

        browser = p.chromium.launch(headless=not args.headed)
        context = browser.new_context(user_agent=USER_AGENT, locale="ko-KR")
        page = context.new_page()

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
                status, method = process_article(context, page, article, out_path, args.timeout)
                writer.writerow([article["sheet"], article["no"], article["date"], article["press"],
                                  article["title"], article["url"], status, method, str(out_path), ""])
                print(f"    -> {status} ({method})")
            except Exception as e:
                writer.writerow([article["sheet"], article["no"], article["date"], article["press"],
                                  article["title"], article["url"], "failed", "", "", str(e)])
                print(f"    -> 실패: {e}")
            log_file.flush()

            time.sleep(args.delay)

        context.close()
        browser.close()

    print(f"완료. 로그: {log_path}")


if __name__ == "__main__":
    main()
