#!/usr/bin/env python3
"""미간프레스 사이트 생성기.

구글 시트 '미간프레스 웹사이트 데이터'를 읽어 dist/ 에 정적 HTML을 만듭니다.
파이썬 표준 라이브러리만 사용하므로 따로 설치할 것이 없습니다.

    python3 build.py                     # 구글 시트에서 읽기
    python3 build.py --local data/x.xlsx # 내려받은 xlsx 파일로 시험하기
"""
import base64
import hashlib
import json
import mimetypes
import os
import re
import shutil
import sys
import urllib.request
from html import escape
from pathlib import Path

from sheet import fetch_xlsx, read_workbook, records, serial_to_month

ROOT = Path(__file__).parent
STATIC = ROOT / "static"
DIST = ROOT / "dist"
SITE_URL = "https://miganpress.com"
SHEET_ID = os.environ.get("SHEET_ID", "1OJ-63umjkmWIJITMQwiwXcVU7wcFpFSMmSTBzhElDOg")

# 시트 '사이트 정보' 탭의 항목 이름 → 코드에서 쓰는 이름
SITE_KEYS = {
    "첫 화면 문장": "tagline",        # 검색·공유용 사이트 설명
    "소개 큰 문장": "about_heading",
    "소개 문단": "about_text",
    "영문 소개 문단": "en_text",
    "대표 이름": "person_name",
    "대표 소개": "person_bio",
    "대표 링크": "person_links",     # 한 줄에 하나: 이름 | 주소
    "이메일": "email",
    "찾아오시는 길": "visit",
    "인스타그램": "instagram",
    "블루스카이": "bluesky",
    "트위터": "twitter",
    "X": "twitter",
}

# '출간 도서' 탭 머리글 → 코드에서 쓰는 이름
BOOK_KEYS = {
    "공개": "public", "번호": "no", "주소": "slug", "제목": "title", "부제": "subtitle",
    "지은이": "author", "옮긴이": "translator", "원서 제목": "original_title", "원제": "original_title",
    "원서 부제": "original_subtitle", "원서 저자": "original_author", "저자명 원문": "original_author", "출간 연월": "pub_date",
    "상태": "status", "표지": "cover", "한 줄 소개": "tagline", "태그라인": "tagline", "책 소개": "description",
    "짧은 소개": "short", "긴 소개": "description",
    "저자 소개": "author_bio", "지은이 소개": "author_bio",
    "역자 소개": "translator_bio", "옮긴이 소개": "translator_bio",
    "추천사": "praise", "추천의 말": "praise",
    "수상 내역": "awards", "수상/선정 내역": "awards", "수상·선정 내역": "awards", "수상": "awards", "수상 경력": "awards",
    "사양": "binding", "제본": "binding", "판형": "size", "판형과 사양": "size", "차례": "toc", "목차": "toc", "부가기호": "add_code", "정가": "price", "가격": "price", "면수": "pages", "쪽수": "pages", "ISBN": "isbn", "보도자료": "press_kit", "표지 파일": "cover_file", "고해상도 표지": "cover_file", "교보문고": "kyobo", "예스24": "yes24", "알라딘": "aladin",
}

# '편집한 책' 탭 머리글 → 코드에서 쓰는 이름
WORK_KEYS = {
    "공개": "public", "순서": "order", "제목": "title", "지은이": "author", "옮긴이": "translator",
    "출판사": "publisher", "출간 연도": "year", "표지": "cover", "메모": "note", "ISBN": "isbn", "분야": "category", "분류": "category",
}


def e(value):
    return escape(str(value or ""), quote=True)


def is_public(row):
    return str(row.get("public", "Y")).strip().upper() not in ("N", "NO", "아니오", "비공개")


def rename(row, keys):
    return {keys.get(k, k): v for k, v in row.items()}


# 제목 줄바꿈: 시트 제목에 "/"를 넣으면 그 자리에서 줄이 바뀝니다(예: 우리가 / 함께 읽는다면).
# 시트를 고치기 전까지 쓰는 임시 목록입니다. 시트에 "/"를 넣었다면 여기서 지워도 됩니다.
TITLE_BREAKS = {"우리가 함께 읽는다면": "우리가 / 함께 읽는다면"}


def load(local_path=None):
    data = Path(local_path).read_bytes() if local_path else fetch_xlsx(SHEET_ID)
    tabs = read_workbook(data)
    # '편집한 책' 탭은 '만든 책'으로 이름을 바꿔도 됩니다
    if "만든 책" in tabs and "편집한 책" not in tabs:
        tabs["편집한 책"] = tabs["만든 책"]
    for need in ("출간 도서", "편집한 책", "사이트 정보"):
        if need not in tabs:
            raise SystemExit(f"시트에 '{need}' 탭이 없습니다. 탭 이름을 확인하세요.")

    site = {}
    for row in tabs["사이트 정보"][1:]:
        if len(row) >= 2 and row[0].strip() in SITE_KEYS:
            site[SITE_KEYS[row[0].strip()]] = row[1].strip()

    books = []
    for row in records(tabs["출간 도서"]):
        b = rename(row, BOOK_KEYS)
        if not b.get("title") or not is_public(b):
            continue
        t = b["title"].strip()
        t = TITLE_BREAKS.get(t, t)
        b["title"] = re.sub(r"\s*/\s*", " ", t)
        b["title_html"] = "<br>".join(e(x.strip()) for x in t.split("/"))
        no = b.get("no", "")
        b["no"] = no.zfill(3) if no.isdigit() else no
        b["pub_date"] = serial_to_month(b.get("pub_date", ""))
        b["status"] = "published" if b.get("status", "").strip() == "출간" else "upcoming"
        slug = re.sub(r"[^a-z0-9-]+", "-", (b.get("slug") or "").lower()).strip("-")
        b["slug"] = slug or b["no"] or hashlib.md5(b["title"].encode()).hexdigest()[:8]
        books.append(b)
    books.sort(key=lambda b: str(b.get("no", "")))

    works = []
    for row in records(tabs["편집한 책"]):
        w = rename(row, WORK_KEYS)
        if w.get("title") and is_public(w):
            works.append(w)
    works.sort(key=lambda w: float(w["order"]) if re.fullmatch(r"\d+(\.\d+)?", w.get("order", "")) else 9999)
    return site, books, works


def drive_download_url(url):
    """구글 드라이브 공유 링크를 바로 내려받을 수 있는 주소로 바꿉니다."""
    m = re.search(r"/d/([A-Za-z0-9_-]{10,})", url) or re.search(r"[?&]id=([A-Za-z0-9_-]{10,})", url)
    if "drive.google.com" in url and m:
        return f"https://drive.google.com/uc?export=download&id={m.group(1)}"
    return url


def download_image(url, name):
    req = urllib.request.Request(url, headers={"User-Agent": "migan-site-build"})
    with urllib.request.urlopen(req, timeout=60) as res:
        ctype = res.headers.get_content_type()
        data = res.read()
    if not ctype.startswith("image/"):
        raise RuntimeError(f"이미지가 아닙니다 ({ctype}). 공유 설정을 확인하세요.")
    ext = mimetypes.guess_extension(ctype) or ".jpg"
    ext = ".jpg" if ext in (".jpe", ".jpeg") else ext
    out = DIST / "covers" / f"{name}{ext}"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(data)
    return f"/covers/{out.name}"


def partial(name):
    return (ROOT / "partials" / name).read_text(encoding="utf-8").strip()


def data_json(name):
    path = ROOT / "data" / name
    return json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}


LOGO_STATIC = partial("logo_static_svg.html")    # 머리 로고(모션 로고의 마지막 모양)
MOTION_JS = partial("logo_motion_js.html")
LOGO_HEADER = partial("logo_header_svg.html")    # 머리 로고(국문 세로형, 미간 빨강)
BLINK_JS = partial("logo_blink_js.html")
# 스타일 파일이 바뀌면 주소 끝 ?v= 값도 바뀌어, 브라우저가 예전 파일을 쓰지 않습니다
CSS_VER = hashlib.sha1((STATIC / "style.css").read_bytes()).hexdigest()[:8]         # 머리 로고 눈 깜빡임
EMAIL_JS = partial("email_js.html")              # 이메일 주소 조립
BOOK_NAV_JS = partial("book_nav_js.html")        # 책 상세 왼쪽 목차: 지금 읽는 구역 표시
WORK_CATEGORIES = data_json("work_categories.json")   # 만든 책 제목 → 분야
WORK_COVERS = data_json("work_covers.json")           # 만든 책 제목 → 표지 파일


def email_link(addr):
    """이메일 수집 방지: 주소를 base64로 숨기고, 페이지가 열릴 때 JS가 조립합니다.
    JS가 꺼져 있으면 'name [at] domain' 글자만 보입니다."""
    addr = (addr or "").strip()
    enc = base64.b64encode(addr.encode()).decode()
    return f'<a data-e="{enc}">{e(addr.replace("@", " [at] "))}</a>'


def resolve_cover(item, name):
    """표지 칸 → 이미지 파일을 정합니다. 주소(구글 드라이브 포함)면 내려받고, 파일 이름이면 covers 폴더에서 씁니다. 비어 있으면 글자 표지가 나옵니다."""
    value = (item.get("cover") or "").strip()
    if not value:
        return ""
    if not value.startswith("http"):
        return f"/covers/{value.lstrip('/').removeprefix('covers/')}"
    try:
        return download_image(drive_download_url(value), name)
    except Exception as err:  # noqa: BLE001
        print(f"  ! 표지를 가져오지 못했습니다 ({name}): {err}", file=sys.stderr)
        return ""


def fmt_date(value):
    """'2026-11' -> '2026년 11월', '2027' -> '2027년'"""
    v = str(value or "").strip()
    m = re.fullmatch(r"(\d{4})-(\d{1,2})", v)
    if m:
        return f"{m.group(1)}년 {int(m.group(2))}월"
    if re.fullmatch(r"\d{4}", v):
        return f"{v}년"
    return v


def status_label(book):
    date = fmt_date(book.get("pub_date"))
    return f"{date} 출간" if book.get("status") == "published" else f"{date} 출간 예정"


# 사람 이름: 시트에는 여러 명을 쉼표(또는 ·)로 나눠 적습니다.
# 사이트에서는 가운뎃점(·)으로 잇고, 좁은 자리(표지·목록·홈)에서는 세 명 이상이면 '첫 사람 외'로 줄입니다.
SHORTEN_FROM = 3


def names(value):
    return [n.strip() for n in re.split(r"\s*[,，·、/]\s*", str(value or "")) if n.strip()]


def names_ko(value, short=False):
    ns = names(value)
    if short and len(ns) >= SHORTEN_FROM:
        return f"{ns[0]} 외"
    return "·".join(ns)


def names_en(value, short=False):
    """영문 이름: A / A and B / A, B, and C. 줄이면 A et al."""
    ns = names(value)
    if short and len(ns) >= SHORTEN_FROM:
        return f"{ns[0]} et al."
    if len(ns) <= 2:
        return " and ".join(ns)
    return ", ".join(ns[:-1]) + ", and " + ns[-1]


def credit_parts(book, short=False):
    parts = []
    if book.get("author"):
        parts.append(f"{names_ko(book['author'], short)} 지음")
    if book.get("translator"):
        parts.append(f"{names_ko(book['translator'], short)} 옮김")
    return parts


def credit(book, short=False):
    return ", ".join(credit_parts(book, short))


def credit_html(book):
    """목록용(줄인 이름): 데스크톱은 '지음, 옮김' 한 줄, 모바일은 지은이/옮긴이를 줄로 나눕니다."""
    return '<span class="credit-sep">, </span>'.join(f"<span>{e(x)}</span>" for x in credit_parts(book, short=True))


def praise_html(text):
    """추천사 칸: 추천사 사이는 빈 줄로 나누고, 각 추천사의 마지막 줄을 '—'나 '-'로 시작하면 출처로 씁니다.
    예)
    읽는 내내 마음이 저렸다.
    — 김OO, 소설가
    """
    items = []
    for block in [b.strip() for b in str(text or "").split("\n\n") if b.strip()]:
        lines = [l.strip() for l in block.split("\n") if l.strip()]
        source = ""
        if len(lines) > 1 and re.match(r"^[—–-]", lines[-1]):
            source = re.sub(r"^[—–-]+\s*", "", lines.pop())
        else:
            # 같은 줄 끝에 ' — 출처'로 적은 경우
            m = re.match(r"^(.*\S)\s+[—–]\s*(\S.{0,80})$", lines[-1])
            if m:
                lines[-1], source = m.group(1), m.group(2)
        # 시트에 따옴표를 이미 적었으면 지웁니다(사이트가 따옴표를 붙여 줍니다)
        if lines:
            lines[0] = re.sub(r'^["“”\'‘’]+\s*', "", lines[0])
            lines[-1] = re.sub(r'\s*["“”\'‘’]+$', "", lines[-1])
        quote = "<br>".join(rich(l) for l in lines)
        cite = f"<figcaption>{e(source)}</figcaption>" if source else ""
        items.append(f"<figure><blockquote><p>{quote}</p></blockquote>{cite}</figure>")
    return "\n".join(items)


def intro_html(book):
    """짧은 소개와 긴 소개를 함께 씁니다.
    둘 다 있으면 짧은 소개를 먼저 보여 주고, 긴 소개는 '책 소개 전체 읽기'를 눌러 펼쳐요.
    하나만 있으면 그것을 그대로 보여 줍니다."""
    short, long_ = book.get("short"), book.get("description")
    if short and long_:
        return (
            f"{paragraphs(short)}\n"
            f'<details class="more-intro"><summary>책 소개 전체 읽기</summary>'
            f'<div class="prose">{paragraphs(long_)}</div></details>'
        )
    return paragraphs(short or long_)


def toc_html(text):
    """차례: 한 줄에 한 항목. 빈 줄로 부를 나누면 사이가 조금 벌어집니다. 길어서 접어 둡니다."""
    text = (text or "").strip()
    if not text:
        return ""
    parts = [p for p in re.split(r"\n\s*\n", text) if p.strip()]
    lists = "".join(
        "<ul>" + "".join(f"<li>{rich(line.strip())}</li>" for line in part.splitlines() if line.strip()) + "</ul>"
        for part in parts
    )
    return f'<details class="more-intro toc"><summary>차례 보기</summary><div class="toc__body">{lists}</div></details>'


def awards_html(text):
    """수상 내역: 한 줄에 하나씩. 예) 2025 퓰리처상 회고록·자서전 부문 최종 후보"""
    lines = [l.strip().lstrip("-·•").strip() for l in str(text or "").splitlines() if l.strip()]
    if not lines:
        return ""
    return '<ul class="awards">' + "".join(f"<li>{rich(l)}</li>" for l in lines) + "</ul>"


def book_section(title, inner, cls="", sid=""):
    if not inner.strip():
        return ""
    id_attr = f' id="{sid}"' if sid else ""
    return f'<section class="book-section {cls}"{id_attr}><h2 class="caps-ko">{title}</h2>{inner}</section>'


def rich(text):
    """글자를 안전하게 바꾸고, **굵게**·__기울임__ 표시한 부분만 살립니다. 줄바꿈은 그대로 살립니다."""
    html = e(text)
    html = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", html, flags=re.S)
    html = re.sub(r"__(.+?)__", r"<em>\1</em>", html, flags=re.S)
    return html.replace("\n", "<br>")


def plain(text):
    """굵게 표시(**)를 뺀 글자. 검색 설명 등에 씁니다."""
    return re.sub(r"__(.+?)__", r"\1", re.sub(r"\*\*(.+?)\*\*", r"\1", str(text or ""), flags=re.S), flags=re.S)


def paragraphs(text):
    blocks = [p.strip() for p in str(text or "").split("\n\n") if p.strip()]
    return "\n".join(f"<p>{rich(p)}</p>" for p in blocks)


def cover(book, size="m", link=True, dark=False):
    """실제 표지가 있으면 이미지를, 없으면 흰 바탕의 글자 표지를 그립니다.
    글자 표지 안에는 제목·지은이·옮긴이와, 오른쪽 위에 출간 연도가 들어갑니다."""
    year = str(book.get("pub_date") or "")[:4]
    if book.get("cover"):
        body = f'<img src="{e(book["cover"])}" alt="" loading="lazy">'
        cls = f"cover cover--{size} cover--image"
    else:
        credit_lines = "".join(f"<span>{e(x)}</span>" for x in credit_parts(book, short=True))
        year_html = f'<span class="cover__year">{e(year)}</span>' if year else ""
        body = (
            f'{year_html}'
            f'<span class="cover__title">{book["title_html"]}</span>'
            f'<span class="cover__credit">{credit_lines}</span>'
        )
        cls = f"cover cover--{size}" + (" cover--dark" if dark else " cover--next" if book.get("is_next") else "")
    label = f"『{book['title']}』 {credit(book)}, {status_label(book)}"
    if link:
        return f'<a class="{cls}" href="/books/{e(book["slug"])}/" aria-label="{e(label)}">{body}</a>'
    return f'<div class="{cls}" role="img" aria-label="{e(label)}">{body}</div>'


# SNS: 시트 칸 이름(코드 이름) → 링크 모양. 이름은 페이지마다 다르게 씁니다.
SOCIALS = (
    ("instagram", "https://www.instagram.com/{}/"),
    ("twitter", "https://x.com/{}"),
    ("bluesky", "https://bsky.app/profile/{}"),
)
SOCIAL_LABELS = {
    "ko": {"instagram": "인스타그램", "twitter": "트위터", "bluesky": "블루스카이"},
    "en": {"instagram": "Instagram", "twitter": "Twitter", "bluesky": "Bluesky"},
}


def social_url(value, pattern):
    """'@migan.press' 같은 계정 이름도, 전체 링크도 받아 줍니다."""
    value = (value or "").strip()
    if not value or value.startswith("http"):
        return value
    return pattern.format(value.lstrip("@"))


def social_handle(value):
    value = (value or "").strip()
    if value.startswith("http"):
        return value.rstrip("/").rsplit("/", 1)[-1]
    return value if value.startswith("@") else f"@{value}"


def social_rows(site, lang="ko"):
    """연락처 표에 넣을 SNS 줄: (이름, 링크)."""
    return [
        (SOCIAL_LABELS[lang][key], f'<a href="{e(social_url(site[key], pattern))}">{e(social_handle(site[key]))}</a>')
        for key, pattern in SOCIALS
        if site.get(key)
    ]


def contact_table(rows):
    return '<dl class="meta about-contact">' + "".join(f"<div><dt>{k}</dt><dd>{v}</dd></div>" for k, v in rows) + "</dl>"


def social_icons(site, lang="ko"):
    """꼬리의 SNS 아이콘 링크. static/icons/{instagram,twitter,bluesky}.svg 파일이 있으면 그 아이콘을,
    없으면 이름 글자를 보여 줍니다. 공식 아이콘은 각 회사 브랜드 자료에서 받아 넣으세요."""
    out = []
    for key, pattern in SOCIALS:
        if not site.get(key):
            continue
        name = SOCIAL_LABELS["en"][key]
        path = STATIC / "icons" / f"{key}.svg"
        if path.exists():
            svg = path.read_text(encoding="utf-8")
            svg = re.sub(r"<\?xml.*?\?>|<!--.*?-->", "", svg, flags=re.S).strip()
            svg = re.sub(r'\s(fill|stroke)="(?!none)(#[0-9a-fA-F]{3,8}|black|white|rgb\([^)]*\))"', r' \1="currentColor"', svg)
            svg = re.sub(r"<svg\b[^>]*>", lambda m: re.sub(r'\s(width|height)="[^"]*"', "", m.group(0)).replace("<svg", '<svg class="icon" aria-hidden="true"', 1), svg, count=1)
            inner = svg
        else:
            inner = f'<span class="social__text">{name}</span>'
        label = f"Migan Press on {name}" if lang == "en" else f"미간 {SOCIAL_LABELS['ko'][key]}"
        out.append(f'<a href="{e(social_url(site[key], pattern))}" aria-label="{e(label)}">{inner}</a>')
    return "".join(out)


def curly_quotes(t):
    """곧은 따옴표(' ")를 굽은 따옴표로 바꿉니다. 글자 사이의 '는 ’(don’t), 여는 자리는 ‘ “, 닫는 자리는 ’ ”."""
    t = t.replace("&#x27;", "'").replace("&quot;", '"')
    t = re.sub(r"(?<=\w)'(?=\w)", "’", t)
    t = re.sub(r"(^|[\s(\[>—-])'", "\\1‘", t)
    t = t.replace("'", "’")
    t = re.sub(r'(^|[\s(\[>—-])"', "\\1“", t)
    return t.replace('"', "”")


def kern_brackets(html_text):
    """본문의 곧은 따옴표를 둥근 따옴표로 바꿉니다(낫표 「」『』 간격은 글꼴 그대로 둡니다).
    태그 안(속성)과 script·style 안은 건드리지 않습니다."""
    parts = re.split(r"(<script.*?</script>|<style.*?</style>|<[^>]+>)", html_text, flags=re.S)
    for i in range(0, len(parts), 2):
        parts[i] = curly_quotes(parts[i])
    return "".join(parts)


def layout(site, *, title, body, current="", description="", image="", path="/", lang="ko"):
    body = kern_brackets(body)
    full_title = f"{title} — 미간 MIGAN" if title else "미간 MIGAN"
    if lang == "en":
        full_title = "Migan Press"
    desc = description or site.get("tagline", "").replace("/", " ")
    og_image = f'<meta property="og:image" content="{SITE_URL}{e(image)}">\n' if image else ""
    socials = social_icons(site, lang)
    head_extra = '<meta name="robots" content="noindex">' if path == "/404.html" else f'<link rel="canonical" href="{SITE_URL}{e(path)}">'
    alternates = (
        f'<link rel="alternate" hreflang="ko" href="{SITE_URL}/">\n'
        f'<link rel="alternate" hreflang="en" href="{SITE_URL}/en/">'
        if path in ("/", "/en/") else ""
    )

    def nav_link(href, label, key):
        attr = ' aria-current="page"' if current == key else ""
        return f'<a href="{href}"{attr}>{label}</a>'

    if lang == "en":
        nav_html = '<nav class="site-nav" aria-label="Language"><a class="lang-ko" href="/" lang="ko">한국어</a></nav>'
    else:
        nav_html = f"""<nav class="site-nav" aria-label="주 메뉴">
{nav_link('/books/', '책', 'books')}
{nav_link('/about/', '소개', 'about')}
<a class="lang" href="/en/" aria-label="English" title="English"><svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="9.25"/><ellipse cx="12" cy="12" rx="4" ry="9.25"/><path d="M2.75 12h18.5M4.2 7.2h15.6M4.2 16.8h15.6"/></svg></a>
</nav>"""

    return f"""<!doctype html>
<html lang="{lang}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(full_title)}</title>
<meta name="description" content="{e(desc)}">
<meta property="og:title" content="{e(full_title)}">
<meta property="og:description" content="{e(desc)}">
<meta property="og:url" content="{SITE_URL}{e(path)}">
<meta property="og:type" content="website">
<meta property="og:locale" content="{'en_US' if lang == 'en' else 'ko_KR'}">
{og_image}{head_extra}
{alternates}
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Montserrat:wght@500;600;700&amp;family=Noto+Sans+KR:wght@400;500;700{'&amp;family=Work+Sans:ital,wght@0,400;0,500;0,600;1,400' if lang == 'en' else ''}&amp;display=swap">
<link rel="stylesheet" href="/style.css?v={CSS_VER}">
<link rel="icon" href="/favicon.svg" type="image/svg+xml">
</head>
<body{' class="is-en"' if lang == 'en' else ''}>
<header class="site-header">
<a class="logo" href="/" aria-label="{'Migan Press — Korean homepage' if lang == 'en' else '미간 처음으로'}">{LOGO_HEADER}</a>
{nav_html}
</header>
<main>
{body}
</main>
<footer class="site-footer">
<nav class="social" aria-label="{'Social media' if lang == 'en' else '미간 SNS'}">{socials}</nav>
<div class="site-footer__right">
<p class="copyright">© Migan Press. All rights reserved.</p>
</div>
</footer>
{BLINK_JS}
{EMAIL_JS}
</body>
</html>
"""


def book_row(book):
    return f"""<a class="book-row split" href="/books/{e(book['slug'])}/">
<span class="book-row__title">{book['title_html']}</span>
<span class="book-row__meta"><span class="credit-line">{credit_html(book)}</span><span class="muted">{e(status_label(book))}</span></span>
</a>"""


def book_list(books):
    return '<div class="book-list">' + "\n".join(book_row(b) for b in books) + "</div>"


def cover_card(b, size="m"):
    """표지만 보여 줍니다. 제목·지은이·옮긴이·연도는 표지 안에 있어요."""
    return f'<article class="card">{cover(b, size)}</article>'


def next_book(books):
    """바로 다음에 나올 책: 출간 예정 가운데 가장 이른 책(없으면 가장 최근에 나온 책)."""
    upcoming = sorted((b for b in books if b.get("status") != "published"), key=lambda b: str(b.get("pub_date", "")))
    published = sorted((b for b in books if b.get("status") == "published"), key=lambda b: str(b.get("pub_date", "")), reverse=True)
    return upcoming[0] if upcoming else (published[0] if published else None)


def page_home(site, books):
    feature = next_book(books)

    feature_html = ""
    if feature:
        upcoming_book = feature.get("status") != "published"
        label = "출간 예정 도서" if upcoming_book else "새로 나온 책"
        when = fmt_date(feature.get("pub_date")) + (" 출간 예정" if upcoming_book else " 출간")
        feature_html = f"""<section class="feature split" aria-labelledby="feature-title">
<div class="feature__cover">{cover(feature, 'l', dark=True)}</div>
<div class="feature__text">
<p class="feature__when muted">{e(when)}</p>
<h2 id="feature-title"><a href="/books/{e(feature['slug'])}/">{feature['title_html']}</a></h2>
<p>{e(credit(feature, short=True))}</p>
<div class="feature__tagline prose">{paragraphs(feature.get('tagline') or feature.get('short'))}</div>
<a class="more" href="/books/{e(feature['slug'])}/">자세히 보기</a>
</div>
</section>"""

    cards = "\n".join(cover_card(b, "s") for b in books)
    body = f"""<h1 class="visually-hidden">미간 MIGAN</h1>
{feature_html}
<section class="shelf" aria-labelledby="shelf-title">
<div class="shelf-head"><h2 id="shelf-title">미간에서 나올 책</h2><a class="more" href="/books/">모두 보기</a></div>
<div class="book-grid book-grid--5">
{cards}
</div>
</section>"""
    return layout(site, title="", body=body, path="/")


def page_books(site, books):
    cards = "\n".join(cover_card(b) for b in books)
    body = f"""<section class="page-head split">
<h1 class="visually-hidden">책</h1>
<div class="view-toggle" role="group" aria-label="보기 방식">
<button type="button" data-view="list" aria-pressed="true" aria-label="목록으로 보기" title="목록으로 보기"><svg viewBox="0 0 20 20" aria-hidden="true"><path d="M2 4.5h16M2 10h16M2 15.5h16"/></svg></button>
<button type="button" data-view="grid" aria-pressed="false" aria-label="표지로 보기" title="표지로 보기"><svg viewBox="0 0 20 20" aria-hidden="true"><rect x="2" y="2" width="6.5" height="6.5"/><rect x="11.5" y="2" width="6.5" height="6.5"/><rect x="2" y="11.5" width="6.5" height="6.5"/><rect x="11.5" y="11.5" width="6.5" height="6.5"/></svg></button>
</div>
</section>
<section class="books" data-current="list">
<div class="book-grid">
{cards}
</div>
{book_list(books)}
</section>
<script>
(function () {{
  var root = document.querySelector('.books');
  var buttons = document.querySelectorAll('.view-toggle button');
  function set(view) {{
    root.setAttribute('data-current', view);
    buttons.forEach(function (b) {{ b.setAttribute('aria-pressed', String(b.dataset.view === view)); }});
  }}
  buttons.forEach(function (b) {{ b.addEventListener('click', function () {{ set(b.dataset.view); }}); }});
}})();
</script>"""
    return layout(site, title="책", body=body, current="books", path="/books/")


def download_href(url):
    """구글 드라이브 공유 링크는 바로 내려받는 주소로 바꿉니다. 그 밖의 주소는 그대로 씁니다."""
    url = (url or "").strip()
    if not url:
        return ""
    if not url.startswith("http"):
        return f"/files/{url.lstrip('/').removeprefix('files/')}"
    return drive_download_url(url)


def downloads_html(book):
    """관련 자료 받기: 보도자료와 표지 파일. 시트에 링크가 있는 것만 보여 줍니다."""
    items = []
    if book.get("press_kit"):
        items.append(("보도자료", download_href(book["press_kit"])))
    if book.get("cover_file"):
        items.append(("표지 이미지", download_href(book["cover_file"])))
    if not items:
        return ""
    links = "".join(
        f'<li><a href="{e(href)}" download target="_blank" rel="noopener">{label}<span aria-hidden="true"> ↓</span></a></li>'
        for label, href in items
    )
    return f'<div class="downloads" id="downloads"><p class="caps-ko">관련 자료 받기</p><ul>{links}</ul></div>'


def fmt_size(v):
    """'120*188' → '120×188mm'"""
    v = (v or "").strip()
    return re.sub(r"(\d+)\s*[*xX×]\s*(\d+)(\s*mm)?", lambda m: f"{m.group(1)}×{m.group(2)}mm", v)


def fmt_pages(v):
    """'320' 또는 '320면' → '320면'"""
    v = (v or "").strip()
    return f"{v}면" if re.fullmatch(r"\d+", v) else v


def fmt_price(v):
    """'18000' 또는 '18000원' → '18,000원'"""
    v = (v or "").strip()
    m = re.fullmatch(r"([\d,]+)\s*원?", v)
    return f"{int(m.group(1).replace(',', '')):,}원" if m else v


def page_book(site, books, i):
    book = books[i]
    nxt = books[i + 1] if i + 1 < len(books) else None
    rows = []
    if book.get("original_title"):
        orig = f"<cite>{e(book['original_title'])}</cite>"   # 영문 책 제목은 이탤릭, 부제는 싣지 않음
        if book.get("original_author"):
            orig += f" by {e(names_en(book['original_author']))}"
        rows.append(("원서", orig))
    rows.append(("출간", e(status_label(book))))
    if book.get("size"):
        rows.append(("판형", e(fmt_size(book["size"]))))
    if book.get("binding"):
        rows.append(("사양", e(book["binding"])))
    if book.get("pages"):
        rows.append(("면수", e(fmt_pages(book["pages"]))))
    if book.get("price"):
        rows.append(("정가", e(fmt_price(book["price"]))))
    if book.get("isbn"):
        isbn = book["isbn"] + (f" {book['add_code']}" if book.get("add_code") else "")
        rows.append(("ISBN", e(isbn)))
    stores = [(n, book.get(k)) for n, k in (("교보문고", "kyobo"), ("예스24", "yes24"), ("알라딘", "aladin")) if book.get(k)]
    if stores:
        rows.append(("구입처", ", ".join(f'<a href="{e(u)}">{n}</a>' for n, u in stores)))
    dl = "\n".join(f"<div><dt>{k}</dt><dd>{v}</dd></div>" for k, v in rows)
    note = '<p class="meta-note">출간 이후 상세 정보는 바뀔 수 있습니다.</p>' if book.get("status") != "published" else ""
    downloads = downloads_html(book)
    subtitle = f'<span class="sub">{e(book["subtitle"])}</span>' if book.get("subtitle") else ""
    nxt_title = nxt["title_html"].replace("<br>", ' <br class="br-m">') if nxt else ""
    next_html = (
        f'<a class="next" href="/books/{e(nxt["slug"])}/"><span class="pager__label">다음 책</span><span>{nxt_title}</span></a>'
        if nxt
        else '<a class="next" href="/books/"><span class="pager__label">처음으로</span><span>책 목록</span></a>'
    )
    # 본문 구역: (구역 id, 머리말, 내용). 내용이 있는 구역만 싣고, 왼쪽 목차에도 그 구역만 넣습니다.
    bio = lambda key: f'<div class="prose">{paragraphs(book.get(key))}</div>' if book.get(key) else ""
    sections = [
        ("toc", "차례", toc_html(book.get("toc")), "toc-section"),
        ("praise", "추천의 말", praise_html(book.get("praise")), "praise"),
        ("awards", "수상 내역", awards_html(book.get("awards")), "awards-section"),
        ("author", "지은이", bio("author_bio"), ""),
        ("translator", "옮긴이", bio("translator_bio"), ""),
    ]
    sections = [x for x in sections if x[2].strip()]
    sections_html = "\n".join(book_section(t, inner, cls, sid) for sid, t, inner, cls in sections)
    nav_items = [("intro", "책 소개")] + [(sid, t) for sid, t, _, _ in sections]
    if downloads:
        nav_items.append(("downloads", "관련 자료 받기"))
    section_nav = (
        '<nav class="book-nav" aria-label="이 책의 구역"><ul>'
        + "".join(f'<li><a href="#{sid}">{t}</a></li>' for sid, t in nav_items)
        + "</ul></nav>"
    ) if len(nav_items) > 1 else ""
    body = f"""<article class="book">
<div class="book__cover">{cover(book, 'xl', link=False, dark=True)}</div>
<div class="book__meta">
<dl class="meta">
{dl}
</dl>
{downloads}
{note}
{section_nav}
</div>
<div class="book__text">
<header>
<p class="muted">{e(status_label(book))}</p>
<h1>{e(book['title'])}{subtitle}</h1>
<p>{e(credit(book))}</p>
</header>
<div class="prose" id="intro">
<p class="lead">{rich(book.get('tagline'))}</p>
{intro_html(book)}
</div>
{sections_html}
</div>
</article>
{BOOK_NAV_JS if section_nav else ""}
<nav class="pager" aria-label="다른 책">
<a class="back" href="/books/">← 목록으로</a>
{next_html}
</nav>"""
    return layout(site, title=book["title"], body=body, current="books",
                  description=plain(book.get("short") or book.get("tagline", "")), image=book.get("cover", ""),
                  path=f"/books/{book['slug']}/")


def work_card(w):
    meta = " · ".join(x for x in (w.get("author"), w.get("year")) if x)
    full = ", ".join(x for x in (
        f"{w['author']} 지음" if w.get("author") else "",
        f"{w['translator']} 옮김" if w.get("translator") else "",
        w.get("publisher"), w.get("year"),
    ) if x)
    if w.get("cover"):
        img = f'<img src="{e(w["cover"])}" alt="『{e(w["title"])}』 표지" loading="lazy">'
        cover_html = f'<div class="cover cover--xs cover--image">{img}</div>'
    else:
        cover_html = (
            f'<div class="cover cover--xs">'
            f'<span class="cover__title">{e(w["title"])}</span></div>'
        )
    return f"""<li class="work" title="{e(full)}">
{cover_html}
<span class="work__title">{e(w['title'])}</span>
<span class="muted">{e(meta)}</span>
</li>"""


def about_contact(site):
    """소개글 아래 연락처 표: 찾아오시는 길, 이메일, 인스타그램, 트위터."""
    rows = [("찾아오시는 길", e(site.get("visit") or "아직은 찾아오실 수 없습니다. 문의는 이메일로 부탁드립니다."))]
    if site.get("email"):
        rows.append(("이메일", email_link(site["email"])))
    return contact_table(rows + social_rows(site, "ko"))


PERSON_LINKS_TITLE = "인터뷰·기고 보기"   # 대표 링크 묶음을 펼치는 단추 글자
# 시트 '대표 링크'가 비어 있을 때 쓰는 기본 링크
DEFAULT_PERSON_LINKS = (
    "「언니들이 보내는 행운」 채널예스 | https://m.ch.yes24.com/article/details/46310\n"
    "「2027년으로 보내는 행운의 편지」 『함께가는 여성』 | https://www.womenlink.or.kr/archives/?idx=85742291&bmode=view"
)


def person_links_html(text):
    """대표 소개 아래 바깥 링크(인터뷰 등). 한 줄에 하나, '이름 | 주소' 또는 주소만."""
    items = []
    for line in str(text or "").splitlines():
        line = line.strip()
        if not line:
            continue
        label, _, url = line.rpartition("|") if "|" in line else ("", "", line)
        label, url = label.strip() or url.strip(), url.strip()
        if url.startswith("http"):
            items.append(f'<li><a href="{e(url)}" target="_blank" rel="noopener">{e(label)}<span aria-hidden="true"> ↗</span></a></li>')
    return f'<ul class="person__links" id="person-links-list" hidden>{"".join(items)}</ul>' if items else ""


def page_about(site, works):
    works_html = ""
    if works:
        # '분야' 칸이 있으면 분야별로 묶습니다. 분야 순서는 시트에 처음 나온 순서를 따릅니다.
        groups = {}
        for w in works:
            cat = (w.get("category") or "").strip() or WORK_CATEGORIES.get(w.get("title", "").strip(), "")
            groups.setdefault(cat, []).append(w)
        # 분야가 있는 책(예: 만화)은 제목 없이 같은 목록의 뒤쪽에 이어 붙입니다
        ordered = sorted(groups.items(), key=lambda kv: kv[0] != "")
        lists = '<ul class="works">' + "".join(
            work_card(w).replace('<li class="work"', f'<li class="work" data-group="{gi}"', 1)
            for gi, (_cat, ws) in enumerate(ordered)
            for w in sorted(ws, key=lambda w: -int(re.sub(r"\D", "", w.get("year") or "") or 0))
        ) + "</ul>"
        works_html = f"""<section class="works-block" id="works" aria-label="만든 책" hidden>
{lists}
</section>
"""
    body = f"""<section class="about-motion"><img class="logo-motion" src="/motion/migan-motion-ko-red.svg" alt="미간 로고" width="879" height="583"></section>
{MOTION_JS}
<section class="about-hero split">
<h1 class="about-hero__left">{e(site.get('about_heading'))}</h1>
<div class="about-hero__right"><div class="prose"><p>{e(site.get('about_text'))}</p></div>
{about_contact(site)}</div>
</section>
<section class="person split" aria-labelledby="person-name">
<div class="person__left">
<p class="caps-ko">만드는 사람</p>
<h2 id="person-name">{e(site.get('person_name'))}</h2>
</div>
<div class="person__right"><div class="prose">{paragraphs(site.get('person_bio'))}</div>
{'<button type="button" class="works-toggle" aria-expanded="false" aria-controls="works" data-label="만든 책 보기">만든 책 보기</button>' if works else ''}</div>
</section>
{works_html}
{f'<div class="person-links split"><div></div><div class="person-links__body"><button type="button" class="works-toggle" aria-expanded="false" aria-controls="person-links-list" data-label="{PERSON_LINKS_TITLE}">{PERSON_LINKS_TITLE}</button>{links}</div></div>' if (links := person_links_html(site.get('person_links') or DEFAULT_PERSON_LINKS)) else ''}"""
    body += """
<script>
(function () {
  /* '만든 책 보기', '인터뷰·기고 보기'를 누르면 아래 목록을 펼치고 접습니다 */
  [].slice.call(document.querySelectorAll('.works-toggle')).forEach(function (btn) {
    var list = document.getElementById(btn.getAttribute('aria-controls'));
    if (!list) return;
    btn.addEventListener('click', function () {
      var open = btn.getAttribute('aria-expanded') !== 'true';
      btn.setAttribute('aria-expanded', String(open));
      btn.textContent = open ? '접기' : btn.getAttribute('data-label');
      list.hidden = !open;
    });
  });
})();
</script>"""
    return layout(site, title="소개", body=body, current="about", path="/about/")


def page_404(site):
    body = f"""<section class="not-found">
<div class="not-found__logo" aria-hidden="true">{LOGO_STATIC}</div>
<h1>페이지를 찾을 수 없습니다</h1>
<a class="more" href="/">처음으로</a>
</section>"""
    return layout(site, title="페이지 없음", body=body, path="/404.html")


def page_en(site, books):
    """해외 출판사·에이전시용 영문 소개 한 장: 로고 모션, 소개, 연락처, 책 목록(원제·저자)."""
    text = site.get("en_text") or (
        "Migan is a publisher based in Seoul.\n"
        "Migan (미간, 眉間) takes its name from the Korean word for the space between the eyebrows. "
        "We believe a good story makes us think rather than handing us answers."
    )
    rows = ([("Email", email_link(site["email"]))] if site.get("email") else []) + social_rows(site, "en")
    contact = f'<div class="en-block"><h2 class="caps-ko">Contact</h2>{contact_table(rows)}</div>' if rows else ""
    items = []
    for b in sorted(books, key=lambda b: str(b.get("pub_date", ""))):
        title = (b.get("original_title") or "").strip()
        if not title:
            continue
        author = names_en(b.get("original_author"), short=True)
        items.append(f'<li><cite>{e(title)}</cite><span class="en-books__author">{e(author)}</span></li>')
    books_html = f'<div class="en-block"><h2 class="caps-ko">Forthcoming Titles</h2><ul class="en-books">{"".join(items)}</ul></div>' if items else ""
    body = f"""<section class="about-motion"><img class="logo-motion" src="/motion/migan-motion-en-red.svg" alt="MIGAN logo" width="875" height="564"></section>
{MOTION_JS}
<section class="about-hero about-hero--solo">
<h1 class="visually-hidden">Migan Press</h1>
<div class="about-hero__right"><div class="prose">{paragraphs(text)}</div>
{books_html}
{contact}</div>
</section>"""
    return layout(site, title="Migan Press", body=body, path="/en/", lang="en", description=text)


# 미리보기용: BASE_PATH="/preview-..." 로 빌드하면 모든 내부 링크 앞에 그 주소를 붙이고,
# 검색엔진에 잡히지 않게 막습니다. 평소(본 사이트)에는 비워 둡니다.
BASE_PATH = os.environ.get("BASE_PATH", "").rstrip("/")


def apply_base(html):
    if not BASE_PATH:
        return html
    html = re.sub(r'((?:href|src)=")/(?!/)', rf'\g<1>{BASE_PATH}/', html)
    html = re.sub(r'url\((["\']?)/(?!/)', rf'url(\g<1>{BASE_PATH}/', html)
    html = re.sub(r'<link rel="canonical"[^>]*>', '', html)
    if 'name="robots"' not in html:
        html = html.replace("<head>", '<head>\n<meta name="robots" content="noindex,nofollow">', 1)
    return html


def write(path, html):
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.suffix == ".html":
        html = apply_base(html)
    path.write_text(html, encoding="utf-8")


def main():
    local = sys.argv[2] if len(sys.argv) > 2 and sys.argv[1] == "--local" else None
    site, books, works = load(local)
    if DIST.exists():
        shutil.rmtree(DIST)
    shutil.copytree(STATIC, DIST)
    for b in books:
        b["cover"] = resolve_cover(b, f"book-{b['slug']}")
    nb = next_book(books)
    if nb:
        nb["is_next"] = True   # 표지 목록에서 이 책만 산호색 표지
    for i, w in enumerate(works, 1):
        if not (w.get("cover") or "").strip():
            # 시트 '표지' 칸이 비어 있으면 data/work_covers.json(제목 → 파일)에서 찾습니다
            w["cover"] = WORK_COVERS.get(w.get("title", "").strip(), "")
        w["cover"] = resolve_cover(w, f"work-{i:02d}")
    write(DIST / "index.html", page_home(site, books))
    write(DIST / "books" / "index.html", page_books(site, books))
    for i, book in enumerate(books):
        write(DIST / "books" / book["slug"] / "index.html", page_book(site, books, i))
    write(DIST / "about" / "index.html", page_about(site, works))
    write(DIST / "en" / "index.html", page_en(site, books))
    write(DIST / "404.html", page_404(site))
    # 검색엔진용: 페이지 목록과 안내 파일
    pages = ["/", "/books/", "/about/", "/en/"] + [f"/books/{b['slug']}/" for b in books]
    urls = "".join(f"<url><loc>{SITE_URL}{p}</loc></url>" for p in pages)
    write(DIST / "sitemap.xml", f'<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">{urls}</urlset>\n')
    if BASE_PATH:
        (DIST / "sitemap.xml").unlink()
        write(DIST / "robots.txt", "User-agent: *\nDisallow: /\n")
    else:
        write(DIST / "robots.txt", f"User-agent: *\nAllow: /\nSitemap: {SITE_URL}/sitemap.xml\n")
    print(f"책 {len(books)}권, 만든 책 {len(works)}권으로 사이트를 만들었습니다 -> {DIST}")


if __name__ == "__main__":
    main()
