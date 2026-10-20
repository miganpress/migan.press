"""구글 시트(xlsx 내보내기)를 읽는 작은 도구. 파이썬 표준 라이브러리만 씁니다."""
import re
import urllib.request
import zipfile
from datetime import date, timedelta
from io import BytesIO
from xml.etree import ElementTree as ET

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def fetch_xlsx(sheet_id):
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    req = urllib.request.Request(url, headers={"User-Agent": "migan-site-build"})
    with urllib.request.urlopen(req, timeout=60) as res:
        data = res.read()
    if not data.startswith(b"PK"):
        raise RuntimeError(
            "시트를 읽지 못했습니다. 시트 공유 설정이 '링크가 있는 모든 사용자: 뷰어'인지 확인하세요."
        )
    return data


def _col_index(ref):
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _is_bold(run):
    rpr = run.find("m:rPr", NS)
    if rpr is None:
        return False
    b = rpr.find("m:b", NS)
    return b is not None and b.get("val", "1") not in ("0", "false")


def _is_italic(run):
    rpr = run.find("m:rPr", NS)
    if rpr is None:
        return False
    i = rpr.find("m:i", NS)
    return i is not None and i.get("val", "1") not in ("0", "false")


def _wrap_italic(text):
    return "\n".join(f"__{part}__" if part.strip() else part for part in text.split("\n"))


def _wrap_bold(text):
    # 줄바꿈을 넘어가는 굵은 글씨는 줄마다 따로 감쌉니다
    return "\n".join(f"**{part}**" if part.strip() else part for part in text.split("\n"))


def _text(el):
    """셀 글자를 꺼냅니다. 셀 안에서 일부만 굵게 한 글자는 **굵게**, 기울인 글자는 __기울임__으로 표시해 둡니다."""
    runs = el.findall("m:r", NS)
    if not runs:
        return "".join(t.text or "" for t in el.iter(f"{{{NS['m']}}}t"))
    out = []
    for child in el:
        tag = child.tag.split("}")[1]
        if tag == "t":
            out.append(child.text or "")
        elif tag == "r":
            t = "".join(x.text or "" for x in child.iter(f"{{{NS['m']}}}t"))
            if _is_italic(child):
                t = _wrap_italic(t)
            out.append(_wrap_bold(t) if _is_bold(child) else t)
    return "".join(out).replace("****", "")


def read_workbook(data):
    """{탭 이름: [[셀, ...], ...]} 를 돌려줍니다. 모든 값은 문자열입니다."""
    z = zipfile.ZipFile(BytesIO(data))
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
        shared = [_text(si) for si in root.findall("m:si", NS)]

    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels.findall("pr:Relationship", NS)}
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    sheets = {}
    for s in wb.find("m:sheets", NS).findall("m:sheet", NS):
        rid = s.get(f"{{{NS['r']}}}id")
        target = targets[rid].lstrip("/")
        path = target if target.startswith("xl/") else f"xl/{target}"
        root = ET.fromstring(z.read(path))
        rows = []
        for row in root.iter(f"{{{NS['m']}}}row"):
            cells = {}
            for c in row.findall("m:c", NS):
                t = c.get("t")
                v = c.find("m:v", NS)
                if t == "s" and v is not None:
                    val = shared[int(v.text)]
                elif t == "inlineStr":
                    is_ = c.find("m:is", NS)
                    val = _text(is_) if is_ is not None else ""
                elif v is not None:
                    val = v.text or ""
                    if t is None and re.fullmatch(r"-?\d+\.0", val):
                        val = val[:-2]
                else:
                    val = ""
                cells[_col_index(c.get("r"))] = val
            if cells:
                width = max(cells) + 1
                rows.append([cells.get(i, "") for i in range(width)])
        sheets[s.get("name")] = rows
    return sheets


def records(rows):
    """첫 줄을 머리글로 삼아 dict 목록을 만듭니다. 빈 줄은 건너뜁니다."""
    if not rows:
        return []
    header = [h.strip() for h in rows[0]]
    out = []
    for row in rows[1:]:
        if not any(str(v).strip() for v in row):
            continue
        out.append({h: (row[i].strip() if i < len(row) else "") for i, h in enumerate(header) if h})
    return out


def serial_to_month(value):
    """시트가 '2027-03'을 날짜로 바꿔 버린 경우(일련번호)를 되돌립니다."""
    if re.fullmatch(r"\d{5}(\.\d+)?", value or ""):
        d = date(1899, 12, 30) + timedelta(days=int(float(value)))
        return f"{d.year}-{d.month:02d}"
    return value
