# logic/toc_linker.py
import io, re
import fitz  # PyMuPDF

NUM_AT_END = re.compile(r"(\d+)\s*$")
TOC_LINE_RE = re.compile(
    r"^(?P<title>.*?)(?:[\s\.\u2026·•]*?)(?P<num>\d+)(?:[\s\.\u2026·•]*)$"
)

# FAST MODE: disable label lookups to avoid heavy memory/CPU on small instances
USE_LABELS = False


def parse_range(text: str):
    t = text.strip()
    if not t:
        raise ValueError("Empty TOC range.")
    if "-" in t:
        a, b = t.split("-", 1)
        a, b = int(a), int(b)
        if a < 1 or b < 1 or b < a:
            raise ValueError("Invalid range (use like 6-8).")
        return a - 1, b - 1
    n = int(t)
    if n < 1:
        raise ValueError("Page numbers must be >= 1.")
    return n - 1, n - 1


def _numeric_label(page) -> int | None:
    if not USE_LABELS:
        return None
    try:
        lbl = page.get_label()
        if not lbl:
            return None
        m = NUM_AT_END.search(str(lbl))
        return int(m.group(1)) if m else None
    except Exception:
        return None


def _find_index_by_numeric_label(doc, wanted_num: int, start_from: int) -> int | None:
    if not USE_LABELS:
        return None
    # If you later flip USE_LABELS=True, keep the scan bounded (faster):
    end = min(len(doc), start_from + 50)  # search at most the next 50 pages
    for i in range(max(0, start_from), end):
        n = _numeric_label(doc.load_page(i))
        if n == wanted_num:
            return i
    return None


def _get_lines_dict_sorted(page):
    d = page.get_text("dict")
    lines = []
    for b in d.get("blocks", []):
        if b.get("type") != 0:
            continue
        for ln in b.get("lines", []) or []:
            spans = ln.get("spans", []) or []
            if not spans:
                continue
            text = "".join(s.get("text", "") for s in spans).strip()
            bbox = ln.get("bbox")
            if not bbox:
                continue
            x0, y0, x1, y1 = bbox
            lines.append(
                {
                    "text": text,
                    "bbox": bbox,
                    "x0": x0,
                    "y0": y0,
                    "x1": x1,
                    "y1": y1,
                    "spans": spans,
                }
            )
    lines.sort(key=lambda L: (L["y0"], L["x0"]))
    return lines


def _ends_with_number(text: str) -> bool:
    t = (text or "").strip()
    return bool(re.search(r"\d(?:[\s\.\u2026·•]*)$", t))


def create_links_for_toc(doc, toc_start, toc_end):
    created = 0
    rows = []
    all_nums = []
    LINE_JOIN_Y_GAP = 10
    LEFT_X_TOL = 14

    for pno in range(toc_start, toc_end + 1):
        page = doc.load_page(pno)
        lines = _get_lines_dict_sorted(page)
        for idx, L in enumerate(lines):
            text = L["text"]
            m = TOC_LINE_RE.match(text)
            if not m:
                continue
            if not _ends_with_number(text):
                continue
            try:
                pnum = int(m.group("num"))
            except ValueError:
                continue
            x0, y0, x1, y1 = L["bbox"]
            if idx > 0:
                prev = lines[idx - 1]
                if not _ends_with_number(prev["text"]):
                    if (L["y0"] - prev["y1"]) <= LINE_JOIN_Y_GAP and abs(
                        prev["x0"] - L["x0"]
                    ) <= LEFT_X_TOL:
                        px0, py0, px1, py1 = prev["bbox"]
                        x0, y0, x1, y1 = (
                            min(x0, px0),
                            min(y0, py0),
                            max(x1, px1),
                            max(y1, py1),
                        )
            rows.append((pno, (x0, y0, x1, y1), pnum, m.group("title").strip()))
            all_nums.append(pnum)

    search_start = toc_end + 1
    baseline = min(all_nums) if all_nums else 1

    for pno, bbox, pnum, _title in rows:
        dest_index = _find_index_by_numeric_label(doc, pnum, search_start)
        if dest_index is None:
            dest_index = search_start + (pnum - baseline)
        if 0 <= dest_index < len(doc):
            rect = fitz.Rect(*bbox)
            src_page = doc.load_page(pno)
            src_page.insert_link(
                {"kind": fitz.LINK_GOTO, "from": rect, "page": dest_index, "zoom": 0}
            )
            created += 1
    return created


def process_pdf(infile_fs, toc_range_text: str) -> io.BytesIO:
    toc_start, toc_end = parse_range(toc_range_text)
    infile_fs.stream.seek(0)
    doc = fitz.open(stream=infile_fs.stream.read(), filetype="pdf")
    toc_start = max(0, min(toc_start, len(doc) - 1))
    toc_end = max(0, min(toc_end, len(doc) - 1))
    _ = create_links_for_toc(doc, toc_start, toc_end)
    out = io.BytesIO()
    doc.save(out, garbage=4, deflate=True)
    doc.close()
    out.seek(0)
    return out
