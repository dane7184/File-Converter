"""PDF -> editable formats.

Each converter takes (pdf_path, out_path, pages) where `pages` is a list of
0-based page indexes or None for the whole document, and writes out_path.
"""
from __future__ import annotations

import csv
import io
import math
import re
import zipfile
from pathlib import Path

import pymupdf

FORMATS = {
    "docx": {"label": "Word", "ext": ".docx", "desc": "Editable document, layout preserved"},
    "xlsx": {"label": "Excel", "ext": ".xlsx", "desc": "Tables into worksheets, one sheet per page"},
    "pptx": {"label": "PowerPoint", "ext": ".pptx", "desc": "One slide per page, editable text boxes"},
    "txt": {"label": "Plain text", "ext": ".txt", "desc": "Text only"},
    "md": {"label": "Markdown", "ext": ".md", "desc": "Headings, lists and tables"},
    "html": {"label": "HTML", "ext": ".html", "desc": "Web page with positioned text"},
    "csv": {"label": "CSV", "ext": ".csv", "desc": "Tables as comma-separated rows"},
    "png": {"label": "Images", "ext": ".zip", "desc": "PNG per page, zipped"},
}


class ConversionError(Exception):
    pass


# ---------------------------------------------------------------- helpers

def parse_pages(spec: str | None, page_count: int) -> list[int] | None:
    """'1-3,5' -> [0, 1, 2, 4]. Empty/None -> None (all pages)."""
    if not spec or not spec.strip():
        return None
    out: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        m = re.fullmatch(r"(\d+)\s*(?:-\s*(\d+))?", part)
        if not m:
            raise ConversionError(f"Bad page range: '{part}'. Use e.g. 1-3,5")
        a = int(m.group(1))
        b = int(m.group(2)) if m.group(2) else a
        if a < 1 or b < a:
            raise ConversionError(f"Bad page range: '{part}'")
        if a > page_count:
            raise ConversionError(f"Page {a} is past the end (document has {page_count} pages)")
        out.extend(range(a - 1, min(b, page_count)))
    return sorted(set(out)) or None


def open_pdf(pdf_path, password: str | None = None) -> pymupdf.Document:
    try:
        doc = pymupdf.open(pdf_path)
    except Exception as e:
        raise ConversionError("Could not open this file as a PDF - it may be damaged or not a real PDF.") from e
    if doc.needs_pass:
        if not password or not doc.authenticate(password):
            doc.close()
            raise ConversionError("This PDF is password protected - supply the correct password.")
    return doc


def _iter_pages(doc, pages):
    for i in (pages if pages is not None else range(doc.page_count)):
        yield i, doc[i]


def has_text_layer(pdf_path, password=None) -> bool:
    """False for scanned PDFs (pictures of pages, nothing to extract)."""
    doc = open_pdf(pdf_path, password)
    try:
        for page in doc:
            if page.get_text("text").strip():
                return True
        return False
    finally:
        doc.close()


# PDF base fonts are not installed on Windows/macOS; name the look-alike that is.
_STANDARD_FONTS = {
    "helvetica": "Arial", "helv": "Arial", "arial": "Arial",
    "times": "Times New Roman", "timesnewroman": "Times New Roman", "tiro": "Times New Roman",
    "courier": "Courier New", "couriernew": "Courier New", "cour": "Courier New",
    "symbol": "Symbol",
}


def _font_info(name: str) -> tuple[str, bool, bool]:
    """'ABCDEF+TimesNewRomanPS-BoldMT' -> ('Times New Roman', True, False)."""
    base = re.sub(r"^[A-Z]{6}\+", "", name or "")
    low = base.lower()
    bold = bool(re.search(r"bold|black|heavy|demi", low))
    italic = bool(re.search(r"italic|oblique", low))
    fam = re.split(r"[-,]", base)[0]
    fam = re.sub(r"(PS)?MT$|PS$", "", fam)
    key = fam.lower().replace(" ", "")
    if key not in _STANDARD_FONTS:
        fam = re.sub(r"(Bold|Italic|Oblique|Regular)+$", "", fam) or fam
        key = fam.lower().replace(" ", "")
    if key in _STANDARD_FONTS:
        return _STANDARD_FONTS[key], bold, italic
    fam = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", fam).strip()  # SegoeUI -> Segoe UI
    return fam or "Arial", bold, italic


def _hex(color) -> str:
    """PyMuPDF colour (0xRRGGBB int, or floats 0-1) -> 'RRGGBB'."""
    if isinstance(color, int):
        return f"{color & 0xFFFFFF:06X}"
    if len(color) == 1:  # gray
        color = (color[0],) * 3
    elif len(color) == 4:  # CMYK
        c, m, y, k = color
        color = ((1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))
    return "".join(f"{max(0, min(255, round(v * 255))):02X}" for v in color[:3])


def _page_spans(page) -> list[tuple]:
    """(bbox, size, bold, italic, 'RRGGBB') for every piece of text on the page."""
    spans = []
    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for s in line["spans"]:
                if s["text"].strip():
                    _, fb, fi = _font_info(s["font"])
                    spans.append((s["bbox"], s["size"], bool(s["flags"] & 16) or fb,
                                  bool(s["flags"] & 2) or fi, _hex(s["color"])))
    return spans


def _style_at(spans, bbox):
    """(size, bold, italic, colour) of the text that fills most of bbox, or None."""
    best, area = None, 0.0
    for sb, *style in spans:
        ix = min(sb[2], bbox[2]) - max(sb[0], bbox[0])
        iy = min(sb[3], bbox[3]) - max(sb[1], bbox[1])
        if ix > 0 and iy > 0 and ix * iy > area:
            best, area = style, ix * iy
    return best


def _word_lines(words) -> list[list]:
    """Group words (x0, y0, x1, y1, text, ...) into visual lines, left to right."""
    words = sorted(words, key=lambda w: (round((w[1] + w[3]) / 2), w[0]))
    lines: list[list] = []
    for w in words:
        mid = (w[1] + w[3]) / 2
        if lines and abs(lines[-1][0] - mid) <= max(3, (w[3] - w[1]) * 0.5):
            lines[-1][1].append(w)
        else:
            lines.append([mid, [w]])
    return [sorted(ws, key=lambda w: w[0]) for _, ws in lines]


def _bbox_of(words):
    return (min(w[0] for w in words), min(w[1] for w in words),
            max(w[2] for w in words), max(w[3] for w in words))


def _text_rows(words) -> list[dict]:
    """Lines of text split into cells wherever there is a wide gap."""
    rows = []
    for ws in _word_lines(words):
        segs, cur = [], [ws[0]]
        for w in ws[1:]:
            if w[0] - cur[-1][2] > max(12, (w[3] - w[1]) * 1.2):  # wide gap = new column
                segs.append(cur)
                cur = [w]
            else:
                cur.append(w)
        segs.append(cur)
        rows.append({
            "bbox": _bbox_of(ws),
            "segs": [(s[0][0], s[-1][2], " ".join(w[4] for w in s)) for s in segs],
        })
    return rows


def _align_rows(rows: list[dict]) -> list[list[str]]:
    """Put the cells of neighbouring multi-cell lines into shared columns."""
    out: list[list[str]] = []
    i = 0
    while i < len(rows):
        if len(rows[i]["segs"]) < 2:
            out.append([rows[i]["segs"][0][2]])
            i += 1
            continue
        j = i
        while j < len(rows):
            n = len(rows[j]["segs"])
            sandwiched = (n == 1 and len(rows[j]["segs"][0][2]) <= 40
                          and j + 1 < len(rows) and len(rows[j + 1]["segs"]) >= 2)
            if n < 2 and not sandwiched:
                break
            j += 1
        run = rows[i:j]
        cols: list[list[float]] = []
        for a, b in sorted((s[0], s[1]) for r in run for s in r["segs"]):
            if cols and a < cols[-1][1]:
                cols[-1][1] = max(cols[-1][1], b)
            else:
                cols.append([a, b])
        aligned, clash = [], False
        for r in run:
            cells = [""] * len(cols)
            for a, b, text in r["segs"]:
                mid = (a + b) / 2
                k = next((k for k, c in enumerate(cols) if c[0] <= mid <= c[1]), 0)
                if cells[k]:
                    clash = True
                cells[k] = text
            aligned.append(cells)
        # A line that bridges two columns would glue them together: keep lines as they are.
        out.extend([[s[2] for s in r["segs"]] for r in run] if clash else aligned)
        i = j
    return out


def _edges(values, tol=2.0) -> list[float]:
    out: list[float] = []
    for v in sorted(values):
        if not out or v - out[-1] > tol:
            out.append(v)
    return out


def _nearest(edges, v) -> int:
    return min(range(len(edges)), key=lambda k: abs(edges[k] - v))


def _page_fills(page, drawings) -> list[tuple]:
    """Filled areas on the page as (Rect, 'RRGGBB'), in paint order."""
    fills = []
    page_area = abs(page.rect)
    for d in drawings:
        if d.get("fill") is None or "f" not in d["type"]:
            continue
        r = pymupdf.Rect(d["rect"])
        if r.width < 3 or r.height < 3 or abs(r) > 0.9 * page_area:
            continue
        fills.append((r, _hex(d["fill"])))
    return fills


def _page_layout(page) -> list[tuple]:
    """Page content top to bottom: ('table', {...}) and ('text', [row, ...]) items."""
    words = page.get_text("words")
    spans = _page_spans(page)
    try:
        drawings = page.get_drawings()
    except Exception:
        drawings = []
    fills = _page_fills(page, drawings)
    try:
        found = list(page.find_tables().tables)
    except Exception:
        found = []

    tables, in_table = [], set()
    for t in found:
        boxes = [c for c in t.cells if c]
        if not boxes:
            continue
        tb = pymupdf.Rect(t.bbox) + (-2, -2, 2, 2)
        inside = [k for k, w in enumerate(words)
                  if tb.contains(pymupdf.Point((w[0] + w[2]) / 2, (w[1] + w[3]) / 2))]
        xs = _edges([b[0] for b in boxes] + [b[2] for b in boxes])
        ys = _edges([b[1] for b in boxes] + [b[3] for b in boxes])
        cells = []
        for b in boxes:
            c0, c1 = _nearest(xs, b[0]), _nearest(xs, b[2])
            r0, r1 = _nearest(ys, b[1]), _nearest(ys, b[3])
            if c1 <= c0 or r1 <= r0:
                continue
            mine = [words[k] for k in inside
                    if b[0] <= (words[k][0] + words[k][2]) / 2 <= b[2]
                    and b[1] <= (words[k][1] + words[k][3]) / 2 <= b[3]]
            cell = {"r0": r0, "c0": c0, "r1": r1, "c1": c1, "text": "", "style": None, "align": None, "fill": None}
            if mine:
                cell["text"] = " ".join(" ".join(w[4] for w in ln) for ln in _word_lines(mine))
                tb_ = _bbox_of(mine)
                cell["style"] = _style_at(spans, tb_)
                left, right = tb_[0] - b[0], b[2] - tb_[2]
                if abs(left - right) <= 2 and left > 4:
                    cell["align"] = "center"
                elif right < left - 2:
                    cell["align"] = "right"
                else:
                    cell["align"] = "left"
            centre = pymupdf.Point((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)
            for rect, colour in fills:
                if rect.contains(centre):
                    cell["fill"] = colour
            if cell["fill"] == "FFFFFF":
                cell["fill"] = None
            cells.append(cell)
        if not any(c["text"] for c in cells):
            continue
        ruled = any("s" in d["type"] and pymupdf.Rect(d["rect"]).intersects(tb) for d in drawings)
        in_table.update(inside)
        tables.append((t.bbox[1], {"rows": len(ys) - 1, "cols": len(xs) - 1, "cells": cells, "ruled": ruled}))

    rows = _text_rows([w for k, w in enumerate(words) if k not in in_table])
    for r in rows:
        r["style"] = _style_at(spans, r["bbox"])
        b = r["bbox"]
        centre = pymupdf.Point((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)
        r["fill"] = None
        for rect, colour in fills:
            if rect.contains(centre):
                r["fill"] = colour
        if r["fill"] == "FFFFFF":
            r["fill"] = None
        if r["style"] and not r["fill"]:
            # Pale text on a background we could not carry over would be unreadable.
            red, green, blue = (int(r["style"][3][n:n + 2], 16) for n in (0, 2, 4))
            if 0.299 * red + 0.587 * green + 0.114 * blue > 200:
                r["style"] = [*r["style"][:3], "000000"]

    tables.sort(key=lambda t: t[0])
    items, pending = [], []
    for r in rows:
        while tables and tables[0][0] < r["bbox"][1]:
            if pending:
                items.append(("text", pending))
                pending = []
            items.append(("table", tables.pop(0)[1]))
        pending.append(r)
    if pending:
        items.append(("text", pending))
    items.extend(("table", t[1]) for t in tables)
    return items


def _table_grid(table) -> list[list[str]]:
    grid = [[""] * table["cols"] for _ in range(table["rows"])]
    for c in table["cells"]:
        grid[c["r0"]][c["c0"]] = c["text"]
    return grid


_CURRENCY = "$€£¥₹៛"


def _coerce(value: str):
    """'1,234.50' / '$48,500.00' / '12%' / '(7)' -> (number, Excel format) so Excel can calculate.

    Anything that is not clearly a number comes back unchanged with format None.
    """
    s = value.strip()
    keep = (value, None)
    if not s or len(s) > 24:
        return keep
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg, s = True, s[1:-1].strip()
    if s and s[0] in "-−–":
        neg, s = not neg, s[1:].strip()
    symbol, suffix = "", False
    if s and s[0] in _CURRENCY:
        symbol, s = s[0], s[1:].strip()
    elif s and s[-1] in _CURRENCY:
        symbol, suffix, s = s[-1], True, s[:-1].strip()
    if s and s[0] in "-−–":
        neg, s = not neg, s[1:].strip()
    percent = s.endswith("%")
    if percent:
        s = s[:-1].strip()
    m = re.fullmatch(r"(\d{1,3}(?:,\d{3})+|\d+)?(\.\d+)?", s)
    if not m or not (m.group(1) or m.group(2)):
        return keep
    whole, frac = m.group(1) or "0", m.group(2) or ""
    grouped = "," in whole
    digits = whole.replace(",", "")
    if len(digits) > 15:  # ID numbers: Excel would round them
        return keep
    if len(digits) > 1 and digits.startswith("0") and not (symbol or percent):  # codes, phone numbers
        return keep
    n = float(digits + frac) if frac else int(digits)
    if neg:
        n = -n
    decimals = "." + "0" * (len(frac) - 1) if frac else ""
    if percent:
        return n / 100, "0" + decimals + "%"
    if symbol:
        quoted = f'"{symbol}"'
        body = "#,##0" + decimals
        return n, (body + quoted) if suffix else (quoted + body)
    if grouped:
        return n, "#,##0" + decimals
    return n, ("0" + decimals) if frac else None


# ------------------------------------------------------------- converters

def to_docx(pdf_path, out_path, pages=None, password=None):
    import logging
    import warnings

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", FutureWarning)
        from pdf2docx import Converter
    logging.getLogger().setLevel(logging.WARNING)  # pdf2docx logs every page at INFO

    cv = Converter(str(pdf_path), password=password)
    try:
        cv.convert(str(out_path), pages=pages)
    finally:
        cv.close()
    _fix_docx_fonts(out_path)


def _fix_docx_fonts(path):
    """Replace PDF font names (Helvetica, ABCDEF+Calibri-Bold) with installed families."""
    from docx import Document
    from docx.oxml.ns import qn

    doc = Document(str(path))
    attrs = [qn(f"w:{a}") for a in ("ascii", "hAnsi", "eastAsia", "cs")]
    roots = [doc.element, doc.styles.element]
    for section in doc.sections:
        for part in (section.header, section.footer):
            roots.append(part._element)
    changed = False
    for root in roots:
        for el in root.iter(qn("w:rFonts")):
            for a in attrs:
                name = el.get(a)
                if name:
                    family = _font_info(name)[0]
                    if family != name:
                        el.set(a, family)
                        changed = True
    # A fixed line height smaller than the text cuts the tops off the letters in Word.
    for p in doc.element.iter(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        spacing = ppr.find(qn("w:spacing")) if ppr is not None else None
        if spacing is None or spacing.get(qn("w:lineRule")) != "exact":
            continue
        sizes = [int(sz.get(qn("w:val"))) for sz in p.iter(qn("w:sz")) if (sz.get(qn("w:val")) or "").isdigit()]
        line = spacing.get(qn("w:line")) or ""
        if sizes and line.isdigit() and int(line) < max(sizes) * 10 * 0.95:  # half-points -> twentieths
            spacing.set(qn("w:lineRule"), "atLeast")
            changed = True
    if changed:
        doc.save(str(path))


def to_xlsx(pdf_path, out_path, pages=None, password=None):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    doc = open_pdf(pdf_path, password)
    wb = Workbook()
    wb.remove(wb.active)
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def font_for(style, base_size=None):
        if not style:
            return None
        size, bold, italic, colour = style
        kw = {"bold": bold, "italic": italic}
        if colour != "000000":
            kw["color"] = colour
        if base_size is None and size >= 12:  # headings keep their size
            kw["size"] = round(size)
        return Font(**kw) if (bold or italic or len(kw) > 2) else None

    try:
        for i, page in _iter_pages(doc, pages):
            ws = wb.create_sheet(f"Page {i + 1}")
            widths: dict[int, int] = {}
            r = 1
            prev_kind = None
            for kind, payload in _page_layout(page):
                if kind == "table":
                    if prev_kind is not None:
                        r += 1  # blank row before the table
                    for c in payload["cells"]:
                        row, col = r + c["r0"], 1 + c["c0"]
                        value, fmt = _coerce(c["text"]) if c["text"] else (None, None)
                        cell = ws.cell(row=row, column=col, value=value)
                        if fmt:
                            cell.number_format = fmt
                        font = font_for(c["style"], base_size=10)
                        if font:
                            cell.font = font
                        horizontal = c["align"]
                        if horizontal == "right" and not isinstance(value, str):
                            horizontal = None  # numbers sit on the right anyway
                        if horizontal == "left" and isinstance(value, str):
                            horizontal = None
                        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal=horizontal)
                        spans_cols, spans_rows = c["c1"] - c["c0"], c["r1"] - c["r0"]
                        for rr in range(row, row + spans_rows):
                            for cc in range(col, col + spans_cols):
                                box = ws.cell(row=rr, column=cc)
                                if payload["ruled"]:
                                    box.border = border
                                if c["fill"]:
                                    box.fill = PatternFill("solid", fgColor=c["fill"])
                        if spans_cols > 1 or spans_rows > 1:
                            ws.merge_cells(start_row=row, start_column=col,
                                           end_row=row + spans_rows - 1, end_column=col + spans_cols - 1)
                        elif value is not None:
                            shown = c["text"] if fmt else str(value)
                            widths[col] = max(widths.get(col, 0), len(shown))
                    r += payload["rows"]
                else:
                    if prev_kind is not None:
                        r += 1  # blank row after the table
                    last_bottom = None
                    for row_info, cells in zip(payload, _align_rows(payload)):
                        top, bottom = row_info["bbox"][1], row_info["bbox"][3]
                        if last_bottom is not None and top - last_bottom > (bottom - top) * 0.9:
                            r += 1  # keep the gap between paragraphs
                        last_bottom = bottom
                        font = font_for(row_info["style"])
                        for ci, text in enumerate(cells, start=1):
                            if not text:
                                continue
                            value, fmt = _coerce(text) if len(cells) > 1 else (text, None)
                            cell = ws.cell(row=r, column=ci, value=value)
                            if fmt:
                                cell.number_format = fmt
                            if font:
                                cell.font = font
                            if row_info["fill"]:
                                cell.fill = PatternFill("solid", fgColor=row_info["fill"])
                            if len(cells) > 1:
                                widths[ci] = max(widths.get(ci, 0), len(text))
                        if row_info["style"] and row_info["style"][0] >= 12:
                            ws.row_dimensions[r].height = round(row_info["style"][0] * 1.4, 1)
                        r += 1
                prev_kind = kind
            for col, width in widths.items():
                ws.column_dimensions[get_column_letter(col)].width = min(max(10, width + 3), 60)
        if not wb.sheetnames:
            wb.create_sheet("Sheet1")
        wb.save(out_path)
    finally:
        doc.close()


def to_csv(pdf_path, out_path, pages=None, password=None):
    doc = open_pdf(pdf_path, password)
    try:
        with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            for _, page in _iter_pages(doc, pages):
                items = _page_layout(page)
                tables = [p for kind, p in items if kind == "table"]
                if tables:
                    for t in tables:
                        w.writerows(_table_grid(t))
                else:
                    for kind, rows in items:
                        w.writerows(_align_rows(rows))
    finally:
        doc.close()


# PowerPoint ---------------------------------------------------------------

_EMU = 12700  # per PDF point
_MAX_NATIVE_SHAPES = 600  # busier pages get their line art as one background picture


def _flatten(path) -> list[tuple[list, bool]]:
    """A PyMuPDF drawing -> [(points, closed)] with curves turned into short lines."""
    subs: list[list] = []
    closed: list[bool] = []

    def start(p):
        subs.append([(p.x, p.y)])
        closed.append(False)

    def extend(p):
        subs[-1].append((p.x, p.y))

    for item in path["items"]:
        kind = item[0]
        if kind == "re":
            r = item[1]
            subs.append([(r.x0, r.y0), (r.x1, r.y0), (r.x1, r.y1), (r.x0, r.y1)])
            closed.append(True)
        elif kind == "qu":
            q = item[1]
            subs.append([(q.ul.x, q.ul.y), (q.ur.x, q.ur.y), (q.lr.x, q.lr.y), (q.ll.x, q.ll.y)])
            closed.append(True)
        elif kind in ("l", "c"):
            p0 = item[1]
            joined = (subs and not closed[-1]
                      and abs(subs[-1][-1][0] - p0.x) < 0.05 and abs(subs[-1][-1][1] - p0.y) < 0.05)
            if not joined:
                start(p0)
            if kind == "l":
                extend(item[2])
            else:
                _, a, b, c, d = item
                for k in range(1, 9):
                    t = k / 8
                    u = 1 - t
                    extend(pymupdf.Point(
                        u ** 3 * a.x + 3 * u * u * t * b.x + 3 * u * t * t * c.x + t ** 3 * d.x,
                        u ** 3 * a.y + 3 * u * u * t * b.y + 3 * u * t * t * c.y + t ** 3 * d.y))
    filled = path.get("fill") is not None
    out = []
    for pts, was_closed in zip(subs, closed):
        loop = len(pts) > 2 and abs(pts[0][0] - pts[-1][0]) < 0.05 and abs(pts[0][1] - pts[-1][1]) < 0.05
        if loop:
            pts = pts[:-1]
        out.append((pts, was_closed or loop or (filled and len(pts) > 2) or bool(path.get("closePath"))))
    return out


def _is_ellipse(path) -> bool:
    items = path["items"]
    if len(items) != 4 or any(it[0] != "c" for it in items):
        return False
    r = pymupdf.Rect(path["rect"])
    cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
    mids = [(cx, r.y0), (cx, r.y1), (r.x0, cy), (r.x1, cy)]
    return all(any(abs(it[1].x - mx) < 0.6 and abs(it[1].y - my) < 0.6 for mx, my in mids) for it in items)


def _line_art_png(doc, pno, dpi=150) -> bytes:
    """The page's vector graphics only (no text, no images) as a picture."""
    tmp = pymupdf.open()
    try:
        tmp.insert_pdf(doc, from_page=pno, to_page=pno)
        page = tmp[0]
        page.add_redact_annot(page.rect, fill=False)
        page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_REMOVE,
                              graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                              text=pymupdf.PDF_REDACT_TEXT_REMOVE)
        return page.get_pixmap(dpi=dpi).tobytes("png")
    finally:
        tmp.close()


def _stacks(prev, line) -> bool:
    """True if `line` is the next line of the same paragraph as `prev`."""
    p, b = prev["bbox"], line["bbox"]
    height = p[3] - p[1]
    if b[1] < p[1] + 0.5 * height or b[1] - p[3] > 0.8 * height:
        return False
    if min(p[2], b[2]) - max(p[0], b[0]) <= 0:
        return False
    s1 = max(s["size"] for s in prev["spans"])
    s2 = max(s["size"] for s in line["spans"])
    return max(s1, s2) / min(s1, s2) < 1.25


def _paragraphs(lines) -> list[tuple[list, str]]:
    """Split a text block into runs of lines that share one alignment: (lines, how).

    `how` is 'left', 'right', 'center', 'justify' or 'rotated'.
    """
    groups: list[list] = []
    for line in lines:
        horizontal = abs(line["dir"][0] - 1) < 0.01 and abs(line["dir"][1]) < 0.01
        if not horizontal:
            groups.append([line, None])  # marker: rotated
            continue
        if groups and groups[-1][-1] is not None and _stacks(groups[-1][-1], line):
            groups[-1].append(line)
        else:
            groups.append([line])
    out = []
    for g in groups:
        if g[-1] is None:
            out.append(([g[0]], "rotated"))
            continue
        if len(g) == 1:
            out.append((g, "left"))
            continue
        x0s = [ln["bbox"][0] for ln in g]
        x1s = [ln["bbox"][2] for ln in g]
        mids = [(a + b) / 2 for a, b in zip(x0s, x1s)]
        tol = 1.5
        left = max(x0s) - min(x0s) <= tol
        full = [x for x in x1s[:-1] if max(x1s) - x <= tol]
        if left and len(full) == len(g) - 1 and len(g) >= 3:
            out.append((g, "justify"))
        elif left:
            out.append((g, "left"))
        elif max(x1s) - min(x1s) <= tol:
            out.append((g, "right"))
        elif max(mids) - min(mids) <= tol:
            out.append((g, "center"))
        else:  # ragged on both sides: place every line on its own
            out.extend(([ln], "left") for ln in g)
    return out


def to_pptx(pdf_path, out_path, pages=None, password=None):
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.dml import MSO_LINE
    from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
    from pptx.enum.text import MSO_AUTO_SIZE, PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Emu, Pt

    doc = open_pdf(pdf_path, password)
    prs = Presentation()
    blank = prs.slide_layouts[6]
    align = {"left": PP_ALIGN.LEFT, "right": PP_ALIGN.RIGHT, "center": PP_ALIGN.CENTER,
             "justify": PP_ALIGN.JUSTIFY}

    def set_alpha(parent, opacity):
        if opacity is None or opacity >= 0.995:
            return
        clr = parent.find(f"{qn('a:solidFill')}/{qn('a:srgbClr')}") if parent is not None else None
        if clr is not None:
            clr.append(clr.makeelement(qn("a:alpha"), {"val": str(int(max(opacity, 0) * 100000))}))

    def paint(shape, path, k, closed=True):
        """Give a shape the PDF path's fill and outline instead of the theme's."""
        style = shape._element.find(qn("p:style"))
        if style is not None:
            shape._element.remove(style)
        if closed and path.get("fill") is not None and "f" in path["type"]:
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string(_hex(path["fill"]))
            set_alpha(shape._element.spPr, path.get("fill_opacity"))
        else:
            shape.fill.background()
        if path.get("color") is not None and "s" in path["type"]:
            stroke(shape.line, path, k)
        else:
            shape.line.fill.background()

    def stroke(line, path, k):
        line.color.rgb = RGBColor.from_string(_hex(path["color"]))
        line.width = Pt(max(round((path.get("width") or 1) * k, 2), 0.25))
        dashes = (path.get("dashes") or "").replace(" ", "")
        if dashes and not dashes.startswith("[]"):
            line.dash_style = MSO_LINE.DASH
        set_alpha(line._ln, path.get("stroke_opacity"))

    try:
        first = doc[pages[0] if pages else 0]
        # PowerPoint allows slides from 1 to 56 inches.
        shrink = min(1.0, 4032 / max(first.rect.width, first.rect.height))
        slide_w = max(first.rect.width * shrink, 72)
        slide_h = max(first.rect.height * shrink, 72)
        prs.slide_width = Emu(int(slide_w * _EMU))
        prs.slide_height = Emu(int(slide_h * _EMU))

        for pno, page in _iter_pages(doc, pages):
            slide = prs.slides.add_slide(blank)
            shapes = slide.shapes
            # Fit this page on the slide without distorting it.
            k = min(slide_w / page.rect.width, slide_h / page.rect.height)
            ox = (slide_w - page.rect.width * k) / 2 - page.rect.x0 * k
            oy = (slide_h - page.rect.height * k) / 2 - page.rect.y0 * k

            def X(x):
                return int(round((ox + x * k) * _EMU))

            def Y(y):
                return int(round((oy + y * k) * _EMU))

            def L(v):
                return max(int(round(v * k * _EMU)), 1)

            data = page.get_text("dict")
            page_area = abs(page.rect)
            xrefs = {}
            try:
                for info in page.get_image_info(xrefs=True):
                    if info.get("xref"):
                        xrefs[tuple(round(v) for v in info["bbox"])] = info["xref"]
            except Exception:
                pass

            def add_image(block):
                x0, y0, x1, y1 = block["bbox"]
                if x1 - x0 < 0.5 or y1 - y0 < 0.5:
                    return
                try:
                    blob = None
                    xref = xrefs.get(tuple(round(v) for v in block["bbox"]))
                    smask = doc.extract_image(xref).get("smask", 0) if xref else 0
                    if smask:  # transparency lives in a separate mask image
                        pix = pymupdf.Pixmap(doc, xref)
                        if pix.colorspace is not None and pix.colorspace.n > 3:
                            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                        try:
                            pix = pymupdf.Pixmap(pix, pymupdf.Pixmap(doc, smask))
                        except Exception:
                            pass
                        blob = pix.tobytes("png")
                    elif block.get("ext") in ("png", "jpeg", "jpg") and block.get("colorspace", 3) <= 3:
                        blob = block["image"]  # keep the original, no re-encoding
                    else:
                        pix = pymupdf.Pixmap(block["image"])
                        if pix.alpha or pix.colorspace is None or pix.colorspace.n > 3:
                            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                        blob = pix.tobytes("png")
                    shapes.add_picture(io.BytesIO(blob), X(x0), Y(y0), L(x1 - x0), L(y1 - y0))
                except Exception:
                    try:  # last resort: a picture of that part of the page
                        pix = page.get_pixmap(clip=pymupdf.Rect(block["bbox"]), dpi=150)
                        shapes.add_picture(io.BytesIO(pix.tobytes("png")), X(x0), Y(y0), L(x1 - x0), L(y1 - y0))
                    except Exception:
                        pass

            images = [b for b in data["blocks"] if b["type"] == 1]
            backdrops = [b for b in images
                         if (b["bbox"][2] - b["bbox"][0]) * (b["bbox"][3] - b["bbox"][1]) > 0.5 * page_area]
            for b in backdrops:
                add_image(b)

            # ---- lines, boxes, fills
            try:
                drawings = page.get_drawings()
            except Exception:
                drawings = []
            try:
                shaded = any(entry[0] == "fill-shade" for entry in page.get_bboxlog())
            except Exception:
                shaded = False
            if drawings and (shaded or len(drawings) > _MAX_NATIVE_SHAPES):
                try:
                    shapes.add_picture(io.BytesIO(_line_art_png(doc, pno)), X(page.rect.x0), Y(page.rect.y0),
                                       L(page.rect.width), L(page.rect.height))
                    drawings = []
                except Exception:
                    pass
            for path in drawings:
                r = pymupdf.Rect(path["rect"])
                if not r.intersects(page.rect + (-1, -1, 1, 1)) and not (r.width == 0 or r.height == 0):
                    continue
                has_fill = path.get("fill") is not None and "f" in path["type"]
                has_stroke = path.get("color") is not None and "s" in path["type"]
                if not has_fill and not has_stroke:
                    continue
                if has_fill and not has_stroke and _hex(path["fill"]) == "FFFFFF" and abs(r) > 0.95 * page_area:
                    continue  # white page background
                try:
                    items = path["items"]
                    if len(items) == 1 and items[0][0] == "re" and r.width > 0 and r.height > 0:
                        shp = shapes.add_shape(MSO_SHAPE.RECTANGLE, X(r.x0), Y(r.y0), L(r.width), L(r.height))
                        paint(shp, path, k)
                        continue
                    if _is_ellipse(path):
                        shp = shapes.add_shape(MSO_SHAPE.OVAL, X(r.x0), Y(r.y0), L(r.width), L(r.height))
                        paint(shp, path, k)
                        continue
                    subs = _flatten(path)
                    flat = r.width < 0.05 or r.height < 0.05
                    if flat or (not has_fill and all(len(p) == 2 for p, _ in subs)):
                        if not has_stroke:
                            continue
                        for pts, _ in subs:
                            for (ax, ay), (bx, by) in zip(pts, pts[1:]):
                                conn = shapes.add_connector(MSO_CONNECTOR.STRAIGHT, X(ax), Y(ay), X(bx), Y(by))
                                style = conn._element.find(qn("p:style"))
                                if style is not None:
                                    conn._element.remove(style)
                                stroke(conn.line, path, k)
                        continue
                    subs = [(p, c) for p, c in subs if len(p) >= 2]
                    if not subs:
                        continue
                    builder = None
                    for pts, is_closed in subs:
                        emu = [(X(x), Y(y)) for x, y in pts]
                        if builder is None:
                            builder = shapes.build_freeform(emu[0][0], emu[0][1])
                        else:
                            builder.move_to(emu[0][0], emu[0][1])
                        builder.add_line_segments(emu[1:], close=is_closed)
                    shp = builder.convert_to_shape()
                    paint(shp, path, k, closed=any(c for _, c in subs))
                except Exception:
                    continue

            for b in images:
                if b not in backdrops:
                    add_image(b)

            # ---- text
            def add_runs(para, line, joiner=""):
                if joiner and para.runs:
                    para.runs[-1].text += joiner
                for span in line["spans"]:
                    if not span["text"]:
                        continue
                    run = para.add_run()
                    run.text = span["text"]
                    family, name_bold, name_italic = _font_info(span["font"])
                    font = run.font
                    font.name = family
                    font.size = Pt(max(round(span["size"] * k, 1), 1))
                    font.bold = bool(span["flags"] & 16) or name_bold
                    font.italic = bool(span["flags"] & 2) or name_italic
                    font.color.rgb = RGBColor.from_string(_hex(span["color"]))
                    if any(ord(ch) > 0x2FF for ch in span["text"]):  # non-Latin scripts
                        latin = run._r.rPr.find(qn("a:latin"))
                        if latin is not None:
                            latin.addnext(latin.makeelement(qn("a:cs"), {"typeface": family}))
                            latin.addnext(latin.makeelement(qn("a:ea"), {"typeface": family}))

            def new_box(left, top, width, height, wrap):
                box = shapes.add_textbox(left, top, max(width, 1), max(height, 1))
                tf = box.text_frame
                tf.word_wrap = wrap
                tf.auto_size = MSO_AUTO_SIZE.NONE
                tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                return box, tf

            for block in data["blocks"]:
                if block["type"] != 0:
                    continue
                lines = [ln for ln in block.get("lines", []) if any(s["text"].strip() for s in ln["spans"])]
                for group, how in _paragraphs(lines):
                    x0 = min(ln["bbox"][0] for ln in group)
                    x1 = max(ln["bbox"][2] for ln in group)
                    y0 = min(ln["bbox"][1] for ln in group)
                    y1 = max(ln["bbox"][3] for ln in group)
                    first_line = group[0]
                    size = max(s["size"] for s in first_line["spans"])
                    baseline = first_line["spans"][0]["origin"][1]

                    if how == "rotated":
                        dx, dy = first_line["dir"]
                        angle = math.degrees(math.atan2(dy, dx)) % 360
                        upright = abs(abs(dy) - 1) < 0.01
                        length = (y1 - y0) if upright else math.hypot(x1 - x0, y1 - y0)
                        thick = (x1 - x0) if upright else size * 1.2
                        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
                        box, tf = new_box(X(cx - length / 2), Y(cy - thick / 2), L(length), L(thick), False)
                        box.rotation = angle
                        add_runs(tf.paragraphs[0], first_line)
                        continue

                    if len(group) == 1:
                        top = min(baseline - 0.93 * size, y0 + 0.2 * size)
                        box, tf = new_box(X(x0), Y(top), L(x1 - x0), L(y1 - top), False)
                        add_runs(tf.paragraphs[0], first_line)
                        continue

                    pitch = (group[-1]["spans"][0]["origin"][1] - baseline) / (len(group) - 1)
                    top = baseline - 0.8 * pitch
                    wrap = how == "justify"
                    width = (x1 - x0) * (1.01 if wrap else 1.0)
                    box, tf = new_box(X(x0), Y(top), L(width), L(max(y1 - top, pitch * len(group))), wrap)
                    para = tf.paragraphs[0]
                    para.alignment = align[how]
                    if pitch > 0:
                        para.line_spacing = Pt(round(pitch * k, 2))
                    for li, line in enumerate(group):
                        if li and not wrap:
                            para.add_line_break()
                        ended = group[li - 1]["spans"][-1]["text"][-1:] if li else " "
                        add_runs(para, line, joiner=" " if (wrap and li and ended not in " -") else "")
        prs.save(out_path)
    finally:
        doc.close()


def to_txt(pdf_path, out_path, pages=None, password=None):
    doc = open_pdf(pdf_path, password)
    try:
        parts = [page.get_text("text", sort=True).rstrip() for _, page in _iter_pages(doc, pages)]
        Path(out_path).write_text("\n\n\f".join(parts) + "\n", encoding="utf-8")
    finally:
        doc.close()


def to_md(pdf_path, out_path, pages=None, password=None):
    doc = open_pdf(pdf_path, password)
    try:
        try:
            import pymupdf4llm

            md = pymupdf4llm.to_markdown(doc, pages=pages, show_progress=False)
        except Exception:
            md = "\n\n---\n\n".join(page.get_text("text", sort=True).strip() for _, page in _iter_pages(doc, pages))
        Path(out_path).write_text(md, encoding="utf-8")
    finally:
        doc.close()


def to_html(pdf_path, out_path, pages=None, password=None):
    doc = open_pdf(pdf_path, password)
    try:
        body = "\n".join(page.get_text("html") for _, page in _iter_pages(doc, pages))
        title = Path(pdf_path).stem
        html = (
            "<!doctype html>\n<html><head><meta charset=\"utf-8\">"
            f"<title>{title}</title>"
            "<style>body{background:#eee;margin:0;padding:16px}"
            "body>div{background:#fff;margin:0 auto 16px;box-shadow:0 1px 4px rgba(0,0,0,.2)}</style>"
            f"</head><body contenteditable=\"true\">\n{body}\n</body></html>"
        )
        Path(out_path).write_text(html, encoding="utf-8")
    finally:
        doc.close()


def to_png(pdf_path, out_path, pages=None, password=None, dpi=150):
    doc = open_pdf(pdf_path, password)
    try:
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_STORED) as z:
            for i, page in _iter_pages(doc, pages):
                z.writestr(f"page-{i + 1:03d}.png", page.get_pixmap(dpi=dpi).tobytes("png"))
    finally:
        doc.close()


_CONVERTERS = {
    "docx": to_docx, "xlsx": to_xlsx, "pptx": to_pptx, "txt": to_txt,
    "md": to_md, "html": to_html, "csv": to_csv, "png": to_png,
}


def convert(pdf_path, fmt: str, out_path=None, pages: str | None = None, password: str | None = None) -> Path:
    """Convert pdf_path to fmt. `pages` is a human range like '1-3,5'. Returns the output path."""
    fmt = fmt.lower().lstrip(".")
    if fmt not in _CONVERTERS:
        raise ConversionError(f"Unknown format '{fmt}'. Choose from: {', '.join(FORMATS)}")
    pdf_path = Path(pdf_path)
    if not pdf_path.is_file():
        raise ConversionError(f"File not found: {pdf_path}")
    out_path = Path(out_path) if out_path else pdf_path.with_suffix(FORMATS[fmt]["ext"])

    doc = open_pdf(pdf_path, password)
    try:
        if doc.page_count == 0:
            raise ConversionError("The PDF has no pages.")
        page_list = parse_pages(pages, doc.page_count)
    finally:
        doc.close()

    try:
        _CONVERTERS[fmt](pdf_path, out_path, page_list, password)
    except ConversionError:
        raise
    except Exception as e:
        raise ConversionError(f"Conversion to {FORMATS[fmt]['label']} failed: {e}") from e
    return out_path
