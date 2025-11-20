import io, re
from datetime import datetime
import streamlit as st
from openpyxl import load_workbook
from pdfminer.high_level import extract_text
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from pypdf import PdfReader, PdfWriter, Transformation
from pypdf._page import PageObject

# ---- Config ----
SIDE_MARGIN_MM = 2
TOP_MARGIN_MM = 4
STAMP_BOTTOM_MM = 12
INTER_GAP_MM = 1

BASE_CROP_L = 6
BASE_CROP_R = 6
BASE_CROP_T = 8
BASE_CROP_B = 8

LOW_TEXT_LINES = 4
SHORT_TEXT_CHARS = 80
EXTRA_CROP_LR = 14
EXTRA_CROP_T  = 18
EXTRA_CROP_B  = 28

def strip_diacritics(s: str) -> str:
    import unicodedata
    if s is None:
        return ""
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if ord(c) < 128)

def read_excel_lookup(file_like):
    """
    Czyta Excela i zwraca:
    - lookup: numery -> (zlecenie, ilość palet, przewoźnik, uwagi)
    - all_nums: zbiór wszystkich numerów zleceń wyciągniętych z kolumny ZLECENIE
    Kolumna UWAGI jest opcjonalna (nagłówek 'UWAGI').
    """
    wb = load_workbook(file_like, data_only=True); ws = wb.active
    headers = {}
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=col).value
        if v is None: 
            continue
        headers[str(v).strip().lower()] = col

    z_col   = headers.get("zlecenie")
    ilo_col = headers.get("ilośc palet") or headers.get("ilosc palet") or headers.get("ilość palet")
    pr_col  = headers.get("przewoźnik") or headers.get("przewoznik")
    uw_col  = headers.get("uwagi")  # opcjonalne

    if not z_col or not ilo_col or not pr_col:
        raise ValueError("Excel musi mieć kolumny: ZLECENIE, ilość palet, przewoźnik (nagłówki w 1. wierszu).")

    lookup = {}
    all_nums = set()

    for row in range(2, ws.max_row + 1):
        z  = ws.cell(row=row, column=z_col).value
        il = ws.cell(row=row, column=ilo_col).value
        pr = ws.cell(row=row, column=pr_col).value
        uw = ws.cell(row=row, column=uw_col).value if uw_col else None

        z  = "" if z  is None else str(z).strip()
        il = "" if il is None else str(il).strip()
        pr = "" if pr is None else str(pr).strip()
        uw = "" if uw is None else str(uw).strip()

        parts = [p.strip() for p in re.split(r"[+;,/\s]+", z) if p.strip()]
        for p in parts:
            p2 = "".join(ch for ch in p if ch.isdigit())
            if p2.isdigit():
                all_nums.add(p2)
                lookup[p2] = (z, il, pr, uw)

    return lookup, all_nums

NBSP = "\u00A0"; NNBSP = "\u202F"; THINSP = "\u2009"
def normalize_digits(s: str) -> str:
    import re
    return re.sub(r"[\s\-{}{}{}]".format(NBSP, NNBSP, THINSP), "", s)

def extract_candidates(text: str):
    import re
    normal = re.findall(r"\b\d{4,8}\b", text)
    fancy = re.findall(r"(?<!\d)(?:\d[\s\u00A0\u202F\u2009\-]?){4,9}(?!\d)", text)
    fancy = [normalize_digits(s) for s in fancy]
    so = [normalize_digits(m.group(1)) for m in re.finditer(r"Sales\s*[\r\n ]*Order[\s:]*([0-9\s\u00A0\u202F\u2009\-]{4,12})", text, flags=re.I)]
    cands = normal + fancy + so
    cands = [c for c in cands if c.isdigit() and 4 <= len(c) <= 8]
    out, seen = [], set()
    for c in cands:
        if c not in seen:
            out.append(c); seen.add(c)
    return out

def adaptive_crop_extra(text: str):
    lines = [ln for ln in (text or "").splitlines() if ln.strip()]
    sparse = (len(lines) <= LOW_TEXT_LINES) or (len((text or "")) < SHORT_TEXT_CHARS)
    if sparse:
        return (EXTRA_CROP_LR*mm, EXTRA_CROP_LR*mm, EXTRA_CROP_T*mm, EXTRA_CROP_B*mm)
    return (0,0,0,0)

def make_overlay(width, height, header, footer, uwagi="", font_size=12, margin_mm=8):
    """
    Nadruk w prawym dolnym rogu:
    - nagłówek (ZLECENIE / ZLECENIA)
    - stopka (ilość palet / przewoźnik)
    - opcjonalnie linia z uwagami z Excela
    Układ pozostaje bardzo zbliżony do wersji v1.6 (dwie główne linie),
    trzecia linia z UWAGAMI jest dokładana niżej.
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height))
    try:
        c.setFont("Helvetica-Bold", font_size)
    except Exception:
        c.setFont("Helvetica", font_size)
    m = margin_mm * mm

    c.drawRightString(width - m, m + font_size + 1, header)
    if footer:
        c.drawRightString(width - m, m, footer)
    if uwagi:
        try:
            c.setFont("Helvetica", font_size - 1)
        except Exception:
            pass
        c.drawRightString(width - m, m - (font_size + 1), "uwagi: {}".format(strip_diacritics(uwagi)))

    c.save()
    return buf.getvalue()

def make_summary_page(width, height, missing_from_pdf, missing_from_excel):
    buf = io.BytesIO(); c = canvas.Canvas(buf, pagesize=(width, height))
    W, H = width, height
    try: c.setFont("Helvetica-Bold", 16)
    except Exception: c.setFont("Helvetica", 16)
    c.drawString(30, H-40, "RAPORT POROWNANIA DANYCH")

    y = H-80
    c.setFont("Helvetica-Bold", 12); c.drawString(30, y, "ZLECENIA Z EXCELA NIEZNALEZIONE W PDF:")
    y -= 20; c.setFont("Helvetica-Bold", 10); c.drawString(30, y, "ZLECENIE"); 
    y -= 12; c.setLineWidth(0.5); c.line(30, y, W-30, y); y -= 10
    c.setFont("Helvetica", 10)
    if not missing_from_pdf:
        c.drawString(30, y, "(brak)"); y -= 16
    else:
        for num in missing_from_pdf:
            c.drawString(30, y, str(num)); y -= 14
            if y < 80:
                c.showPage(); y = H-60; c.setFont("Helvetica", 10)

    if y < 140:
        c.showPage(); y = H-60

    c.setFont("Helvetica-Bold", 12); c.drawString(30, y, "ZLECENIA Z PDF-A NIEZNALEZIONE W EXCELU:")
    y -= 20; c.setFont("Helvetica-Bold", 10); c.drawString(30, y, "ZLECENIE")
    y -= 12; c.setLineWidth(0.5); c.line(30, y, W-30, y); y -= 10
    c.setFont("Helvetica", 10)
    if not missing_from_excel:
        c.drawString(30, y, "(brak)"); y -= 16
    else:
        for num in missing_from_excel:
            c.drawString(30, y, str(num)); y -= 14
            if y < 60:
                c.showPage(); y = H-60; c.setFont("Helvetica", 10)

    c.save(); return buf.getvalue()

def annotate_pdf_web(pdf_bytes, xlsx_bytes, max_per_sheet):
    lookup, excel_numbers = read_excel_lookup(io.BytesIO(xlsx_bytes))
    reader = PdfReader(io.BytesIO(pdf_bytes))
    groups, page_meta, page_text_cache = {}, {}, {}
    found_in_pdf = set()
    pdf_candidates_all = set()

    for i, _ in enumerate(reader.pages):
        page_text = extract_text(io.BytesIO(pdf_bytes), page_numbers=[i]) or ""
        page_text_cache[i] = page_text

        cands = extract_candidates(page_text)
        for c in cands:
            if c in excel_numbers:
                found_in_pdf.add(c)
            else:
                pdf_candidates_all.add(c)

        picked = next((n for n in cands if n in excel_numbers), None)
        mapped = lookup.get(picked) if picked else None
        if mapped:
            if len(mapped) == 4:
                z_full, il, pr, uw = mapped
            else:
                z_full, il, pr = mapped
                uw = ""
            key = z_full
            header = ("ZLECENIA (laczone): {}".format(strip_diacritics(z_full))
                      if "+" in z_full else "ZLECENIE: {}".format(strip_diacritics(z_full)))
            footer = "ilosc palet: {} | przewoznik: {}".format(strip_diacritics(il), strip_diacritics(pr))
        elif picked:
            key = picked
            header = "ZLECENIE: {}".format(picked)
            footer = "(brak danych w Excelu)"
            uw = ""
        else:
            key = "_NO_ORDER_{}".format(i+1)
            header = "(nie znaleziono numeru zlecenia na tej stronie)"
            footer = ""
            uw = ""

        groups.setdefault(key, []).append(i)
        page_meta[i] = (header, footer, uw)

    def key_sort(k: str):
        import re
        nums = [int(x) for x in re.findall(r"\d+", k)]
        return (min(nums) if nums else 10**9, k)

    ordered_keys = sorted(groups.keys(), key=key_sort)

    W, H = A4
    margin_x = SIDE_MARGIN_MM * mm
    top_margin = TOP_MARGIN_MM * mm
    bot_stamp = STAMP_BOTTOM_MM * mm
    gap = INTER_GAP_MM * mm
    avail_w = W - 2*margin_x
    avail_h = H - top_margin - bot_stamp
    base_crop_l = BASE_CROP_L*mm
    base_crop_r = BASE_CROP_R*mm
    base_crop_t = BASE_CROP_T*mm
    base_crop_b = BASE_CROP_B*mm

    writer = PdfWriter()
    writer.add_metadata({"/Producer": "Kersia PDF Stamper v1.7b (pypdf, uwagi + poprawiony raport 2)"})

    for gkey in ordered_keys:
        idxs = groups[gkey]
        for start in range(0, len(idxs), max_per_sheet):
            batch = idxs[start:start+max_per_sheet]
            items, total_h = [], 0.0
            for idx in batch:
                src = reader.pages[idx]
                sw = float(src.mediabox.width)
                sh = float(src.mediabox.height)
                ex_l, ex_r, ex_t, ex_b = adaptive_crop_extra(page_text_cache[idx])
                cl = base_crop_l + ex_l
                cr = base_crop_r + ex_r
                ct = base_crop_t + ex_t
                cb = base_crop_b + ex_b
                cw = max(10.0, sw - cl - cr)
                ch = max(10.0, sh - ct - cb)
                s  = avail_w / cw
                dh = s * ch
                items.append((idx, cl, cr, ct, cb, s, dh))
                total_h += dh
            total_h += gap * max(0, len(batch)-1)
            down = min(1.0, avail_h / total_h) if total_h > 0 else 1.0

            writer.add_blank_page(width=W, height=H)
            base_page = writer.pages[-1]
            y = H - top_margin
            for (idx, cl, cr, ct, cb, s, dh) in items:
                s *= down
                dh *= down
                x = margin_x - s * cl
                y2 = y - dh
                tmp = PageObject.create_blank_page(width=W, height=H)
                tmp.merge_page(reader.pages[idx])
                T = Transformation().translate(-cl, -cb).scale(s, s).translate(x, y2)
                tmp.add_transformation(T)
                base_page.merge_page(tmp)
                y = y2 - gap

            header, footer, uw = page_meta[batch[0]]
            ov = PdfReader(io.BytesIO(make_overlay(W, H, header, footer, uw)))
            base_page.merge_page(ov.pages[0])

    # --- RAPORT KOŃCOWY ---
    # ZLECENIA Z EXCELA NIEZNALEZIONE W PDF
    # i ZLECENIA Z PDF-A NIEZNALEZIONE W EXCELU
    # liczymy teraz na podstawie WSZYSTKICH kandydatów z PDF (extract_candidates),
    # żeby łapać także numery takie jak 55667, 55226 itp.
    pdf_all = found_in_pdf.union(pdf_candidates_all)
    if excel_numbers:
        excel_missing = sorted([n for n in excel_numbers if n not in pdf_all], key=lambda x: int(x))
    else:
        excel_missing = []
    pdf_only = sorted([n for n in pdf_all if n not in excel_numbers], key=lambda x: int(x))

    rep = PdfReader(io.BytesIO(make_summary_page(W, H, excel_missing, pdf_only)))
    writer.add_page(rep.pages[0])

    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    r2 = PdfReader(buf, strict=False)
    w2 = PdfWriter()
    for p in r2.pages:
        w2.add_page(p)
    out = io.BytesIO()
    w2.write(out)
    return out.getvalue()

# ---- UI ----
st.set_page_config(page_title="Kersia PDF Stamper v1.7b (Raport + UWAGI)", page_icon="🧰", layout="centered")
st.title("Kersia — PDF Stamper (raport braków + uwagi)")
excel_file = st.file_uploader("Plik Excel:", type=["xlsx", "xlsm", "xls"])
pdf_file = st.file_uploader("Plik PDF:", type=["pdf"])
max_per_sheet = st.slider("Maks. stron na kartkę", 1, 6, 3, 1)

if st.button("GENERUJ PDF", type="primary", disabled=not (excel_file and pdf_file)):
    try:
        data = annotate_pdf_web(pdf_file.read(), excel_file.read(), max_per_sheet)
        fname = "zlecenia_{}.pdf".format(datetime.now().strftime('%Y%m%d'))
        st.success("Gotowe! Pobierz poniżej.")
        st.download_button("Pobierz wynik", data=data, file_name=fname, mime="application/pdf")
    except Exception as e:
        st.error("Błąd: {}".format(repr(e)))