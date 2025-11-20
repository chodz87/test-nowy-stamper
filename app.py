import io
import os
import re
from datetime import datetime, date
from typing import Any, Dict, List, Tuple, Set

import streamlit as st
from openpyxl import load_workbook
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2._page import PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="Kersia PDF Stamper", page_icon="🧰", layout="centered")
st.title("Kersia — PDF Stamper (wersja z raportem i uwagami)")


# ----------------------- POMOCNICZE -----------------------

def _coerce_int(value: Any) -> int:
    """Bezpieczne rzutowanie na int (dla ilości palet)."""
    if value is None:
        return 0
    if isinstance(value, (datetime, date)):
        return 0
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        if value != value:  # NaN
            return 0
        return int(round(value))
    s = str(value).strip()
    if not s:
        return 0
    m = re.search(r"-?\d+", s.replace(",", "."))
    return int(m.group(0)) if m else 0


def _to_str(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return str(v)


def _parse_excel(excel_bytes: bytes) -> Tuple[List[Dict[str, Any]], Set[str]]:
    """
    Parsuje Excela i zwraca:
    - listę wierszy: dict z kluczami: zlecenie, ilosc, przewoznik, uwagi
    - zbiór wszystkich numerów zleceń (string)
    """
    wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active

    # Dopasowanie nagłówków „elastycznie”
    header_row = {}
    for c in range(1, ws.max_column + 1):
        name = str(ws.cell(1, c).value or "").strip().lower()
        if name:
            header_row[name] = c

    # Które kolumny nas interesują
    col_z = header_row.get("zlecenie", header_row.get("nr zlecenia", 1))
    col_i = (
        header_row.get(
            "ilość palet",
            header_row.get(
                "ilosc palet",
                header_row.get("ilosc", header_row.get("ilość", 2)),
            ),
        )
        or 2
    )
    col_p = (
        header_row.get(
            "przewoźnik",
            header_row.get(
                "przewoznik", header_row.get("przewoź", header_row.get("przewoz", 3))
            ),
        )
        or 3
    )
    col_u = header_row.get("uwagi")  # może nie istnieć

    rows: List[Dict[str, Any]] = []
    zlecenia_set: Set[str] = set()

    for r in range(2, ws.max_row + 1):
        z = ws.cell(r, col_z).value if col_z else None
        i = ws.cell(r, col_i).value if col_i else None
        p = ws.cell(r, col_p).value if col_p else None
        u = ws.cell(r, col_u).value if col_u else None

        if z is None and i is None and p is None and u is None:
            continue

        z_str = _to_str(z).strip()
        i_int = _coerce_int(i)
        p_str = _to_str(p).strip()
        u_str = _to_str(u).strip()

        if z_str:
            zlecenia_set.add(z_str)

        rows.append(
            {
                "zlecenie": z_str,
                "ilosc": i_int,
                "przewoznik": p_str,
                "uwagi": u_str,
            }
        )

    return rows, zlecenia_set


def _register_fonts() -> str:
    """Rejestruje font z polskimi znakami (jeśli dostępny)."""
    try_paths = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/local/share/fonts/DejaVuSans.ttf",
        os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf"),
    ]
    for p in try_paths:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont("DejaVuSans", p))
            return "DejaVuSans"
    return "Helvetica"


def _make_stamp_page(
    zlecenie: str,
    ilosc: int,
    przewoznik: str,
    uwagi: str,
    width: float,
    height: float,
) -> bytes:
    """Tworzy pojedynczą stronę z nadrukiem (overlay)."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height), bottomup=True)
    font_name = _register_fonts()
    c.setAuthor("Kersia PDF Stamper")
    c.setTitle(f"Zlecenie {zlecenie}")
    c.setCreator("Kersia PDF Stamper (wersja z raportem)")

    c.setFont(font_name, 14)
    margin = 15 * mm
    x = margin
    y = height - margin

    c.drawString(x, y, f"ZLECENIE: {zlecenie}")
    y -= 8 * mm
    c.drawString(x, y, f"ILOŚĆ PALET: {ilosc}")
    y -= 8 * mm
    c.drawString(x, y, f"PRZEWOŹNIK: {przewoznik}")
    y -= 8 * mm
    if uwagi:
        c.drawString(x, y, f"UWAGI: {uwagi}")
        y -= 8 * mm

    c.showPage()
    c.save()
    return buf.getvalue()


def _extract_sales_orders_from_pdf(pdf_bytes: bytes) -> Set[str]:
    """
    Skanuje tekst z PDF i wyciąga numery „Sales order ...”.
    Zwraca zbiór stringów.
    """
    reader = PdfReader(io.BytesIO(pdf_bytes), strict=False)
    sales_orders: Set[str] = set()

    pattern = re.compile(
        r"sales\s*order[^\d]*(\d+)", re.IGNORECASE
    )  # „Sales order 123456”, „Sales order: 123456” itd.

    for page in reader.pages:
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        for match in pattern.finditer(text):
            num = match.group(1).strip()
            if num:
                sales_orders.add(num)

    return sales_orders


def _make_report_page(
    not_in_excel: List[str], width: float = 595.27, height: float = 841.89
) -> bytes:
    """
    Tworzy stronę raportową na końcu:
    „ZLECENIA Z PDF-A NIEZNALEZIONE W EXCELU”
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(width, height), bottomup=True)
    font_name = _register_fonts()
    c.setAuthor("Kersia PDF Stamper")
    c.setTitle("Raport zleceń")
    c.setCreator("Kersia PDF Stamper (raport)")

    c.setFont(font_name, 16)
    margin = 20 * mm
    x = margin
    y = height - margin
    c.drawString(x, y, "ZLECENIA Z PDF-A NIEZNALEZIONE W EXCELU")
    y -= 12 * mm

    c.setFont(font_name, 12)

    if not not_in_excel:
        c.drawString(x, y, "Wszystkie numery Sales order z PDF występują w Excelu.")
    else:
        for num in not_in_excel:
            if y < margin:
                c.showPage()
                y = height - margin
                c.setFont(font_name, 12)
            c.drawString(x, y, f"- {num}")
            y -= 7 * mm

    c.showPage()
    c.save()
    return buf.getvalue()


# ----------------------- GŁÓWNA FUNKCJA -----------------------

def annotate_pdf(pdf_bytes: bytes, excel_bytes: bytes, max_per_sheet: int = 3) -> bytes:
    # 1. Parsujemy Excela
    rows, excel_zlecenia = _parse_excel(excel_bytes)
    if not rows:
        raise ValueError(
            "Nie znaleziono danych w Excelu. Upewnij się, że masz kolumny: ZLECENIE, ILOŚĆ PALET, PRZEWOŹNIK (i opcjonalnie UWAGI)."
        )

    # 2. Zbieramy numery Sales order z PDF
    pdf_sales_orders = _extract_sales_orders_from_pdf(pdf_bytes)

    # 3. Tworzymy nowy PDF z nadrukami
    reader = PdfReader(io.BytesIO(pdf_bytes), strict=False)
    writer = PdfWriter()
    writer.add_metadata(
        {
            "/Producer": "Kersia PDF Stamper (PyPDF2 + ReportLab)",
            "/Creator": "Kersia PDF Stamper (wersja z raportem)",
            "/Title": "Zlecenia",
        }
    )

    data_idx = 0
    pages = list(reader.pages)

    for page in pages:
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)

        # Na tej stronie nakładamy max_per_sheet wpisów z Excela
        for _ in range(max_per_sheet):
            if data_idx >= len(rows):
                break
            row = rows[data_idx]
            data_idx += 1

            overlay_bytes = _make_stamp_page(
                row["zlecenie"], row["ilosc"], row["przewoznik"], row["uwagi"], w, h
            )
            overlay_reader = PdfReader(io.BytesIO(overlay_bytes), strict=False)
            overlay_page = overlay_reader.pages[0]
            page.merge_page(overlay_page)

        writer.add_page(page)

    # Jeżeli zostały wiersze z Excela, dorabiamy kolejne strony na bazie ostatniego rozmiaru
    while data_idx < len(rows):
        if pages:
            w = float(pages[-1].mediabox.width)
            h = float(pages[-1].mediabox.height)
        else:
            # domyślnie A4 w punktach
            w, h = (595.27, 841.89)

        blank = PageObject.create_blank_page(width=w, height=h)

        for _ in range(max_per_sheet):
            if data_idx >= len(rows):
                break
            row = rows[data_idx]
            data_idx += 1
            overlay_bytes = _make_stamp_page(
                row["zlecenie"], row["ilosc"], row["przewoznik"], row["uwagi"], w, h
            )
            overlay_reader = PdfReader(io.BytesIO(overlay_bytes), strict=False)
            overlay_page = overlay_reader.pages[0]
            blank.merge_page(overlay_page)

        writer.add_page(blank)

    # 4. Raport: zlecenia, które są w PDF (Sales order), ale NIE MA ich w Excelu
    not_in_excel = sorted(pdf_sales_orders - excel_zlecenia, key=str)
    report_bytes = _make_report_page(not_in_excel)
    report_reader = PdfReader(io.BytesIO(report_bytes), strict=False)
    for p in report_reader.pages:
        writer.add_page(p)

    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


# ----------------------- UI -----------------------

st.markdown("Ta wersja dodaje stronę raportu oraz pole **UWAGI** z Excela na każdej etykiecie.")

excel_file = st.file_uploader(
    "Plik Excel (ZLECENIE, ilość palet, przewoźnik, opcjonalnie UWAGI):",
    type=["xlsx", "xlsm", "xls"],
)
pdf_file = st.file_uploader(
    "Plik PDF (szablon/strony z Sales order):",
    type=["pdf"],
)
max_per_sheet = st.slider(
    "Maks. wpisów na stronę PDF", min_value=1, max_value=6, value=3, step=1
)

if st.button("GENERUJ PDF", type="primary", disabled=not (excel_file and pdf_file)):
    try:
        result = annotate_pdf(pdf_file.read(), excel_file.read(), max_per_sheet)
        fname = "zlecenia_{}.pdf".format(datetime.now().strftime("%Y%m%d"))
        st.success("Gotowe! Poniżej przycisk pobierania.")
        st.download_button(
            "Pobierz wynik", data=result, file_name=fname, mime="application/pdf"
        )
    except Exception as e:
        st.error(f"Błąd: {e}")