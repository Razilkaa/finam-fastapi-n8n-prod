"""Generate quotes_all Word document by filling a .docx template with multiple tables (including textboxes)."""

from __future__ import annotations

import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
from typing import Any, Optional

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

GREEN = "00B050"
RED = "FF0000"
BLACK = "000000"


@dataclass(frozen=True)
class QuoteAll:
    symbol: str
    old_price_raw: Optional[str]
    new_price_raw: Optional[str]
    change_value_raw: Optional[str]
    change_unit_raw: Optional[str]
    report_date_raw: Optional[str]


def _norm_text(value: str) -> str:
    return (
        value.replace("\u00a0", " ")
        .replace("\u2212", "-")  # minus sign
        .strip()
    )


def _norm_symbol(value: str) -> str:
    value = " ".join(_norm_text(value).split()).lower()
    if value == "msci ac asia pacific":
        return "msci asia pacific"
    return value


def _to_decimal(value: str | None) -> Decimal | None:
    if value is None:
        return None
    value = _norm_text(str(value))
    if not value:
        return None
    try:
        return Decimal(value.replace(",", "."))
    except InvalidOperation:
        return None


def _parse_report_date(value: Any) -> Optional[date]:
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value.date() if isinstance(value, datetime) else value

    s = str(value).strip()
    if not s:
        return None

    if s.endswith("Z"):
        try:
            return datetime.fromisoformat(s.replace("Z", "+00:00")).date()
        except ValueError:
            pass

    try:
        return datetime.fromisoformat(s).date()
    except ValueError:
        pass

    try:
        return datetime.strptime(s, "%d.%m.%Y").date()
    except ValueError:
        return None


def parse_quotes_all(payload: list[dict]) -> tuple[list[QuoteAll], Optional[date]]:
    quotes: list[QuoteAll] = []
    report_dt: Optional[date] = None

    for item in payload:
        if not isinstance(item, dict):
            continue

        report_date_raw = item.get("report_date")
        if report_dt is None:
            report_dt = _parse_report_date(report_date_raw)

        symbol = str(item.get("symbol", "")).strip()
        if not symbol:
            continue

        old_price_raw = item.get("old_price")
        new_price_raw = item.get("new_price")
        change_value_raw = item.get("change_value", item.get("pct_change"))
        change_unit_raw = item.get("change_unit", "%")

        quotes.append(
            QuoteAll(
                symbol=symbol,
                old_price_raw=None if old_price_raw is None else str(old_price_raw).strip(),
                new_price_raw=None if new_price_raw is None else str(new_price_raw).strip(),
                change_value_raw=None
                if change_value_raw is None
                else str(change_value_raw).strip(),
                change_unit_raw=None
                if change_unit_raw is None
                else str(change_unit_raw).strip(),
                report_date_raw=None
                if report_date_raw is None
                else str(report_date_raw).strip(),
            )
        )

    return quotes, report_dt


def _group_thousands(int_part: str) -> str:
    s = int_part
    if len(s) <= 3:
        return s
    out: list[str] = []
    while s:
        out.append(s[-3:])
        s = s[:-3]
    return " ".join(reversed(out))


def format_number(value: str | None) -> str | None:
    dec = _to_decimal(value)
    if dec is None:
        return None

    raw = _norm_text(str(value))
    if "." in raw:
        frac_len = len(raw.split(".", 1)[1])
    elif "," in raw:
        frac_len = len(raw.split(",", 1)[1])
    else:
        frac_len = 0

    sign = "-" if dec < 0 else ""
    dec = abs(dec)

    quant = Decimal("1") if frac_len == 0 else Decimal("1." + ("0" * frac_len))
    dec = dec.quantize(quant)

    s = f"{dec:f}"
    if "." in s:
        int_part, frac_part = s.split(".", 1)
    else:
        int_part, frac_part = s, ""

    int_part = _group_thousands(int_part)
    if frac_len:
        return f"{sign}{int_part},{frac_part}"
    return f"{sign}{int_part}"


def format_change(
    change_value: str | None,
    change_unit: str | None,
    *,
    fallback_from_old_new: tuple[str | None, str | None] | None = None,
) -> tuple[str | None, str | None]:
    unit = (change_unit or "").strip().lower()

    dec = _to_decimal(change_value)
    if dec is None and fallback_from_old_new is not None:
        old_s, new_s = fallback_from_old_new
        old = _to_decimal(old_s)
        new = _to_decimal(new_s)
        if old is not None and new is not None:
            if unit in {"%", "percent"} and old != 0:
                dec = (new - old) / old * 100
            elif unit in {"bp", "b.p.", "б.п.", "бп"}:
                dec = (new - old) * 100

    if dec is None:
        return None, None

    sign = "+" if dec > 0 else "-" if dec < 0 else ""
    abs_dec = abs(dec)

    if unit in {"bp", "b.p.", "б.п.", "бп"}:
        abs_dec = abs_dec.quantize(Decimal("1"))
        num = format_number(str(abs_dec))
        color = GREEN if sign == "+" else RED if sign == "-" else BLACK
        return f"{sign}{num} б.п.", color

    if unit in {"%", "percent", "pct"} or unit == "":
        abs_dec = abs_dec.quantize(Decimal("1.00"))
        num = format_number(str(abs_dec))
        color = GREEN if sign == "+" else RED if sign == "-" else BLACK
        return f"{sign}{num}%", color

    num = format_number(str(abs_dec))
    color = GREEN if sign == "+" else RED if sign == "-" else BLACK
    return f"{sign}{num} {change_unit}".strip(), color


def get_text(node: etree._Element) -> str:
    return _norm_text("".join(node.itertext()))


def set_cell_text(tc: etree._Element, text: str) -> None:
    ts = tc.xpath(".//w:t", namespaces=NS)
    if not ts:
        p = tc.find(".//w:p", namespaces=NS)
        if p is None:
            p = etree.SubElement(tc, f"{{{W_NS}}}p")
        r = p.find("./w:r", namespaces=NS)
        if r is None:
            r = etree.SubElement(p, f"{{{W_NS}}}r")
        t = etree.SubElement(r, f"{{{W_NS}}}t")
        t.text = text
        return

    ts[0].text = text
    for t in ts[1:]:
        t.text = ""


def set_color(container: etree._Element, color: str | None) -> None:
    if not color:
        return

    for p in container.xpath(".//w:p", namespaces=NS):
        ppr = p.find("./w:pPr", namespaces=NS)
        if ppr is None:
            ppr = etree.SubElement(p, f"{{{W_NS}}}pPr")
        prpr = ppr.find("./w:rPr", namespaces=NS)
        if prpr is None:
            prpr = etree.SubElement(ppr, f"{{{W_NS}}}rPr")
        c = prpr.find("./w:color", namespaces=NS)
        if c is None:
            c = etree.SubElement(prpr, f"{{{W_NS}}}color")
        c.set(f"{{{W_NS}}}val", color)

    for r in container.xpath(".//w:r", namespaces=NS):
        rpr = r.find("./w:rPr", namespaces=NS)
        if rpr is None:
            rpr = etree.SubElement(r, f"{{{W_NS}}}rPr")
        c = rpr.find("./w:color", namespaces=NS)
        if c is None:
            c = etree.SubElement(rpr, f"{{{W_NS}}}color")
        c.set(f"{{{W_NS}}}val", color)


def fill_template_all(*, template_path: Path, quotes: list[QuoteAll]) -> tuple[BytesIO, int]:
    quotes_by_symbol = { _norm_symbol(q.symbol): q for q in quotes }

    with zipfile.ZipFile(template_path, "r") as zin:
        doc_xml = zin.read("word/document.xml")
        root = etree.fromstring(doc_xml)

        updated_rows = 0
        for tr in root.xpath(".//w:tr", namespaces=NS):
            tcs = tr.xpath("./w:tc", namespaces=NS)
            if len(tcs) < 3:
                continue

            symbol = get_text(tcs[0])
            if not symbol:
                continue

            quote = quotes_by_symbol.get(_norm_symbol(symbol))
            if quote is None:
                continue

            changed = False

            new_price = format_number(quote.new_price_raw)
            if new_price is not None:
                set_cell_text(tcs[1], new_price)
                changed = True

            change_text, change_color = format_change(
                quote.change_value_raw,
                quote.change_unit_raw,
                fallback_from_old_new=(quote.old_price_raw, quote.new_price_raw),
            )
            if change_text is None:
                unit = (quote.change_unit_raw or "").strip().lower()
                if unit in {"bp", "b.p.", "б.п.", "бп"}:
                    change_text = "0 б.п."
                else:
                    change_text = "0,00%"
                change_color = BLACK

            set_cell_text(tcs[2], change_text)
            set_color(tcs[2], change_color)
            changed = True

            if changed:
                updated_rows += 1

        new_doc_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8")

        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/document.xml":
                    data = new_doc_xml
                zout.writestr(info, data)

        buffer.seek(0)
        return buffer, updated_rows


def get_quotes_all_filename(report_dt: Optional[date]) -> str:
    if report_dt is not None:
        return f"Daily_quotes_all_{report_dt.strftime('%d.%m.%Y')}.docx"
    return "Daily_quotes_all.docx"
