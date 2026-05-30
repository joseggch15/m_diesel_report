# -*- coding: utf-8 -*-
"""
Extraccion de datos desde los PDFs de reconciliacion detallada de Veridapt
para el circuito Diesel (version mensual).

Estos PDFs se generan por sitio para un periodo MENSUAL:
  - 'Detailed Reconciliation Monthly'      -> consolidado LFO Main+Virtual.
       Tiene dos sub-tanques (LFO - Main Tank y LFO - Virtual Tank);
       el dispensing relevante para el reporte vive en LFO - Virtual Tank.
  - 'Detailed Reconciliation 0846/47/48'   -> service truck (mes completo).

Datos extraidos:
  site, opening, closing, inflow, outflow, to_equipment, other_dispenses,
  transfers_out, period_start, period_end, is_truck.

Practicamente identico a diesel_report.pdf_import. Diferencia:
  - acepta 'Detailed Reconciliation Monthly' como titulo para LFO.
"""
from __future__ import annotations

import os
import re

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


# Identificacion del site.
_TITLE_PATTERNS = [
    ("LFO",     re.compile(r"monthly", re.IGNORECASE)),
    ("LFO",     re.compile(r"weekly\s+diesel", re.IGNORECASE)),
    ("LFO",     re.compile(r"\bLFO\b", re.IGNORECASE)),
    ("TFL0846", re.compile(r"\b0846\b")),
    ("TFL0847", re.compile(r"\b0847\b")),
    ("TFL0848", re.compile(r"\b0848\b")),
]

_TANK_SECTION_HEADERS = {
    "LFO":     ["LFO - Virtual Tank"],
    "TFL0846": ["TFL0846"],
    "TFL0847": ["TFL0847"],
    "TFL0848": ["TFL0848"],
}

_SECTION_BOUNDARIES = [
    "LFO - Virtual Tank",
    "LFO - Main Tank",
    "TFL0846",
    "TFL0847",
    "TFL0848",
    "Daily Reconciliation",
    "Report Configuration",
]


def is_available() -> bool:
    return pdfplumber is not None


def _parse_number(text: str) -> float:
    text = str(text).strip()
    text = re.sub(r'\s*[L%]$', '', text)
    text = text.replace(',', '')
    try:
        return float(text)
    except ValueError:
        return 0.0


_RE_TITLE = re.compile(r'Detailed Reconciliation\s+(.+)')

_RE_PERIOD = re.compile(
    r'(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}\s*-\s*'
    r'(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}'
)

# Product Summary
_RE_PRODUCT_SUMMARY = re.compile(
    r'-?[\d.]+%\s+'
    r'Diesel\s+'
    r'(-?[\d,]+\.?\d*)\s*L\s+'
    r'(-?[\d,]+\.?\d*)\s*L\s+'
    r'(-?[\d,]+\.?\d*)\s*L\s+'
    r'(-?[\d,]+\.?\d*)\s*L'
)

_RE_TO_EQUIPMENT    = re.compile(r'To Equipment\s+(-?[\d,]+\.?\d*)\s*L')
_RE_OTHER_DISPENSES = re.compile(r'Other Dispenses\s+(-?[\d,]+\.?\d*)\s*L')
_RE_TRANSFERS_OUT   = re.compile(r'Transfers out\s+(-?[\d,]+\.?\d*)\s*L')


def _detect_site(title_line: str, full_text: str) -> str | None:
    for site, pattern in _TITLE_PATTERNS:
        if pattern.search(title_line):
            return site
    for site, pattern in _TITLE_PATTERNS:
        if pattern.search(full_text):
            return site
    return None


def _isolate_tank_section(full_text: str, site_key: str) -> str:
    headers = _TANK_SECTION_HEADERS.get(site_key, [])
    if not headers:
        return full_text

    start = -1
    for h in headers:
        pos = 0
        while True:
            idx = full_text.find(h, pos)
            if idx < 0:
                break
            tail = full_text[idx + len(h): idx + len(h) + 80]
            if "\nOpening Stock" in tail or "\nClosing Stock" in tail:
                start = full_text.rfind("\n", 0, idx) + 1
                break
            pos = idx + len(h)
        if start >= 0:
            break
    if start < 0:
        return full_text

    others = [h for h in _SECTION_BOUNDARIES if h not in headers]
    end = len(full_text)
    for h in others:
        idx = full_text.find(h, start + 1)
        if idx >= 0 and idx < end:
            end = idx
    return full_text[start:end]


def parse_veridapt_pdf(path: str) -> dict:
    if pdfplumber is None:
        raise ImportError(
            "Se requiere la libreria 'pdfplumber' para leer PDFs.\n"
            "Instale con:  pip install pdfplumber")

    with pdfplumber.open(path) as pdf:
        pages_text = [page.extract_text() or "" for page in pdf.pages]

    page1 = pages_text[0] if pages_text else ""
    full_text = "\n".join(pages_text)

    result = {
        "site": None,
        "opening": 0.0,
        "closing": 0.0,
        "inflow": 0.0,
        "outflow": 0.0,
        "to_equipment": 0.0,
        "other_dispenses": 0.0,
        "transfers_out": 0.0,
        "period_start": "",
        "period_end": "",
        "is_truck": False,
    }

    mp = _RE_PERIOD.search(page1)
    if mp:
        result["period_start"] = mp.group(1)
        result["period_end"] = mp.group(2)

    title_line = ""
    for line in page1.split("\n"):
        if "Detailed Reconciliation" in line and "00:00" not in line:
            title_line = line
            mt = _RE_TITLE.match(line.strip())
            if mt and mt.group(1).strip().lower() != "detailed reconciliation":
                break
    result["site"] = _detect_site(title_line, full_text)
    result["is_truck"] = (result["site"] in ("TFL0846", "TFL0847", "TFL0848"))

    section = _isolate_tank_section(full_text, result["site"] or "")

    ms = _RE_PRODUCT_SUMMARY.search(page1)
    if ms:
        result["opening"] = _parse_number(ms.group(1))
        result["closing"] = _parse_number(ms.group(2))
        result["inflow"] = _parse_number(ms.group(3))
        result["outflow"] = _parse_number(ms.group(4))

    for val in _RE_TO_EQUIPMENT.findall(section):
        result["to_equipment"] += _parse_number(val)
    for val in _RE_OTHER_DISPENSES.findall(section):
        result["other_dispenses"] += _parse_number(val)
    for val in _RE_TRANSFERS_OUT.findall(section):
        result["transfers_out"] += _parse_number(val)

    return result


def parse_multiple_pdfs(paths: list) -> tuple:
    loaded = []
    skipped = []
    for path in paths:
        fname = os.path.basename(path)
        try:
            info = parse_veridapt_pdf(path)
        except Exception as exc:
            skipped.append((fname, str(exc)))
            continue
        if info["site"] is None:
            skipped.append((fname, "No se pudo detectar el site (LFO/TFL...)."))
            continue
        loaded.append((info["site"], info, fname))
    return loaded, skipped
