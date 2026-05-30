# -*- coding: utf-8 -*-
"""
Lectura del CSV "stock_trend_..." exportado del FMS Veridapt.

Contiene la serie de tiempo del nivel de varios tanques de Diesel (Main Tank,
Virtual Tank y la consolidacion "Newmont DIESEL"). Se usa para construir el
grafico "Daily Diesel Tank Log" (Figura 3) y para calcular estadisticas del
mes (minimo y promedio).

Identico al modulo del proyecto diesel_report — la estructura del CSV es la
misma; lo unico que cambia es la ventana temporal (mensual vs semanal), pero
eso lo decide el llamador.
"""
from __future__ import annotations

import csv
import datetime
import statistics

# Columnas conocidas en el header del CSV.
TANK_PATTERNS = {
    "Main Tank":    ["MAINTANK"],
    "Virtual Tank": ["VIRTUALTANK"],
    "Total":        ["NEWMONTDIESEL"],
}


def _num(value) -> float:
    if value is None or value == "":
        return 0.0
    text = str(value).strip().replace(",", "")
    try:
        return float(text)
    except ValueError:
        return 0.0


def _match_tank(column_name: str) -> str | None:
    up = str(column_name).upper().replace(" ", "").replace("-", "")
    for key, patterns in TANK_PATTERNS.items():
        for pat in patterns:
            p = pat.replace("-", "").replace(" ", "")
            if p in up:
                return key
    return None


def _parse_timestamp(text):
    if not text:
        return None
    raw = str(text).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S %z", "%Y-%m-%d %H:%M:%S",
                "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M"):
        try:
            dt = datetime.datetime.strptime(raw, fmt)
            return dt.replace(tzinfo=None)
        except ValueError:
            continue
    return None


def load_tank_trends(csv_path: str) -> tuple:
    """Devuelve (trends, safe_fill).
      trends    -> {tank_key: [(datetime, volumen), ...]}
      safe_fill -> {tank_key: capacidad}
    """
    with open(csv_path, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    if not rows:
        return {}, {}

    header = rows[0]
    col_tank = {}
    for idx, name in enumerate(header):
        if not name:
            continue
        key = _match_tank(name)
        if key:
            col_tank[idx] = key

    trends = {key: [] for key in col_tank.values()}
    safe_fill = {}

    for row in rows[1:]:
        if not row:
            continue
        label = str(row[0]).strip().lower()
        if "safe fill" in label:
            for idx, key in col_tank.items():
                if idx < len(row):
                    safe_fill[key] = _num(row[idx])
            continue
        ts = _parse_timestamp(row[0])
        if ts is None:
            continue
        for idx, key in col_tank.items():
            if idx < len(row):
                trends[key].append((ts, _num(row[idx])))

    for key in list(trends.keys()):
        trends[key] = sorted([p for p in trends[key] if p[1] is not None],
                             key=lambda x: x[0])

    return trends, safe_fill


def filter_by_month(points: list, month_first) -> list:
    """Recorta una serie temporal a los puntos cuyo dia cae dentro del mes
    cuyo 1er dia es `month_first`."""
    if not points:
        return []
    import calendar
    last_day = calendar.monthrange(month_first.year, month_first.month)[1]
    out = []
    for ts, val in points:
        d = ts.date() if isinstance(ts, datetime.datetime) else ts
        if d.year == month_first.year and d.month == month_first.month:
            out.append((ts, val))
        elif d == datetime.date(month_first.year, month_first.month,
                                 last_day) + datetime.timedelta(days=1):
            # Incluye el snapshot del 1er dia del mes siguiente a las 00:00 si
            # cae justo en la frontera (asi cerramos visualmente la curva).
            out.append((ts, val))
    return out


def tank_log_stats(points: list) -> dict:
    """Calcula min, max y promedio de la serie."""
    if not points:
        return {"min_value": 0.0, "min_date": "", "average": 0.0}
    values = [p[1] for p in points if p[1] > 0]
    if not values:
        return {"min_value": 0.0, "min_date": "", "average": 0.0}
    min_idx = min(range(len(points)), key=lambda i: points[i][1])
    min_dt = points[min_idx][0]
    min_value = points[min_idx][1]
    average = statistics.mean(values)
    months_en = ["January", "February", "March", "April", "May", "June",
                 "July", "August", "September", "October", "November",
                 "December"]
    suffix = "th" if 10 <= min_dt.day % 100 <= 20 else \
        {1: "st", 2: "nd", 3: "rd"}.get(min_dt.day % 10, "th")
    min_date_text = "%d%s %s" % (min_dt.day, suffix, months_en[min_dt.month - 1])
    return {
        "min_value": round(min_value, 0),
        "min_date": min_date_text,
        "average": round(average, 0),
    }
