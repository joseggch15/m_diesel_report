# -*- coding: utf-8 -*-
"""
Generacion de los 7 graficos del reporte mensual de Diesel con matplotlib.

  - Figura 1: barras Daily Inlet Deliveries vs Delivery Tickets (~30 dias).
  - Figura 2: dona "Fuel Outflow Distribution" (Equipment / Other / Transfers).
  - Figura 3: Daily Diesel Tank Log (linea de tiempo del volumen).
  - Figura 4: Reconciliation trend mensual (Recon %).
  - Figura 5: Delivery vs Reconciliation (Delivery % + Recon %, 2 series).
  - Figura 6: Monthly Breakdown of Fuel Consumption (bars semanales 3 series).
  - Figura 7: Historical Transaction Volume by Type (bars 3 series por mes).
"""
from __future__ import annotations

import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt

EXCEL_BLUE = "#4472C4"
EXCEL_ORANGE = "#ED7D31"
EXCEL_GRAY = "#A5A5A5"
EXCEL_YELLOW = "#FFC000"
EXCEL_DARK_BLUE = "#1F3864"


def _num(value) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def _save(fig, path: str) -> str:
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def _fmt_thousands(value) -> str:
    return "{:,.2f}".format(_num(value))


# --------------------------------------------------------------------------
# Figura 1: Detailed Inlet Deliveries vs Delivery Tickets (daily bars)
# --------------------------------------------------------------------------

def delivery_daily_bar_chart(rows: list, path: str) -> str:
    fig, ax = plt.subplots(figsize=(11.0, 4.0))
    ax.set_title("DETAILED INLET DELIVERIES VS DELIVERY TICKETS COMPARISON",
                 fontsize=11, fontweight="bold")
    if not rows:
        ax.text(0.5, 0.5, "No confirmed deliveries in this period",
                ha="center", va="center", fontsize=10, color="#555555")
        ax.set_axis_off()
        return _save(fig, path)
    labels = [str(r.get("date", "")) for r in rows]
    tickets = [_num(r.get("tickets")) for r in rows]
    inlet = [_num(r.get("inlet")) for r in rows]
    # Variance % = (inlet - tickets) / tickets * 100
    variance_pct = []
    for v, d in zip(inlet, tickets):
        variance_pct.append(((v - d) / d * 100.0) if d else 0.0)

    x = list(range(len(rows)))
    width = 0.38
    bars_tick = ax.bar([i - width / 2 for i in x], tickets, width,
                       label="Delivery Tickets", color=EXCEL_BLUE)
    bars_inl = ax.bar([i + width / 2 for i in x], inlet, width,
                      label="Inlet Deliveries", color=EXCEL_ORANGE)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=70, fontsize=7, ha="right")
    ax.set_ylabel("Volume (L)", fontsize=9)
    ax.grid(axis="y", alpha=0.3)
    ax.spines["top"].set_visible(False)

    ax2 = ax.twinx()
    line, = ax2.plot(x, variance_pct, color=EXCEL_GRAY, linewidth=1.5,
                     marker="o", markersize=3, label="Variance %")
    ax2.set_ylabel("Variance %", fontsize=9, color="#555555")
    ax2.yaxis.set_major_formatter(
        plt.FuncFormatter(lambda v, _: "{:.2f}%".format(v).replace(".", ",")))
    ax2.spines["top"].set_visible(False)

    handles = [bars_tick, bars_inl, line]
    legend_labels = ["Delivery Tickets", "Inlet Deliveries", "Variance %"]
    ax.legend(handles, legend_labels, loc="upper center",
              bbox_to_anchor=(0.5, -0.22), ncol=3, frameon=False, fontsize=9)
    return _save(fig, path)


# --------------------------------------------------------------------------
# Figura 2: Donut Outflow Distribution
# --------------------------------------------------------------------------

def outflow_donut_chart(equipment, other, transfers, path: str,
                         month_label: str = "") -> str:
    fig, ax = plt.subplots(figsize=(5.2, 3.6))
    title = "Fuel Outflow Distribution"
    if month_label:
        title = "%s - %s" % (title, month_label)
    ax.set_title(title, fontsize=11, fontweight="bold", pad=24)
    values = [_num(equipment), _num(other), _num(transfers)]
    labels = ["Dispensing to Equipment", "Other Dispenses", "Transfers"]
    colors = [EXCEL_BLUE, EXCEL_ORANGE, EXCEL_GRAY]
    if sum(values) <= 0:
        ax.text(0.5, 0.5, "No outflow in this period",
                ha="center", va="center", fontsize=10, color="#555555")
        ax.set_axis_off()
        return _save(fig, path)
    wedges, _texts, autotexts = ax.pie(
        values, colors=colors, startangle=90, counterclock=False,
        autopct=lambda p: ("%.0f%%" % round(p)) if p > 0 else "",
        pctdistance=0.78, wedgeprops=dict(width=0.42, edgecolor="white"))
    for at in autotexts:
        at.set_color("white")
        at.set_fontsize(9)
    ax.legend(wedges, labels, loc="upper center",
              bbox_to_anchor=(0.5, -0.02), ncol=1, frameon=False, fontsize=8)
    ax.set(aspect="equal")
    return _save(fig, path)


# --------------------------------------------------------------------------
# Figura 3: Daily Tank Log
# --------------------------------------------------------------------------

def daily_tank_log_chart(points: list, title: str, path: str,
                          safe_fill=None) -> str:
    fig, ax = plt.subplots(figsize=(11.0, 3.2))
    ax.set_title(title, loc="left", fontsize=11, color="#7F7F7F",
                 pad=10, fontweight="normal")
    if not points:
        ax.text(0.5, 0.5, "No tank-level data available",
                ha="center", va="center", fontsize=10, color="#555555")
        ax.set_axis_off()
        return _save(fig, path)
    xs = [p[0] for p in points]
    ys = [_num(p[1]) for p in points]
    ax.plot(xs, ys, color=EXCEL_BLUE, linewidth=1.4)
    if safe_fill:
        sf = _num(safe_fill)
        ax.axhline(sf, color="#C00000", linewidth=1.4)
        ax.text(xs[0], sf, " Safe Fill Level (%s L)" % _fmt_thousands(sf),
                color="#C00000", fontsize=9, fontweight="bold",
                va="bottom", ha="left")
    ax.set_ylabel("Volume", color="#7F7F7F", fontsize=10)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d/%m"))
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=2))
    fig.autofmt_xdate(rotation=45)
    ax.grid(axis="y", alpha=0.5, color="#D9D9D9", linewidth=0.7)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    for spine in ("left", "bottom"):
        ax.spines[spine].set_color("#D9D9D9")
    ax.tick_params(colors="#7F7F7F", labelsize=8)
    return _save(fig, path)


# --------------------------------------------------------------------------
# Figura 4: Monthly Recon trend (Recon %)
# --------------------------------------------------------------------------

def monthly_recon_trend_chart(points: list, title: str, path: str) -> str:
    """points = [(date, deliv_pct, recon_pct)]; usa columna recon."""
    fig, ax = plt.subplots(figsize=(9.0, 3.6))
    ax.set_title(title, fontsize=11, fontweight="bold")
    if not points:
        ax.text(0.5, 0.5, "No reconciliation history available",
                ha="center", va="center", fontsize=10, color="#555555")
        ax.set_axis_off()
        return _save(fig, path)
    labels = [_month_label(p[0]) for p in points]
    ys = [_num(p[2]) for p in points]
    x = list(range(len(points)))
    ax.plot(x, ys, color=EXCEL_ORANGE, linewidth=1.8, marker="o",
            markersize=4, label="Recon %")
    for i, v in zip(x, ys):
        ax.annotate("%.2f" % v, (i, v), textcoords="offset points",
                    xytext=(0, 7), ha="center", fontsize=7)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7)
    ax.set_ylabel("Variance (%)", fontsize=9)
    ax.grid(axis="y", alpha=0.3)
    ax.legend(loc="upper left", fontsize=8, frameon=False)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    return _save(fig, path)


# --------------------------------------------------------------------------
# Figura 5: Delivery % vs Recon % (2 series)
# --------------------------------------------------------------------------

def delivery_vs_recon_chart(points: list, title: str, path: str) -> str:
    fig, ax = plt.subplots(figsize=(9.0, 3.6))
    ax.set_title(title, fontsize=11, fontweight="bold")
    if not points:
        ax.text(0.5, 0.5, "No reconciliation history available",
                ha="center", va="center", fontsize=10, color="#555555")
        ax.set_axis_off()
        return _save(fig, path)
    labels = [_month_label(p[0]) for p in points]
    deliv = [_num(p[1]) for p in points]
    recon = [_num(p[2]) for p in points]
    x = list(range(len(points)))
    ax.plot(x, deliv, color=EXCEL_BLUE, linewidth=1.6, marker="o",
            markersize=4, label="Delivery %")
    ax.plot(x, recon, color=EXCEL_ORANGE, linewidth=1.6, marker="o",
            markersize=4, label="Recon %")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7)
    ax.set_ylabel("Variance (%)", fontsize=9)
    ax.grid(axis="y", alpha=0.3)
    ax.legend(loc="upper left", fontsize=8, frameon=False)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    return _save(fig, path)


def _month_label(d) -> str:
    if isinstance(d, datetime.datetime):
        d = d.date()
    if isinstance(d, datetime.date):
        return d.strftime("%b-%y").lower()
    return str(d)


# --------------------------------------------------------------------------
# Figura 6: Monthly Breakdown of Fuel Consumption (weekly bars 3 series)
# --------------------------------------------------------------------------

def weekly_breakdown_bars(weekly_rows: list, title: str, path: str) -> str:
    fig, ax = plt.subplots(figsize=(10.0, 4.0))
    ax.set_title(title, fontsize=11, fontweight="bold")
    if not weekly_rows:
        ax.text(0.5, 0.5, "No weekly data available",
                ha="center", va="center", fontsize=10, color="#555555")
        ax.set_axis_off()
        return _save(fig, path)
    labels = ["W%d (%s)" % (int(r.get("week", i + 1)),
                              str(r.get("date", "")))
              for i, r in enumerate(weekly_rows)]
    trans = [_num(r.get("transactions")) for r in weekly_rows]
    eq = [_num(r.get("to_equipment")) for r in weekly_rows]
    tr = [_num(r.get("transfers")) for r in weekly_rows]
    x = list(range(len(weekly_rows)))
    w = 0.27
    ax.bar([i - w for i in x], trans, w, color=EXCEL_BLUE,
           label="Transactions")
    ax.bar(x, eq, w, color=EXCEL_ORANGE, label="To Equipment")
    ax.bar([i + w for i in x], tr, w, color=EXCEL_GRAY,
           label="Transfers")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=8)
    ax.set_ylabel("Volume (L)", fontsize=9)
    ax.grid(axis="y", alpha=0.3)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.15),
              ncol=3, frameon=False, fontsize=9)
    return _save(fig, path)


# --------------------------------------------------------------------------
# Figura 7: Historical Transaction Volume by Type
# --------------------------------------------------------------------------

def historical_outflows_bars(history_rows: list, title: str,
                              path: str) -> str:
    """history_rows = [(label, transactions, to_equip, transfers, other)]."""
    fig, ax = plt.subplots(figsize=(11.0, 4.0))
    ax.set_title(title, fontsize=11, fontweight="bold")
    if not history_rows:
        ax.text(0.5, 0.5, "No historical data available",
                ha="center", va="center", fontsize=10, color="#555555")
        ax.set_axis_off()
        return _save(fig, path)
    labels = [r[0] for r in history_rows]
    trans = [_num(r[1]) for r in history_rows]
    eq = [_num(r[2]) for r in history_rows]
    tr = [_num(r[3]) for r in history_rows]
    x = list(range(len(history_rows)))
    w = 0.27
    ax.bar([i - w for i in x], trans, w, color=EXCEL_BLUE,
           label="Transactions")
    ax.bar(x, eq, w, color=EXCEL_ORANGE, label="To Equipment")
    ax.bar([i + w for i in x], tr, w, color=EXCEL_GRAY,
           label="Transfers")
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=60, ha="right", fontsize=7)
    ax.set_ylabel("Volume (L)", fontsize=9)
    ax.grid(axis="y", alpha=0.3)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18),
              ncol=3, frameon=False, fontsize=9)
    return _save(fig, path)
