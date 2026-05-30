# -*- coding: utf-8 -*-
"""
Interfaz grafica (PySide6) del Generador de Reporte Mensual de Diesel.

Funcionalidades:
  - Cargar el Excel historico "Reconciliation Monthly.xlsx" y elegir un mes
    de un selector.
  - Cargar opcionalmente el CSV "stock_trend_*" para Figura 3.
  - Cargar opcionalmente PDFs Veridapt mensuales (para actualizar la fila
    del mes en el Excel historico).
  - Cargar opcionalmente el CSV de delivery_transaction (upsert al sheet).
  - Editar manualmente todos los datos (consolidada, service trucks,
    deliveries diarias, dispensing, weekly breakdown, textos, tareas).
  - Generar el .docx final con un click.

Ejecutar:  python run.py
"""
from __future__ import annotations

import datetime
import os
import sys
import traceback

from PySide6.QtCore import Qt, QThread
from PySide6.QtGui import QFont
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDoubleSpinBox, QFileDialog, QFrame, QGroupBox,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow, QMessageBox,
    QPlainTextEdit, QProgressDialog, QPushButton, QScrollArea, QTabWidget,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

import report_model as m
import history
import stock_trend
import pdf_import
import excel_writer
import delivery_csv_import
import monthly_variance_writer

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(HERE, "plantilla_reporte.docx")

PRIMARY = "#1F4E78"
ACCENT = "#2E7D32"
DANGER = "#C62828"
BG = "#F4F6F9"

STYLESHEET = f"""
QMainWindow, QWidget {{ background: {BG}; color: #1A1A1A; }}
QGroupBox {{
    font-weight: bold; color: {PRIMARY};
    border: 1px solid #C9D3DF; border-radius: 8px;
    margin-top: 14px; padding: 10px;
}}
QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 4px; }}
QLabel {{ color: #1A1A1A; }}
QPushButton {{
    background: {PRIMARY}; color: white; border: none;
    border-radius: 6px; padding: 7px 14px; font-weight: bold;
}}
QPushButton:hover {{ background: #2A5F92; }}
QPushButton:disabled {{ background: #9AA8B8; color: white; }}
QPushButton#accent {{ background: {ACCENT}; }}
QPushButton#accent:hover {{ background: #388E3C; }}
QTabWidget::pane {{ border: 1px solid #C9D3DF; border-radius: 6px; background: white; }}
QTabBar::tab {{
    background: #E3E9F0; color: #1A1A1A; padding: 8px 16px; margin-right: 2px;
    border-top-left-radius: 6px; border-top-right-radius: 6px;
}}
QTabBar::tab:selected {{ background: white; color: {PRIMARY}; font-weight: bold; }}
QTabBar::tab:!selected {{ color: #1A1A1A; }}
QTableWidget {{ background: white; gridline-color: #DCE3EB; color: #1A1A1A; }}
QTableWidget QTableCornerButton::section {{ background: {PRIMARY}; }}
QHeaderView::section {{
    background: {PRIMARY}; color: white; padding: 6px; border: none;
    font-weight: bold;
}}
QComboBox, QLineEdit, QPlainTextEdit, QDoubleSpinBox {{
    background: white; color: #1A1A1A; border: 1px solid #C9D3DF;
    border-radius: 5px; padding: 4px;
}}
QComboBox QAbstractItemView {{ background: white; color: #1A1A1A; selection-background-color: #D0E4F7; }}
QLabel#title {{ font-size: 16px; font-weight: bold; color: {PRIMARY}; }}
"""


class _BackgroundWorker(QThread):
    def __init__(self, func, args=(), kwargs=None, holder=None):
        super().__init__()
        self._func = func
        self._args = args
        self._kwargs = kwargs or {}
        self._holder = holder if holder is not None else {}

    def run(self):
        try:
            self._holder["result"] = self._func(*self._args, **self._kwargs)
        except BaseException as exc:  # noqa: BLE001
            self._holder["error"] = exc


def _run_with_progress(parent, title: str, message: str,
                        func, *args, **kwargs):
    """Corre `func` en un hilo aparte con una barra de progreso modal.

    Si `func` acepta un argumento `progress_cb(i, total, label="")`, se le
    inyecta uno que actualiza la barra a porcentaje real, de forma segura
    entre hilos (Qt requiere actualizar widgets desde el hilo de la GUI).
    Si no lo acepta, la barra queda indeterminada (animada)."""
    from PySide6.QtCore import QMetaObject, Q_ARG
    import inspect

    holder: dict = {"result": None, "error": None}

    dlg = QProgressDialog(message, None, 0, 0, parent)
    dlg.setWindowTitle(title)
    dlg.setWindowModality(Qt.ApplicationModal)
    dlg.setMinimumDuration(0)
    dlg.setAutoClose(False)
    dlg.setAutoReset(False)
    dlg.setCancelButton(None)
    dlg.setWindowFlags(
        (dlg.windowFlags() | Qt.CustomizeWindowHint)
        & ~Qt.WindowCloseButtonHint
        & ~Qt.WindowContextHelpButtonHint)
    dlg.setMinimumWidth(440)

    sig_kwargs = dict(kwargs)
    try:
        params = inspect.signature(func).parameters
    except (TypeError, ValueError):
        params = {}

    if "progress_cb" in params:
        def _cb(i, total, label=""):
            try:
                total_i = max(1, int(total))
                cur = max(0, min(int(i), total_i))
            except (TypeError, ValueError):
                return
            QMetaObject.invokeMethod(
                dlg, "setMaximum", Qt.QueuedConnection,
                Q_ARG(int, total_i))
            QMetaObject.invokeMethod(
                dlg, "setValue", Qt.QueuedConnection,
                Q_ARG(int, cur))
            if label:
                QMetaObject.invokeMethod(
                    dlg, "setLabelText", Qt.QueuedConnection,
                    Q_ARG(str, "%s   (%d/%d)" % (label, cur, total_i)))
        sig_kwargs["progress_cb"] = _cb

    worker = _BackgroundWorker(func, args, sig_kwargs, holder)
    worker.finished.connect(dlg.close)
    worker.start()
    dlg.exec()
    worker.wait()
    if holder["error"] is not None:
        raise holder["error"]
    return holder["result"]


def kpi_card(title: str, value: str, color: str) -> QLabel:
    lbl = QLabel(f"<b>{title}</b><br><span style='font-size:14px'>{value}</span>")
    lbl.setTextFormat(Qt.RichText)
    lbl.setAlignment(Qt.AlignCenter)
    lbl.setStyleSheet(
        f"QLabel {{ background: white; border: 2px solid {color}; "
        f"border-radius: 8px; padding: 8px 16px; color: {color}; }}")
    return lbl


def _num(t) -> float:
    return m._num(t)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Generador de Reporte Mensual de Diesel  -  "
                            "Newmont Merian FMS")
        self.resize(1200, 860)
        self.setStyleSheet(STYLESHEET)

        self.data = m.default_data()
        self.history = history.MonthlyHistory()
        self.image_edits: dict = {}
        self.template_path = DEFAULT_TEMPLATE
        self.stock_trend_path: str = ""
        self.pdf_loaded_files: list[str] = []
        self.delivery_csv_path: str = ""
        self.status_chips: dict[str, QLabel] = {}
        # Estado interno: bloquea el callback de itemChanged en refresh.
        self._suspend_signals = False

        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.addWidget(self._build_controls())
        layout.addWidget(self._build_kpis())
        layout.addWidget(self._build_tabs(), stretch=1)
        layout.addWidget(self._build_footer())

        self.statusBar().showMessage(
            "Cargue el Excel historico y elija un mes, o edite los datos "
            "manualmente.")
        self._refresh_all_from_data()

    # ====================================================================
    # Construccion de la interfaz
    # ====================================================================

    def _build_controls(self) -> QWidget:
        box = QGroupBox("Datos y seleccion de mes")
        outer = QVBoxLayout(box)

        row1 = QHBoxLayout()
        btn_hist = QPushButton("Cargar Excel historico...")
        btn_hist.clicked.connect(self._on_load_history)
        btn_trend = QPushButton("Cargar tendencia de tanques (CSV)...")
        btn_trend.clicked.connect(self._on_load_stock_trend)
        btn_pdf = QPushButton("Cargar PDF Veridapt...")
        btn_pdf.clicked.connect(self._on_load_pdf)
        btn_csv = QPushButton("Cargar CSV Deliveries...")
        btn_csv.clicked.connect(self._on_load_delivery_csv)
        btn_in = QPushButton("Cargar Excel de entrada...")
        btn_in.clicked.connect(self._on_load_input)
        btn_blank = QPushButton("Crear Excel en blanco...")
        btn_blank.clicked.connect(self._on_create_blank)
        for b in (btn_hist, btn_trend, btn_pdf, btn_csv, btn_in, btn_blank):
            row1.addWidget(b)
        row1.addStretch(1)
        outer.addLayout(row1)

        # Panel de estado: chips con archivos cargados.
        status_row = QHBoxLayout()
        status_row.setSpacing(8)
        lbl_estado = QLabel("Estado de archivos:")
        lbl_estado.setStyleSheet(
            "QLabel { font-weight: bold; color: #1F4E78; }")
        status_row.addWidget(lbl_estado)
        for key in ("history", "trend", "pdf", "delivery_csv"):
            chip = QLabel()
            chip.setTextFormat(Qt.RichText)
            chip.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)
            chip.setMinimumHeight(26)
            self.status_chips[key] = chip
            status_row.addWidget(chip)
        status_row.addStretch(1)
        outer.addLayout(status_row)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("Mes disponible:"))
        self.date_combo = QComboBox()
        self.date_combo.setMinimumWidth(170)
        row2.addWidget(self.date_combo)
        self.btn_fetch = QPushButton("Traer datos de este mes")
        self.btn_fetch.setObjectName("accent")
        self.btn_fetch.clicked.connect(self._on_fetch_month)
        row2.addWidget(self.btn_fetch)
        row2.addSpacing(16)
        row2.addWidget(QLabel("Umbral % delivery:"))
        self.spin_thr = QDoubleSpinBox()
        self.spin_thr.setRange(-5.0, 0.0)
        self.spin_thr.setSingleStep(0.05)
        self.spin_thr.setDecimals(2)
        self.spin_thr.setValue(-0.50)
        self.spin_thr.setSuffix(" %")
        row2.addWidget(self.spin_thr)
        row2.addSpacing(24)
        row2.addWidget(QLabel("Periodo:"))
        self.period_edit = QLineEdit(self.data["period_full"])
        self.period_edit.setMinimumWidth(190)
        row2.addWidget(self.period_edit)
        row2.addSpacing(8)
        row2.addWidget(QLabel("Mes:"))
        self.month_edit = QLineEdit(self.data.get("month_label", ""))
        self.month_edit.setMinimumWidth(110)
        row2.addWidget(self.month_edit)
        row2.addSpacing(8)
        row2.addWidget(QLabel("Portada:"))
        self.cover_edit = QLineEdit(self.data["cover_date"])
        self.cover_edit.setMinimumWidth(130)
        row2.addWidget(self.cover_edit)
        row2.addStretch(1)
        outer.addLayout(row2)
        return box

    def _build_kpis(self) -> QWidget:
        frame = QFrame()
        frame.setFrameShape(QFrame.StyledPanel)
        frame.setStyleSheet("QFrame { background: transparent; }")
        self.kpi_layout = QHBoxLayout(frame)
        self.kpi_layout.setContentsMargins(0, 0, 0, 0)
        return frame

    def _build_tabs(self) -> QWidget:
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_consolidated_tab(), "Reconciliacion")
        self.tabs.addTab(self._build_trucks_tab(), "Service Trucks")
        self.tabs.addTab(self._build_deliveries_tab(), "Deliveries (diarias)")
        self.tabs.addTab(self._build_dispensing_tab(), "Dispensing / Textos")
        self.tabs.addTab(self._build_weekly_tab(), "Weekly Breakdown")
        self.tabs.addTab(self._build_trend_tab(), "Tendencia mensual")
        self.tabs.addTab(self._build_tasks_tab(), "Tareas")
        self.tabs.addTab(self._build_images_tab(), "Imagenes / Figuras")
        return self.tabs

    def _build_consolidated_tab(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.addWidget(QLabel("Tabla consolidada (LFO Main + Virtual). "
                             "Ingrese Opening, Deliveries, Transactions y "
                             "Closing; el resto se calcula solo."))
        self.tbl_cons = QTableWidget(1, 5)
        self.tbl_cons.setHorizontalHeaderLabels(
            ["Site", "Opening Stock", "Deliveries", "Transactions",
             "Closing Stock"])
        self.tbl_cons.verticalHeader().setVisible(False)
        self.tbl_cons.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)
        self.tbl_cons.itemChanged.connect(self._on_cons_changed)
        lay.addWidget(self.tbl_cons)
        lay.addStretch(1)
        return w

    def _build_trucks_tab(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.addWidget(QLabel("Service Trucks (TFL0846/47/48)."))
        self.tbl_trucks = QTableWidget(0, 3)
        self.tbl_trucks.setHorizontalHeaderLabels(
            ["Site", "Deliveries (inflow)", "Transactions (outflow)"])
        self.tbl_trucks.verticalHeader().setVisible(False)
        self.tbl_trucks.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)
        lay.addWidget(self.tbl_trucks)
        btns = QHBoxLayout()
        btn_add = QPushButton("Agregar fila")
        btn_add.clicked.connect(lambda: self._insert_truck_row("", 0, 0))
        btn_del = QPushButton("Eliminar fila seleccionada")
        btn_del.clicked.connect(self._del_truck_row)
        btns.addWidget(btn_add)
        btns.addWidget(btn_del)
        btns.addStretch(1)
        lay.addLayout(btns)
        return w

    def _build_deliveries_tab(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.addWidget(QLabel("Deliveries diarias del mes "
                             "(una fila por dia). Variance y % se calculan."))
        self.tbl_del = QTableWidget(0, 3)
        self.tbl_del.setHorizontalHeaderLabels(
            ["Fecha (d-mmm)", "Delivery Tickets", "Inlet Deliveries"])
        self.tbl_del.verticalHeader().setVisible(False)
        self.tbl_del.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)
        lay.addWidget(self.tbl_del)
        btns = QHBoxLayout()
        btn_add = QPushButton("Agregar fila")
        btn_add.clicked.connect(lambda: self._insert_delivery_row("", 0, 0))
        btn_del = QPushButton("Eliminar fila seleccionada")
        btn_del.clicked.connect(self._del_delivery_row)
        btns.addWidget(btn_add)
        btns.addWidget(btn_del)
        btns.addStretch(1)
        lay.addLayout(btns)
        return w

    def _build_dispensing_tab(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)

        grp1 = QGroupBox("Distribucion de outflow (Figura 2)")
        l1 = QHBoxLayout(grp1)
        l1.addWidget(QLabel("To Equipment:"))
        self.ed_eq = QLineEdit()
        l1.addWidget(self.ed_eq)
        l1.addWidget(QLabel("Other:"))
        self.ed_ot = QLineEdit()
        l1.addWidget(self.ed_ot)
        l1.addWidget(QLabel("Transfers:"))
        self.ed_tr = QLineEdit()
        l1.addWidget(self.ed_tr)
        lay.addWidget(grp1)

        grp2 = QGroupBox("Estadisticas Tank Log (Figura 3)")
        l2 = QHBoxLayout(grp2)
        l2.addWidget(QLabel("Minimo (L):"))
        self.ed_min = QLineEdit()
        l2.addWidget(self.ed_min)
        l2.addWidget(QLabel("Fecha minimo:"))
        self.ed_min_date = QLineEdit()
        l2.addWidget(self.ed_min_date)
        l2.addWidget(QLabel("Promedio mes (L):"))
        self.ed_avg = QLineEdit()
        l2.addWidget(self.ed_avg)
        lay.addWidget(grp2)

        grp3 = QGroupBox("Monthly Transactions overview (promedios)")
        l3 = QHBoxLayout(grp3)
        l3.addWidget(QLabel("Transactions avg:"))
        self.ed_w_tr = QLineEdit()
        l3.addWidget(self.ed_w_tr)
        l3.addWidget(QLabel("To equipment avg:"))
        self.ed_w_eq = QLineEdit()
        l3.addWidget(self.ed_w_eq)
        l3.addWidget(QLabel("Other avg:"))
        self.ed_w_ot = QLineEdit()
        l3.addWidget(self.ed_w_ot)
        l3.addWidget(QLabel("Transfers avg:"))
        self.ed_w_tf = QLineEdit()
        l3.addWidget(self.ed_w_tf)
        lay.addWidget(grp3)

        lay.addWidget(QLabel("Frase de variance (vacia = automatica):"))
        self.ed_recon = QPlainTextEdit()
        self.ed_recon.setFixedHeight(56)
        lay.addWidget(self.ed_recon)

        lay.addWidget(QLabel("Narrativa de deliveries "
                             "(vacia = automatica):"))
        self.ed_narr = QPlainTextEdit()
        self.ed_narr.setFixedHeight(56)
        lay.addWidget(self.ed_narr)

        lay.addWidget(QLabel("Consideraciones (una por linea):"))
        self.ed_cons = QPlainTextEdit()
        self.ed_cons.setFixedHeight(110)
        lay.addWidget(self.ed_cons)
        return w

    def _build_weekly_tab(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.addWidget(QLabel(
            "Weekly breakdown del mes: una fila por semana cerrada dentro "
            "del mes (Recon_LFO Weekly). Esta tabla alimenta la Tabla 3 y la "
            "Figura 6."))
        self.tbl_weekly = QTableWidget(0, 5)
        self.tbl_weekly.setHorizontalHeaderLabels(
            ["Week", "Date (dd/mm/yy)", "Transactions", "To Equipment",
             "Transfers"])
        self.tbl_weekly.verticalHeader().setVisible(False)
        self.tbl_weekly.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)
        lay.addWidget(self.tbl_weekly)
        btns = QHBoxLayout()
        btn_add = QPushButton("Agregar fila")
        btn_add.clicked.connect(lambda: self._insert_weekly_row(0, "", 0, 0, 0))
        btn_del = QPushButton("Eliminar fila seleccionada")
        btn_del.clicked.connect(self._del_weekly_row)
        btns.addWidget(btn_add)
        btns.addWidget(btn_del)
        btns.addStretch(1)
        lay.addLayout(btns)
        return w

    def _build_trend_tab(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.addWidget(QLabel(
            "Tendencia mensual (Figs. 4 y 5): puntos (Mes, Delivery %, "
            "Recon %). Se carga desde la hoja 'Monthly Variance' del Excel "
            "historico."))
        self.tbl_trend = QTableWidget(0, 3)
        self.tbl_trend.setHorizontalHeaderLabels(
            ["Mes (date)", "Delivery %", "Recon %"])
        self.tbl_trend.verticalHeader().setVisible(False)
        self.tbl_trend.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)
        lay.addWidget(self.tbl_trend)
        btns = QHBoxLayout()
        btn_add = QPushButton("Agregar punto")
        btn_add.clicked.connect(lambda: self._insert_trend_row("", 0, 0))
        btn_del = QPushButton("Eliminar punto seleccionado")
        btn_del.clicked.connect(self._del_trend_row)
        btns.addWidget(btn_add)
        btns.addWidget(btn_del)
        btns.addStretch(1)
        lay.addLayout(btns)
        return w

    def _build_tasks_tab(self) -> QWidget:
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.addWidget(QLabel(
            "Notas/observaciones del bloque 'Pending items / continuous "
            "improvement' (una por linea). NOTA: la Tabla 4 'Recent findings "
            "summary' NO se genera desde el software — se llena a mano "
            "directamente en el .docx generado."))
        self.ed_tasks = QPlainTextEdit()
        lay.addWidget(self.ed_tasks)
        return w

    def _build_images_tab(self) -> QWidget:
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        inner = QWidget()
        lay = QVBoxLayout(inner)
        nota = QLabel(
            "Las Figuras 1 a 7 las construye el software automaticamente:\n"
            " - Fig. 1: barras de Inlet vs Tickets (Deliveries diarias).\n"
            " - Fig. 2: dona de outflow (Dispensing).\n"
            " - Fig. 3: Tank Log (requiere CSV stock_trend).\n"
            " - Fig. 4: tendencia Recon % mensual (Monthly Variance).\n"
            " - Fig. 5: Delivery % + Recon % (Monthly Variance).\n"
            " - Fig. 6: bars semanales del mes (Weekly Breakdown).\n"
            " - Fig. 7: Historical Transaction Volume (PVT OUTFLOW).\n"
            "Solo adjunte una imagen si quiere reemplazar la generada.\n"
            "fig_tasks: imagen opcional de la tabla de tareas si la "
            "preparan aparte.")
        nota.setWordWrap(True)
        nota.setStyleSheet("color:#1F4E78; font-weight:bold;")
        lay.addWidget(nota)
        auto = {"fig1", "fig2", "fig3", "fig4", "fig5", "fig6", "fig7"}
        for fig in m.FIGURE_KEYS:
            row = QHBoxLayout()
            suffix = "   [AUTOMATICA]" if fig in auto else ""
            label = QLabel(m.FIGURE_LABELS[fig] + suffix)
            label.setFixedWidth(360)
            edit = QLineEdit()
            self.image_edits[fig] = edit
            btn = QPushButton("Examinar")
            btn.clicked.connect(lambda _=False, f=fig: self._pick_image(f))
            row.addWidget(label)
            row.addWidget(edit, stretch=1)
            row.addWidget(btn)
            lay.addLayout(row)
        lay.addStretch(1)
        scroll.setWidget(inner)
        return scroll

    def _build_footer(self) -> QWidget:
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        btn_save = QPushButton("Guardar Excel de entrada...")
        btn_save.clicked.connect(self._on_save_input)
        lay.addWidget(btn_save)
        lay.addStretch(1)
        btn_gen = QPushButton("GENERAR REPORTE")
        btn_gen.setObjectName("accent")
        btn_gen.setMinimumWidth(220)
        f = QFont()
        f.setBold(True)
        btn_gen.setFont(f)
        btn_gen.clicked.connect(self._on_generate)
        lay.addWidget(btn_gen)
        return w

    # ====================================================================
    # Sincronizacion datos <-> widgets
    # ====================================================================

    def _refresh_all_from_data(self):
        self._suspend_signals = True
        self.period_edit.setText(self.data["period_full"])
        self.month_edit.setText(self.data.get("month_label", ""))
        self.cover_edit.setText(self.data.get("cover_date", ""))
        self._refresh_consolidated()
        self._refresh_trucks()
        self._refresh_deliveries()
        self._refresh_weekly()
        self._refresh_trend()
        self._refresh_dispensing_tab()
        self._refresh_tasks()
        self._refresh_kpis()
        self._refresh_load_status()
        self._suspend_signals = False

    def _refresh_consolidated(self):
        c = self.data["consolidated"]
        self.tbl_cons.setItem(0, 0, QTableWidgetItem(str(c["site"])))
        self.tbl_cons.setItem(0, 1, QTableWidgetItem(str(c["opening"])))
        self.tbl_cons.setItem(0, 2, QTableWidgetItem(str(c["deliveries"])))
        self.tbl_cons.setItem(0, 3, QTableWidgetItem(str(c["transactions"])))
        self.tbl_cons.setItem(0, 4, QTableWidgetItem(str(c["closing"])))

    def _refresh_trucks(self):
        rows = self.data["service_trucks"]
        self.tbl_trucks.setRowCount(0)
        for r in rows:
            self._insert_truck_row(r["site"], r["deliveries"],
                                    r["transactions"])

    def _insert_truck_row(self, site, deliv, trans):
        row = self.tbl_trucks.rowCount()
        self.tbl_trucks.insertRow(row)
        self.tbl_trucks.setItem(row, 0, QTableWidgetItem(str(site or "")))
        self.tbl_trucks.setItem(row, 1, QTableWidgetItem(str(deliv or "")))
        self.tbl_trucks.setItem(row, 2, QTableWidgetItem(str(trans or "")))

    def _del_truck_row(self):
        row = self.tbl_trucks.currentRow()
        if row >= 0:
            self.tbl_trucks.removeRow(row)

    def _refresh_deliveries(self):
        rows = self.data["deliveries_daily"]
        self.tbl_del.setRowCount(0)
        for r in rows:
            self._insert_delivery_row(r.get("date", ""),
                                       r.get("tickets", 0),
                                       r.get("inlet", 0))

    def _insert_delivery_row(self, date, tickets, inlet):
        row = self.tbl_del.rowCount()
        self.tbl_del.insertRow(row)
        self.tbl_del.setItem(row, 0, QTableWidgetItem(str(date or "")))
        self.tbl_del.setItem(row, 1, QTableWidgetItem(str(tickets or "")))
        self.tbl_del.setItem(row, 2, QTableWidgetItem(str(inlet or "")))

    def _del_delivery_row(self):
        row = self.tbl_del.currentRow()
        if row >= 0:
            self.tbl_del.removeRow(row)

    def _refresh_weekly(self):
        rows = self.data["weekly_breakdown"]
        self.tbl_weekly.setRowCount(0)
        for r in rows:
            self._insert_weekly_row(r.get("week", 0), r.get("date", ""),
                                     r.get("transactions", 0),
                                     r.get("to_equipment", 0),
                                     r.get("transfers", 0))

    def _insert_weekly_row(self, week, date, trans, eq, tr):
        row = self.tbl_weekly.rowCount()
        self.tbl_weekly.insertRow(row)
        self.tbl_weekly.setItem(row, 0, QTableWidgetItem(str(week or "")))
        self.tbl_weekly.setItem(row, 1, QTableWidgetItem(str(date or "")))
        self.tbl_weekly.setItem(row, 2, QTableWidgetItem(str(trans or "")))
        self.tbl_weekly.setItem(row, 3, QTableWidgetItem(str(eq or "")))
        self.tbl_weekly.setItem(row, 4, QTableWidgetItem(str(tr or "")))

    def _del_weekly_row(self):
        row = self.tbl_weekly.currentRow()
        if row >= 0:
            self.tbl_weekly.removeRow(row)

    def _refresh_trend(self):
        rows = self.data["monthly_trend"]
        self.tbl_trend.setRowCount(0)
        for pt in rows:
            d = pt[0]
            if isinstance(d, (datetime.date, datetime.datetime)):
                d_str = d.strftime("%Y-%m-%d") if isinstance(d, datetime.date) \
                    and not isinstance(d, datetime.datetime) \
                    else d.strftime("%Y-%m-%d")
            else:
                d_str = str(d)
            self._insert_trend_row(d_str, pt[1] if len(pt) > 1 else 0,
                                    pt[2] if len(pt) > 2 else 0)

    def _insert_trend_row(self, date, deliv, recon):
        row = self.tbl_trend.rowCount()
        self.tbl_trend.insertRow(row)
        self.tbl_trend.setItem(row, 0, QTableWidgetItem(str(date or "")))
        self.tbl_trend.setItem(row, 1, QTableWidgetItem(str(deliv or "")))
        self.tbl_trend.setItem(row, 2, QTableWidgetItem(str(recon or "")))

    def _del_trend_row(self):
        row = self.tbl_trend.currentRow()
        if row >= 0:
            self.tbl_trend.removeRow(row)

    def _refresh_dispensing_tab(self):
        d = self.data["dispensing"]
        self.ed_eq.setText(str(d["to_equipment"]))
        self.ed_ot.setText(str(d["other"]))
        self.ed_tr.setText(str(d["transfers"]))
        s = self.data["tank_log_stats"]
        self.ed_min.setText(str(s["min_value"]))
        self.ed_min_date.setText(str(s["min_date"]))
        self.ed_avg.setText(str(s["average"]))
        ovw = self.data["monthly_overview"]
        self.ed_w_tr.setText(str(ovw.get("transactions_avg", 0)))
        self.ed_w_eq.setText(str(ovw.get("to_equipment_avg", 0)))
        self.ed_w_ot.setText(str(ovw.get("other_avg", 0)))
        self.ed_w_tf.setText(str(ovw.get("transfers_avg", 0)))
        self.ed_recon.setPlainText(self.data.get("recon_sentence", ""))
        self.ed_narr.setPlainText(self.data.get("delivery_narrative", ""))
        self.ed_cons.setPlainText(
            "\n".join(self.data.get("considerations") or []))

    def _refresh_tasks(self):
        self.ed_tasks.setPlainText("\n".join(self.data.get("tasks_notes")
                                              or []))

    def _refresh_kpis(self):
        # Vacia.
        while self.kpi_layout.count():
            it = self.kpi_layout.takeAt(0)
            if it.widget() is not None:
                it.widget().deleteLater()

        cons = m.compute_consolidated_row(self.data["consolidated"])
        delivery_pct = m.delivery_pct_value(self.data["deliveries_daily"])

        self.kpi_layout.addWidget(kpi_card(
            "Variance %", cons["pct"],
            ACCENT if cons["_pct_value"] >= 0 else DANGER))
        self.kpi_layout.addWidget(kpi_card(
            "Closing Stock", cons["closing"], PRIMARY))
        self.kpi_layout.addWidget(kpi_card(
            "Delivery Variance % (mes)", m.fmt_pct(delivery_pct, 2),
            PRIMARY))
        ds = self.data.get("deliveries_summary") or {}
        self.kpi_layout.addWidget(kpi_card(
            "Confirmed Deliveries", str(int(_num(ds.get("confirmed_count")))),
            PRIMARY))
        self.kpi_layout.addStretch(1)

    def _refresh_load_status(self):
        # Excel historico
        if self.history.is_loaded() and self.history.source_path:
            months = len(self.history.available_months())
            self._set_chip(
                "history", True, "Excel historico",
                "%s &mdash; %d meses"
                % (os.path.basename(self.history.source_path), months))
        else:
            self._set_chip("history", False, "Excel historico")
        # Tendencia
        trend_points = self.data.get("tank_trend") or []
        if self.stock_trend_path and trend_points:
            self._set_chip(
                "trend", True, "Tendencia tanques",
                "%s &mdash; %d pts"
                % (os.path.basename(self.stock_trend_path),
                   len(trend_points)))
        else:
            self._set_chip("trend", False, "Tendencia tanques")
        # PDFs
        if self.pdf_loaded_files:
            self._set_chip(
                "pdf", True, "PDF Veridapt",
                "%d archivo(s)" % len(self.pdf_loaded_files))
        else:
            self._set_chip("pdf", False, "PDF Veridapt")
        # Delivery CSV
        if self.delivery_csv_path:
            self._set_chip(
                "delivery_csv", True, "CSV Deliveries",
                os.path.basename(self.delivery_csv_path))
        else:
            self._set_chip("delivery_csv", False, "CSV Deliveries")

    def _set_chip(self, key: str, loaded: bool, title: str,
                  detail: str = "") -> None:
        chip = self.status_chips.get(key)
        if chip is None:
            return
        if loaded:
            chip.setStyleSheet(
                "QLabel { background: #E6F4EA; color: #1B5E20; "
                "border: 1px solid #2E7D32; border-radius: 12px; "
                "padding: 3px 12px; }")
            html = "<b>&#10003; %s</b>" % title
            if detail:
                html += " <span style='color:#33691E;'>(%s)</span>" % detail
            chip.setText(html)
            chip.setToolTip("%s cargado: %s" % (title, detail or "OK"))
        else:
            chip.setStyleSheet(
                "QLabel { background: #FFF4E5; color: #8B4500; "
                "border: 1px solid #FB8C00; border-radius: 12px; "
                "padding: 3px 12px; }")
            chip.setText("<b>&#9888; %s</b> sin cargar" % title)
            chip.setToolTip("%s aun no se ha cargado en esta sesion." % title)

    # ====================================================================
    # Lectura widgets -> data
    # ====================================================================

    def _collect_data(self) -> dict:
        data = dict(self.data)
        data["period_full"] = self.period_edit.text().strip()
        data["month_label"] = self.month_edit.text().strip()
        data["cover_date"] = self.cover_edit.text().strip()

        data["consolidated"] = {
            "site": self.tbl_cons.item(0, 0).text() if self.tbl_cons.item(0, 0)
            else m.SITE_KEY,
            "opening": _num(self.tbl_cons.item(0, 1).text())
            if self.tbl_cons.item(0, 1) else 0,
            "deliveries": _num(self.tbl_cons.item(0, 2).text())
            if self.tbl_cons.item(0, 2) else 0,
            "transactions": _num(self.tbl_cons.item(0, 3).text())
            if self.tbl_cons.item(0, 3) else 0,
            "closing": _num(self.tbl_cons.item(0, 4).text())
            if self.tbl_cons.item(0, 4) else 0,
        }
        trucks = []
        for r in range(self.tbl_trucks.rowCount()):
            site = self.tbl_trucks.item(r, 0)
            if not site or not site.text().strip():
                continue
            trucks.append({
                "site": site.text().strip(),
                "deliveries": _num(self.tbl_trucks.item(r, 1).text())
                if self.tbl_trucks.item(r, 1) else 0,
                "transactions": _num(self.tbl_trucks.item(r, 2).text())
                if self.tbl_trucks.item(r, 2) else 0,
            })
        data["service_trucks"] = trucks

        delivs = []
        for r in range(self.tbl_del.rowCount()):
            date = self.tbl_del.item(r, 0)
            if not date or not date.text().strip():
                continue
            delivs.append({
                "date": date.text().strip(),
                "tickets": _num(self.tbl_del.item(r, 1).text())
                if self.tbl_del.item(r, 1) else 0,
                "inlet": _num(self.tbl_del.item(r, 2).text())
                if self.tbl_del.item(r, 2) else 0,
            })
        data["deliveries_daily"] = delivs

        weekly = []
        for r in range(self.tbl_weekly.rowCount()):
            wk = self.tbl_weekly.item(r, 0)
            if not wk or not wk.text().strip():
                continue
            weekly.append({
                "week": int(_num(wk.text())),
                "date": self.tbl_weekly.item(r, 1).text().strip()
                if self.tbl_weekly.item(r, 1) else "",
                "transactions": _num(self.tbl_weekly.item(r, 2).text())
                if self.tbl_weekly.item(r, 2) else 0,
                "to_equipment": _num(self.tbl_weekly.item(r, 3).text())
                if self.tbl_weekly.item(r, 3) else 0,
                "transfers": _num(self.tbl_weekly.item(r, 4).text())
                if self.tbl_weekly.item(r, 4) else 0,
            })
        data["weekly_breakdown"] = weekly

        trend = []
        for r in range(self.tbl_trend.rowCount()):
            d_it = self.tbl_trend.item(r, 0)
            if not d_it or not d_it.text().strip():
                continue
            d_txt = d_it.text().strip()
            d_val = self._parse_iso_date(d_txt)
            trend.append((
                d_val or d_txt,
                _num(self.tbl_trend.item(r, 1).text())
                if self.tbl_trend.item(r, 1) else 0,
                _num(self.tbl_trend.item(r, 2).text())
                if self.tbl_trend.item(r, 2) else 0,
            ))
        data["monthly_trend"] = trend

        data["dispensing"] = {
            "to_equipment": _num(self.ed_eq.text()),
            "other": _num(self.ed_ot.text()),
            "transfers": _num(self.ed_tr.text()),
        }
        data["tank_log_stats"] = {
            "min_value": _num(self.ed_min.text()),
            "min_date": self.ed_min_date.text().strip(),
            "average": _num(self.ed_avg.text()),
        }
        data["monthly_overview"] = {
            "transactions_avg": _num(self.ed_w_tr.text()),
            "to_equipment_avg": _num(self.ed_w_eq.text()),
            "other_avg": _num(self.ed_w_ot.text()),
            "transfers_avg": _num(self.ed_w_tf.text()),
        }
        data["recon_sentence"] = self.ed_recon.toPlainText().strip()
        data["delivery_narrative"] = self.ed_narr.toPlainText().strip()
        cons_text = self.ed_cons.toPlainText().strip()
        data["considerations"] = [c.strip() for c in cons_text.split("\n")
                                   if c.strip()] or list(
                                       m.DEFAULT_CONSIDERATIONS)
        tasks_text = self.ed_tasks.toPlainText().strip()
        data["tasks_notes"] = [t.strip() for t in tasks_text.split("\n")
                                if t.strip()]
        # Preserve fields no editables
        for key in ("historical_outflows", "deliveries_summary",
                    "tank_trend", "tank_safe_fill"):
            if key in self.data:
                data[key] = self.data[key]
        return data

    def _parse_iso_date(self, text: str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        return None

    # ====================================================================
    # Slots: carga de archivos
    # ====================================================================

    def _populate_date_combo(self, select_month=None):
        """Rellena el combo "Mes disponible" con los meses de Recon_LFO.

        Si `select_month` es una fecha (1er dia del mes), se selecciona ese
        mes despues de repoblar; de lo contrario se conserva el mes que
        estaba seleccionado antes."""
        prev_iso = self.date_combo.currentData() if not select_month else None
        self.date_combo.blockSignals(True)
        try:
            self.date_combo.clear()
            for m_first in self.history.available_months():
                self.date_combo.addItem(m.format_month_label(m_first),
                                         m_first.isoformat())
            target_iso = (select_month.isoformat()
                           if select_month is not None else prev_iso)
            if target_iso:
                idx = self.date_combo.findData(target_iso)
                if idx >= 0:
                    self.date_combo.setCurrentIndex(idx)
        finally:
            self.date_combo.blockSignals(False)

    def _on_load_history(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Cargar Excel historico", "",
            "Excel files (*.xlsx *.xlsm)")
        if not path:
            return
        try:
            self.history = history.MonthlyHistory()
            _run_with_progress(
                self, "Cargando Excel historico",
                "Leyendo  %s ...\n\nPor favor espere, el archivo puede "
                "ser grande." % os.path.basename(path),
                self.history.load, path)
        except Exception as exc:
            QMessageBox.critical(self, "Error",
                                  "No se pudo leer el Excel:\n%s" % exc)
            return
        months = self.history.available_months()
        if not months:
            QMessageBox.warning(self, "Sin datos",
                                 "El Excel no contiene meses en Recon_LFO.")
            return
        self._populate_date_combo()
        self.statusBar().showMessage(
            "Excel historico cargado: %d meses." % len(months))
        self._refresh_load_status()

    def _on_fetch_month(self):
        if not self.history.is_loaded():
            QMessageBox.warning(self, "Falta Excel",
                                 "Primero cargue el Excel historico.")
            return
        iso = self.date_combo.currentData()
        if not iso:
            return
        month_first = datetime.date.fromisoformat(iso)
        thr = float(self.spin_thr.value())
        summary = history.apply_month_to_data(
            self.data, self.history, month_first, threshold_pct=thr)
        # Preserva tank_trend/safe_fill si ya estan cargadas.
        self._refresh_all_from_data()
        msg = "Mes %s cargado. LFO=%s, Trucks=%s, Deliv tickets=%d, "\
              "Weekly rows=%d, Trend=%d, Hist=%d." % (
            m.format_month_label(month_first),
            "OK" if summary["lfo"] else "FALTA",
            ",".join(summary["trucks"]) or "FALTAN",
            summary["deliveries_count"], summary["weekly_rows"],
            summary["trend_points"], summary["historical_points"])
        self.statusBar().showMessage(msg)

    def _on_load_stock_trend(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Cargar CSV stock_trend", "", "CSV files (*.csv)")
        if not path:
            return
        try:
            trends, safe = _run_with_progress(
                self, "Cargando tendencia de tanques",
                "Leyendo %s..." % os.path.basename(path),
                stock_trend.load_tank_trends, path)
        except Exception as exc:
            QMessageBox.critical(self, "Error",
                                  "No se pudo leer el CSV:\n%s" % exc)
            return
        # Si hay un mes seleccionado, filtrar.
        iso = self.date_combo.currentData()
        if iso:
            month_first = datetime.date.fromisoformat(iso)
            pts = stock_trend.filter_by_month(
                trends.get("Main Tank", []), month_first)
        else:
            pts = trends.get("Main Tank", [])
        self.data["tank_trend"] = pts
        self.data["tank_safe_fill"] = safe.get("Main Tank")
        stats = stock_trend.tank_log_stats(pts)
        self.data["tank_log_stats"] = stats
        self.stock_trend_path = path
        self._refresh_all_from_data()
        self.statusBar().showMessage("Stock trend cargado: %d puntos." %
                                      len(pts))

    @staticmethod
    def _month_first_from_period_end(period_end: str):
        """Convierte el `period_end` del PDF Veridapt (formato dd/mm/yyyy) en
        el 1er dia del mes reportado. Devuelve None si no se puede parsear."""
        if not period_end:
            return None
        try:
            end_d = datetime.datetime.strptime(period_end, "%d/%m/%Y").date()
        except ValueError:
            return None
        if end_d.day == 1:
            # period_end es el 1er dia del mes SIGUIENTE al reportado.
            if end_d.month == 1:
                return end_d.replace(year=end_d.year - 1, month=12, day=1)
            return end_d.replace(month=end_d.month - 1, day=1)
        # period_end es el ultimo dia del mes reportado.
        return end_d.replace(day=1)

    def _on_load_pdf(self):
        if not pdf_import.is_available():
            QMessageBox.warning(self, "Falta pdfplumber",
                                 "Instale pdfplumber: pip install pdfplumber")
            return
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Cargar PDFs Veridapt", "", "PDF files (*.pdf)")
        if not paths:
            return
        try:
            loaded, skipped = _run_with_progress(
                self, "Cargando PDFs Veridapt",
                "Leyendo %d PDF(s)..." % len(paths),
                pdf_import.parse_multiple_pdfs, paths)
        except Exception as exc:
            QMessageBox.critical(self, "Error",
                                  "Fallo al leer los PDFs:\n%s" % exc)
            return
        self.pdf_loaded_files = paths

        if loaded and self.history.is_loaded():
            # 1) Determinar el mes reportado para CADA PDF y avisar si no
            # todos pertenecen al mismo mes.
            per_pdf_month = []
            unparseable = []
            for site, info, fname in loaded:
                mf = self._month_first_from_period_end(
                    info.get("period_end", ""))
                if mf is None:
                    unparseable.append((site, fname,
                                         info.get("period_end", "")))
                else:
                    per_pdf_month.append((site, fname, mf, info))

            if unparseable:
                lines = ["No se pudo determinar el mes de estos PDFs:"]
                for site, fname, raw in unparseable:
                    lines.append("  - %s (%s): period_end=%r"
                                 % (site, fname, raw))
                lines.append("\nEstos PDFs NO se escribiran al Excel "
                              "historico.")
                QMessageBox.warning(self, "Fecha del PDF no reconocida",
                                     "\n".join(lines))

            month_first = None
            if per_pdf_month:
                unique_months = sorted({mf for _, _, mf, _ in per_pdf_month})
                if len(unique_months) > 1:
                    lines = ["Los PDFs cargados pertenecen a meses distintos:"]
                    for mf in unique_months:
                        sites = [s for s, _, mf2, _ in per_pdf_month
                                 if mf2 == mf]
                        lines.append("  - %s: %s"
                                     % (m.format_month_label(mf),
                                        ", ".join(sites)))
                    lines.append("\n¿Desea continuar y escribir cada PDF "
                                  "en su mes correspondiente?")
                    if QMessageBox.question(
                            self, "PDFs de meses distintos",
                            "\n".join(lines),
                            QMessageBox.Yes | QMessageBox.No,
                            QMessageBox.No) != QMessageBox.Yes:
                        per_pdf_month = []
                else:
                    month_first = unique_months[0]

            if per_pdf_month:
                items = [(site, mf, info)
                         for site, _, mf, info in per_pdf_month]

                # 2) Detectar filas legacy (del bug anterior: datos del mes N
                # grabados en la fila 1/N en vez de 1/N+1).
                try:
                    legacy = excel_writer.detect_legacy_rows(
                        self.history.source_path, items)
                except Exception as exc:
                    legacy = []
                    QMessageBox.warning(
                        self, "Aviso",
                        "No se pudo verificar filas legacy:\n%s" % exc)

                if legacy:
                    lines = ["Se detectaron filas en el Excel que parecen "
                              "ser del bug anterior (datos del mes grabados "
                              "en la fecha equivocada):"]
                    for lr in legacy:
                        lines.append(
                            "  - %s, fila %d, fecha %s (deberia ser %s)" % (
                                lr["sheet_name"], lr["row"],
                                lr["legacy_date"].strftime("%d/%m/%Y"),
                                lr["correct_date"].strftime("%d/%m/%Y")))
                    lines.append("\n¿Desea borrar esas filas legacy "
                                  "antes de escribir las nuevas?\n\n"
                                  "Si  = borrar y continuar (recomendado).\n"
                                  "No  = dejar como estan y continuar.\n"
                                  "Cancelar = no escribir nada.")
                    choice = QMessageBox.question(
                        self, "Filas legacy detectadas",
                        "\n".join(lines),
                        QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
                        QMessageBox.Yes)
                    if choice == QMessageBox.Cancel:
                        per_pdf_month = []
                        items = []
                    elif choice == QMessageBox.Yes:
                        try:
                            n = excel_writer.remove_rows(
                                self.history.source_path, legacy)
                            self.statusBar().showMessage(
                                "Filas legacy borradas: %d." % n)
                        except Exception as exc:
                            QMessageBox.critical(
                                self, "Error",
                                "No se pudieron borrar las filas legacy:\n%s"
                                % exc)
                            items = []

            if items:
                try:
                    results, errors = _run_with_progress(
                        self, "Escribiendo Excel historico",
                        "Actualizando hojas Recon_LFO / Recon_FuelTank...",
                        excel_writer.write_multiple,
                        self.history.source_path, items)
                except Exception as exc:
                    QMessageBox.critical(
                        self, "Error",
                        "No se pudo escribir al Excel historico:\n%s" % exc)
                    return
                if errors:
                    QMessageBox.warning(
                        self, "Errores escribiendo Excel",
                        "Errores:\n" + "\n".join(
                            "%s: %s" % (k, v) for k, v in errors))
                # 3) Si alguna fila tenia datos PREVIOS distintos a los del
                # PDF, avisamos: el usuario sobreescribio un mes ya cargado.
                overwritten = [r for (_, r) in results
                                if r.get("action") == "overwrite_diff"]
                if overwritten:
                    lines = ["Las siguientes filas tenian datos DISTINTOS y "
                              "fueron sobreescritas con los del PDF:"]
                    for r in overwritten:
                        lines.append("  - %s (fila %s)"
                                     % (r.get("site"), r.get("row")))
                    lines.append("\nSi esto es un error, recupere la version "
                                  "previa del Excel desde su respaldo.")
                    QMessageBox.warning(
                        self, "Datos sobreescritos", "\n".join(lines))
                # Recargar el Excel para refrescar los datos y poner el mes
                # recien escrito en el combo "Mes disponible".
                try:
                    _run_with_progress(
                        self, "Recargando Excel historico",
                        "Refrescando datos...",
                        self.history.load, self.history.source_path)
                    self._populate_date_combo(select_month=month_first)
                except Exception as exc:
                    QMessageBox.warning(self, "Aviso",
                                         "PDFs aplicados, pero al recargar:\n%s"
                                         % exc)
        skipped_msg = ""
        if skipped:
            skipped_msg = "\nOmitidos:\n" + "\n".join(
                "%s: %s" % (k, v) for k, v in skipped)
        QMessageBox.information(self, "PDFs",
                                 "Procesados: %d.\nCargados: %d.%s" %
                                 (len(paths), len(loaded), skipped_msg))
        self._refresh_load_status()

    def _on_load_delivery_csv(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Cargar CSV delivery_transaction", "",
            "CSV files (*.csv)")
        if not path:
            return
        if not self.history.is_loaded():
            QMessageBox.warning(self, "Falta Excel",
                                 "Primero cargue el Excel historico.")
            return
        try:
            info = _run_with_progress(
                self, "Importando CSV de deliveries",
                "Leyendo CSV y escribiendo al Excel historico...",
                delivery_csv_import.import_csv_to_excel,
                self.history.source_path, path)
        except Exception as exc:
            QMessageBox.critical(self, "Error",
                                  "No se pudo importar el CSV:\n%s" % exc)
            return
        self.delivery_csv_path = path
        # Recargar el Excel para incluir los nuevos tickets y refrescar el
        # combo "Mes disponible" por si la importacion agrego una hoja nueva.
        try:
            _run_with_progress(
                self, "Recargando Excel historico",
                "Refrescando datos...",
                self.history.load, self.history.source_path)
            self._populate_date_combo()
        except Exception as exc:
            QMessageBox.warning(self, "Aviso",
                                 "CSV aplicado, pero al recargar:\n%s" % exc)
        QMessageBox.information(self, "CSV Deliveries",
                                 "Hoja: %s\nNuevos: %d\nActualizados: %d\n"
                                 "Total CSV: %d" %
                                 (info.get("sheet"),
                                  info.get("inserted", 0),
                                  info.get("updated", 0),
                                  info.get("total_csv", 0)))
        self._refresh_load_status()

    def _on_load_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Cargar Excel de entrada", "", "Excel files (*.xlsx)")
        if not path:
            return
        try:
            self.data = _run_with_progress(
                self, "Cargando Excel de entrada",
                "Leyendo %s..." % os.path.basename(path),
                m.load_excel, path)
        except Exception as exc:
            QMessageBox.critical(self, "Error",
                                  "No se pudo cargar el Excel:\n%s" % exc)
            return
        self._refresh_all_from_data()
        self.statusBar().showMessage("Excel de entrada cargado: %s" %
                                      os.path.basename(path))

    def _on_create_blank(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Crear Excel en blanco", "", "Excel files (*.xlsx)")
        if not path:
            return
        try:
            _run_with_progress(
                self, "Creando Excel en blanco",
                "Escribiendo %s..." % os.path.basename(path),
                m.save_excel, m.default_data(), path)
        except Exception as exc:
            QMessageBox.critical(self, "Error",
                                  "No se pudo crear el Excel:\n%s" % exc)
            return
        QMessageBox.information(self, "Listo",
                                 "Excel en blanco creado:\n%s" % path)

    def _on_save_input(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Guardar Excel de entrada", "", "Excel files (*.xlsx)")
        if not path:
            return
        data = self._collect_data()
        try:
            _run_with_progress(
                self, "Guardando Excel de entrada",
                "Escribiendo %s..." % os.path.basename(path),
                m.save_excel, data, path)
        except Exception as exc:
            QMessageBox.critical(self, "Error",
                                  "No se pudo guardar:\n%s" % exc)
            return
        QMessageBox.information(self, "Guardado",
                                 "Excel guardado:\n%s" % path)

    def _on_cons_changed(self, _item):
        if self._suspend_signals:
            return
        # Refresca solo KPIs sin tocar los textos.
        self._refresh_kpis()

    def _pick_image(self, fig: str):
        path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar imagen", "",
            "Images (*.png *.jpg *.jpeg)")
        if path:
            self.image_edits[fig].setText(path)

    def _on_generate(self):
        data = self._collect_data()
        images = {k: e.text().strip()
                  for k, e in self.image_edits.items() if e.text().strip()}
        default_name = "Monthly Reconciliation Diesel Report.docx"
        if data.get("month_label"):
            default_name = ("Monthly Reconciliation Diesel Report - %s.docx"
                            % data["month_label"])
        path, _ = QFileDialog.getSaveFileName(
            self, "Guardar reporte", default_name,
            "Word files (*.docx)")
        if not path:
            return
        try:
            _run_with_progress(
                self, "Generando reporte",
                "Construyendo figuras y renderizando docx...",
                m.generate_report, data, images, self.template_path, path)
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(self, "Error",
                                  "Fallo al generar:\n%s" % exc)
            return
        # Persistir Monthly Variance del mes generado si hay historico.
        try:
            iso = self.date_combo.currentData()
            if self.history.is_loaded() and iso:
                month_first = datetime.date.fromisoformat(iso)
                deliv_pct = m.delivery_pct_value(data["deliveries_daily"])
                recon_pct = m.recon_pct_value(data["consolidated"])
                _run_with_progress(
                    self, "Actualizando Monthly Variance",
                    "Escribiendo % de entrega y reconciliacion del mes...",
                    monthly_variance_writer.upsert_month,
                    self.history.source_path, month_first,
                    deliv_pct, recon_pct)
        except Exception as exc:
            QMessageBox.warning(
                self, "Aviso", "Reporte generado, pero no se pudo actualizar "
                "'Monthly Variance':\n%s" % exc)
        QMessageBox.information(self, "Listo",
                                 "Reporte generado:\n%s" % path)


def launch() -> int:
    app = QApplication.instance() or QApplication(sys.argv)
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(launch())
