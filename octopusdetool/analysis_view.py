from __future__ import annotations

from dataclasses import dataclass
from typing import Sequence

from PySide6.QtCharts import (
    QBarCategoryAxis,
    QBarSet,
    QChart,
    QChartView,
    QStackedBarSeries,
    QValueAxis,
)
from PySide6.QtCore import QMargins, Qt
from PySide6.QtGui import QColor, QCursor, QFont, QPainter
from PySide6.QtWidgets import QToolTip


@dataclass(slots=True)
class DisplayBucket:
    axis_label: str
    tooltip_label: str
    go_kwh: float = 0.0
    standard_kwh: float = 0.0

    @property
    def total_kwh(self) -> float:
        return self.go_kwh + self.standard_kwh

    def go_cost_eur(self, tariff_go_ct: float) -> float:
        return self.go_kwh * tariff_go_ct / 100.0

    def standard_cost_eur(self, tariff_standard_ct: float) -> float:
        return self.standard_kwh * tariff_standard_ct / 100.0

    def total_cost_eur(self, tariff_go_ct: float, tariff_standard_ct: float) -> float:
        return self.go_cost_eur(tariff_go_ct) + self.standard_cost_eur(tariff_standard_ct)


class TariffChartView(QChartView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._buckets: list[DisplayBucket] = []
        self._show_currency = False
        self._tariff_go_ct = 0.0
        self._tariff_standard_ct = 0.0
        self._category_axis_title = ""
        self._value_axis_title = ""
        self.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.setStyleSheet("background: transparent; border: none;")
        self.setMinimumHeight(380)
        self.setChart(self._create_chart())

    def update_buckets(
        self,
        buckets: Sequence[DisplayBucket],
        *,
        show_currency: bool,
        tariff_go_ct: float,
        tariff_standard_ct: float,
        category_axis_title: str,
        value_axis_title: str,
    ) -> None:
        self._buckets = list(buckets)
        self._show_currency = show_currency
        self._tariff_go_ct = tariff_go_ct
        self._tariff_standard_ct = tariff_standard_ct
        self._category_axis_title = category_axis_title
        self._value_axis_title = value_axis_title

        chart = self._create_chart()
        self.setChart(chart)

        categories = [bucket.axis_label for bucket in self._buckets]
        if not categories:
            categories = [""]

        go_values = [
            bucket.go_cost_eur(tariff_go_ct) if show_currency else bucket.go_kwh
            for bucket in self._buckets
        ]
        standard_values = [
            bucket.standard_cost_eur(tariff_standard_ct) if show_currency else bucket.standard_kwh
            for bucket in self._buckets
        ]

        go_set = QBarSet("Tarif Go")
        go_set.setColor(QColor("#6f4df6"))
        go_set.setBorderColor(QColor("#6f4df6"))
        go_set.append(go_values or [0.0])

        standard_set = QBarSet("Tarif Standard")
        standard_set.setColor(QColor("#e98bff"))
        standard_set.setBorderColor(QColor("#e98bff"))
        standard_set.append(standard_values or [0.0])

        go_set.hovered.connect(self._on_bar_hovered)
        standard_set.hovered.connect(self._on_bar_hovered)

        series = QStackedBarSeries()
        series.append(go_set)
        series.append(standard_set)

        chart.addSeries(series)

        axis_x = QBarCategoryAxis()
        axis_x.append(categories)
        axis_x.setLabelsColor(QColor("#f4eeff"))
        axis_x.setGridLineVisible(False)

        axis_y = QValueAxis()
        axis_y.setLabelsColor(QColor("#f4eeff"))
        axis_y.setGridLineColor(QColor("#54408f"))
        axis_y.setLabelFormat("%.2f" if show_currency else "%.1f")
        max_value = max((go + standard) for go, standard in zip(go_values, standard_values, strict=False)) if self._buckets else 0.0
        axis_y.setRange(0.0, max(1.0, max_value * 1.15))
        axis_y.applyNiceNumbers()

        chart.addAxis(axis_x, Qt.AlignmentFlag.AlignBottom)
        chart.addAxis(axis_y, Qt.AlignmentFlag.AlignLeft)
        series.attachAxis(axis_x)
        series.attachAxis(axis_y)
        self.viewport().update()

    def _create_chart(self) -> QChart:
        chart = QChart()
        chart.setBackgroundRoundness(18)
        chart.setBackgroundBrush(QColor("#240748"))
        chart.setPlotAreaBackgroundVisible(False)
        chart.setMargins(QMargins(12, 12, 12, 12))
        chart.legend().setVisible(True)
        chart.legend().setAlignment(Qt.AlignmentFlag.AlignBottom)
        chart.legend().setLabelColor(QColor("#f4eeff"))
        return chart

    def paintEvent(self, event) -> None:
        super().paintEvent(event)

        if not self._category_axis_title and not self._value_axis_title:
            return

        painter = QPainter(self.viewport())
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
        font = QFont(self.font())
        font.setPointSizeF(max(font.pointSizeF(), 11.5))
        font.setBold(True)
        painter.setFont(font)
        painter.setPen(QColor("#f4eeff"))

        plot_area = self.chart().plotArea().toRect()
        if plot_area.isEmpty():
            painter.end()
            return

        metrics = painter.fontMetrics()

        if self._value_axis_title:
            value_rect = plot_area.adjusted(0, -metrics.height() - 28, 0, -plot_area.height() - 8)
            painter.drawText(
                value_rect,
                Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignBottom,
                self._value_axis_title,
            )

        if self._category_axis_title:
            category_rect = plot_area.adjusted(0, plot_area.height() + 24, 0, metrics.height() + 44)
            painter.drawText(
                category_rect,
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop,
                self._category_axis_title,
            )

        painter.end()

    def _on_bar_hovered(self, status: bool, index: int) -> None:
        if not status or index >= len(self._buckets):
            QToolTip.hideText()
            return

        bucket = self._buckets[index]
        lines = [bucket.tooltip_label]
        if self._show_currency:
            total_value = bucket.total_cost_eur(self._tariff_go_ct, self._tariff_standard_ct)
            lines.append(f"Gesamt: {self._format_decimal(total_value, 2)} €")
            if bucket.go_kwh:
                lines.append(
                    f"Tarif Go: {self._format_decimal(bucket.go_cost_eur(self._tariff_go_ct), 2)} €"
                )
            if bucket.standard_kwh:
                lines.append(
                    "Tarif Standard: "
                    f"{self._format_decimal(bucket.standard_cost_eur(self._tariff_standard_ct), 2)} €"
                )
        else:
            lines.append(f"Gesamt: {self._format_decimal(bucket.total_kwh, 3)} kWh")
            if bucket.go_kwh:
                lines.append(f"Tarif Go: {self._format_decimal(bucket.go_kwh, 3)} kWh")
            if bucket.standard_kwh:
                lines.append(f"Tarif Standard: {self._format_decimal(bucket.standard_kwh, 3)} kWh")

        QToolTip.showText(QCursor.pos(), "\n".join(lines), self)

    @staticmethod
    def _format_decimal(value: float, decimals: int) -> str:
        formatted = f"{value:,.{decimals}f}"
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".")
