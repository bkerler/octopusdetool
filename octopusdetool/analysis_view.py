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
    high_kwh: float = 0.0

    @property
    def total_kwh(self) -> float:
        return self.go_kwh + self.standard_kwh + self.high_kwh

    def go_cost_eur(self, tariff_go_ct: float) -> float:
        return self.go_kwh * tariff_go_ct / 100.0

    def standard_cost_eur(self, tariff_standard_ct: float) -> float:
        return self.standard_kwh * tariff_standard_ct / 100.0

    def high_cost_eur(self, tariff_high_ct: float) -> float:
        return self.high_kwh * tariff_high_ct / 100.0

    def total_cost_eur(
        self,
        tariff_go_ct: float,
        tariff_standard_ct: float,
        tariff_high_ct: float = 0.0,
    ) -> float:
        return (
            self.go_cost_eur(tariff_go_ct)
            + self.standard_cost_eur(tariff_standard_ct)
            + self.high_cost_eur(tariff_high_ct)
        )


class TariffChartView(QChartView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._buckets: list[DisplayBucket] = []
        self._show_currency = False
        self._tariff_go_ct = 0.0
        self._tariff_standard_ct = 0.0
        self._tariff_high_ct = 0.0
        self._go_label = "Tarif Go"
        self._standard_label = "Tarif Standard"
        self._high_label = "Tarif Hoch"
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
        tariff_high_ct: float,
        go_label: str,
        standard_label: str,
        high_label: str,
        category_axis_title: str,
        value_axis_title: str,
    ) -> None:
        self._buckets = list(buckets)
        self._show_currency = show_currency
        self._tariff_go_ct = tariff_go_ct
        self._tariff_standard_ct = tariff_standard_ct
        self._tariff_high_ct = tariff_high_ct
        self._go_label = go_label
        self._standard_label = standard_label
        self._high_label = high_label
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
        high_values = [
            bucket.high_cost_eur(tariff_high_ct) if show_currency else bucket.high_kwh
            for bucket in self._buckets
        ]

        go_set = QBarSet(go_label)
        go_set.setColor(QColor("#6f4df6"))
        go_set.setBorderColor(QColor("#6f4df6"))
        go_set.append(go_values or [0.0])

        standard_set = QBarSet(standard_label)
        standard_set.setColor(QColor("#e98bff"))
        standard_set.setBorderColor(QColor("#e98bff"))
        standard_set.append(standard_values or [0.0])

        high_set = QBarSet(high_label)
        high_set.setColor(QColor("#ff8a5b"))
        high_set.setBorderColor(QColor("#ff8a5b"))
        high_set.append(high_values or [0.0])

        go_set.hovered.connect(self._on_bar_hovered)
        standard_set.hovered.connect(self._on_bar_hovered)
        high_set.hovered.connect(self._on_bar_hovered)

        series = QStackedBarSeries()
        series.append(go_set)
        series.append(standard_set)
        if any(value > 0 for value in high_values):
            series.append(high_set)

        chart.addSeries(series)

        axis_x = QBarCategoryAxis()
        axis_x.append(categories)
        axis_x.setLabelsColor(QColor("#f4eeff"))
        axis_x.setGridLineVisible(False)

        axis_y = QValueAxis()
        axis_y.setLabelsColor(QColor("#f4eeff"))
        axis_y.setGridLineColor(QColor("#54408f"))
        axis_y.setLabelFormat("%.2f" if show_currency else "%.1f")
        max_value = (
            max(
                (go + standard + high)
                for go, standard, high in zip(go_values, standard_values, high_values, strict=False)
            )
            if self._buckets
            else 0.0
        )
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
            total_value = bucket.total_cost_eur(
                self._tariff_go_ct,
                self._tariff_standard_ct,
                self._tariff_high_ct,
            )
            lines.append(f"Gesamt: {self._format_decimal(total_value, 2)} €")
            if bucket.go_kwh:
                lines.append(
                    f"{self._go_label}: {self._format_decimal(bucket.go_cost_eur(self._tariff_go_ct), 2)} €"
                )
            if bucket.standard_kwh:
                lines.append(
                    f"{self._standard_label}: "
                    f"{self._format_decimal(bucket.standard_cost_eur(self._tariff_standard_ct), 2)} €"
                )
            if bucket.high_kwh:
                lines.append(
                    f"{self._high_label}: "
                    f"{self._format_decimal(bucket.high_cost_eur(self._tariff_high_ct), 2)} €"
                )
        else:
            lines.append(f"Gesamt: {self._format_decimal(bucket.total_kwh, 3)} kWh")
            if bucket.go_kwh:
                lines.append(f"{self._go_label}: {self._format_decimal(bucket.go_kwh, 3)} kWh")
            if bucket.standard_kwh:
                lines.append(
                    f"{self._standard_label}: {self._format_decimal(bucket.standard_kwh, 3)} kWh"
                )
            if bucket.high_kwh:
                lines.append(f"{self._high_label}: {self._format_decimal(bucket.high_kwh, 3)} kWh")

        QToolTip.showText(QCursor.pos(), "\n".join(lines), self)

    @staticmethod
    def _format_decimal(value: float, decimals: int) -> str:
        formatted = f"{value:,.{decimals}f}"
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".")
