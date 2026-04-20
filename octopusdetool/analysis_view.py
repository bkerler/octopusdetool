from __future__ import annotations

from dataclasses import dataclass, field
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
    rate_values_kwh: dict[str, float] = field(default_factory=dict)
    generation_kwh: float = 0.0
    meter_reading_kwh: float = 0.0

    @property
    def total_kwh(self) -> float:
        return sum(self.rate_values_kwh.values())

    @property
    def total_generation_kwh(self) -> float:
        return self.generation_kwh

    @property
    def net_kwh(self) -> float:
        return self.total_kwh - self.total_generation_kwh

    def rate_kwh(self, rate_name: str) -> float:
        return self.rate_values_kwh.get(rate_name, 0.0)

    def rate_cost_eur(self, rate_name: str, rate_ct: float) -> float:
        return self.rate_kwh(rate_name) * rate_ct / 100.0

    def total_cost_eur(self, rate_prices_ct: dict[str, float]) -> float:
        return sum(
            self.rate_cost_eur(rate_name, rate_prices_ct.get(rate_name, 0.0))
            for rate_name in self.rate_values_kwh
        )


class TariffChartView(QChartView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._buckets: list[DisplayBucket] = []
        self._show_currency = False
        self._show_generation = False
        self._rate_order: list[str] = []
        self._rate_prices_ct: dict[str, float] = {}
        self._category_axis_title = ""
        self._value_axis_title = ""
        self._series: QStackedBarSeries | None = None
        self.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.setStyleSheet("background: transparent; border: none;")
        self.setMinimumHeight(380)
        self.setMouseTracking(True)
        self.viewport().setMouseTracking(True)
        self.viewport().setAttribute(Qt.WidgetAttribute.WA_Hover, True)
        self.setChart(self._create_chart())

    def update_buckets(
        self,
        buckets: Sequence[DisplayBucket],
        *,
        show_currency: bool,
        show_generation: bool,
        rate_order: Sequence[str],
        rate_prices_ct: dict[str, float],
        category_axis_title: str,
        value_axis_title: str,
    ) -> None:
        self._buckets = list(buckets)
        self._show_currency = show_currency
        self._show_generation = show_generation
        self._rate_order = list(rate_order)
        self._rate_prices_ct = dict(rate_prices_ct)
        self._category_axis_title = category_axis_title
        self._value_axis_title = value_axis_title

        chart = self._create_chart()
        self.setChart(chart)

        categories = [bucket.axis_label for bucket in self._buckets] or [""]
        palette = [
            "#6f4df6",
            "#e98bff",
            "#ff8a5b",
            "#48c7ff",
            "#a9e34b",
            "#ffcf5a",
        ]

        rate_series_values: list[list[float]] = []
        series = QStackedBarSeries()
        for index, rate_name in enumerate(self._rate_order):
            rate_values = [
                bucket.rate_cost_eur(rate_name, self._rate_prices_ct.get(rate_name, 0.0))
                if show_currency
                else bucket.rate_kwh(rate_name)
                for bucket in self._buckets
            ]
            rate_series_values.append(rate_values)

            bar_set = QBarSet(rate_name)
            color = QColor(palette[index % len(palette)])
            bar_set.setColor(color)
            bar_set.setBorderColor(color)
            bar_set.append(rate_values or [0.0])
            bar_set.hovered.connect(self._on_bar_hovered)
            if any(value > 0 for value in rate_values):
                series.append(bar_set)

        if self._show_generation:
            generation_values = [
                -bucket.total_generation_kwh if not show_currency else 0.0
                for bucket in self._buckets
            ]
            if any(value != 0 for value in generation_values):
                generation_set = QBarSet("Einspeisung")
                generation_color = QColor("#48c7ff")
                generation_set.setColor(generation_color)
                generation_set.setBorderColor(generation_color)
                generation_set.append(generation_values)
                generation_set.hovered.connect(self._on_bar_hovered)
                series.append(generation_set)

        if not series.barSets():
            empty_set = QBarSet("Keine Daten")
            empty_set.setColor(QColor("#54408f"))
            empty_set.setBorderColor(QColor("#54408f"))
            empty_set.append([0.0] * len(categories))
            empty_set.hovered.connect(self._on_bar_hovered)
            series.append(empty_set)

        self._series = series
        chart.addSeries(series)

        axis_x = QBarCategoryAxis()
        axis_x.append(categories)
        axis_x.setLabelsColor(QColor("#f4eeff"))
        axis_x.setGridLineVisible(False)

        axis_y = QValueAxis()
        axis_y.setLabelsColor(QColor("#f4eeff"))
        axis_y.setGridLineColor(QColor("#54408f"))
        axis_y.setLabelFormat("%.2f" if show_currency else "%.1f")
        max_positive = (
            max(sum(values) for values in zip(*rate_series_values, strict=False))
            if self._buckets and rate_series_values
            else 0.0
        )
        max_negative = (
            max((bucket.total_generation_kwh for bucket in self._buckets), default=0.0)
            if self._show_generation
            else 0.0
        )
        lower_bound = -max_negative * 1.15 if max_negative > 0 else 0.0
        upper_bound = max(1.0, max_positive * 1.15)
        axis_y.setRange(lower_bound, upper_bound)
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
        lines.append(f"Zaehlerstand: {self._format_decimal(bucket.meter_reading_kwh, 3)} kWh")
        if self._show_currency:
            lines.append(f"Gesamt: {self._format_decimal(bucket.total_cost_eur(self._rate_prices_ct), 2)} EUR")
            for rate_name in self._rate_order:
                rate_kwh = bucket.rate_kwh(rate_name)
                if rate_kwh:
                    lines.append(
                        f"{rate_name}: "
                        f"{self._format_decimal(bucket.rate_cost_eur(rate_name, self._rate_prices_ct.get(rate_name, 0.0)), 2)} EUR"
                    )
        else:
            lines.append(f"Verbrauch: {self._format_decimal(bucket.total_kwh, 3)} kWh")
            if self._show_generation and bucket.total_generation_kwh:
                lines.append(f"Einspeisung: {self._format_decimal(bucket.total_generation_kwh, 3)} kWh")
                lines.append(f"Netto: {self._format_decimal(bucket.net_kwh, 3)} kWh")
            else:
                lines.append(f"Gesamt: {self._format_decimal(bucket.total_kwh, 3)} kWh")
            for rate_name in self._rate_order:
                rate_kwh = bucket.rate_kwh(rate_name)
                if rate_kwh:
                    lines.append(f"{rate_name}: {self._format_decimal(rate_kwh, 3)} kWh")

        QToolTip.showText(QCursor.pos(), "\n".join(lines), self.viewport())

    @staticmethod
    def _format_decimal(value: float, decimals: int) -> str:
        formatted = f"{value:,.{decimals}f}"
        return formatted.replace(",", "X").replace(".", ",").replace("X", ".")
