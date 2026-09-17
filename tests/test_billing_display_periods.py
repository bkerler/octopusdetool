from datetime import date

from PySide6.QtCore import QDate

from octopusdetool.octopusdetool_gui import OctopusSmartMeterGUI


class _CheckBoxStub:
    def __init__(self, checked: bool):
        self._checked = checked

    def isChecked(self) -> bool:
        return self._checked


class _DateEditStub:
    def __init__(self, value: QDate):
        self._value = value

    def date(self) -> QDate:
        return self._value


def _gui(*, enabled: bool, billing_date: QDate) -> OctopusSmartMeterGUI:
    gui = OctopusSmartMeterGUI.__new__(OctopusSmartMeterGUI)
    gui.existing_data = []
    gui.billing_month_checkbox = _CheckBoxStub(enabled)
    gui.billing_month_start_date_edit = _DateEditStub(billing_date)
    return gui


def test_calendar_month_period_is_unchanged_when_billing_display_is_off() -> None:
    gui = _gui(enabled=False, billing_date=QDate(2026, 1, 26))

    buckets, title, _first_column_title, start_date, end_date, bucket_ranges = gui._build_analysis_buckets(
        "month",
        date(2026, 5, 1),
    )

    assert title == "Mai 2026"
    assert start_date == date(2026, 5, 1)
    assert end_date == date(2026, 5, 31)
    assert len(buckets) == 31
    assert bucket_ranges[0] == (date(2026, 5, 1), date(2026, 5, 1))
    assert bucket_ranges[-1] == (date(2026, 5, 31), date(2026, 5, 31))


def test_calendar_year_period_is_unchanged_when_billing_display_is_off() -> None:
    gui = _gui(enabled=False, billing_date=QDate(2026, 5, 26))

    buckets, title, _first_column_title, start_date, end_date, bucket_ranges = gui._build_analysis_buckets(
        "year",
        date(2026, 1, 1),
    )

    assert title == "2026"
    assert start_date == date(2026, 1, 1)
    assert end_date == date(2026, 12, 31)
    assert len(buckets) == 12
    assert bucket_ranges[0] == (date(2026, 1, 1), date(2026, 1, 31))
    assert bucket_ranges[-1] == (date(2026, 12, 1), date(2026, 12, 31))


def test_billing_month_period_uses_configured_start_day() -> None:
    gui = _gui(enabled=True, billing_date=QDate(2026, 1, 26))

    buckets, title, _first_column_title, start_date, end_date, bucket_ranges = gui._build_analysis_buckets(
        "month",
        date(2026, 5, 1),
    )

    assert title == "Abrechnungsmonat Mai 2026 (26.05.2026 - 25.06.2026)"
    assert start_date == date(2026, 5, 26)
    assert end_date == date(2026, 6, 25)
    assert len(buckets) == 31
    assert bucket_ranges[0] == (date(2026, 5, 26), date(2026, 5, 26))
    assert bucket_ranges[-1] == (date(2026, 6, 25), date(2026, 6, 25))


def test_billing_month_period_crosses_year_boundary() -> None:
    gui = _gui(enabled=True, billing_date=QDate(2026, 1, 26))

    _buckets, title, _first_column_title, start_date, end_date, _bucket_ranges = gui._build_analysis_buckets(
        "month",
        date(2026, 12, 1),
    )

    assert title == "Abrechnungsmonat Dezember 2026 (26.12.2026 - 25.01.2027)"
    assert start_date == date(2026, 12, 26)
    assert end_date == date(2027, 1, 25)


def test_billing_year_period_uses_configured_day_and_month() -> None:
    gui = _gui(enabled=True, billing_date=QDate(2026, 5, 26))

    buckets, title, _first_column_title, start_date, end_date, bucket_ranges = gui._build_analysis_buckets(
        "year",
        date(2026, 1, 1),
    )

    assert title == "Abrechnungsjahr 2026 (26.05.2026 - 25.05.2027)"
    assert start_date == date(2026, 5, 26)
    assert end_date == date(2027, 5, 25)
    assert len(buckets) == 12
    assert bucket_ranges[0] == (date(2026, 5, 26), date(2026, 6, 25))
    assert bucket_ranges[-1] == (date(2027, 4, 26), date(2027, 5, 25))


def test_billing_period_clamps_missing_month_days_to_month_end() -> None:
    gui = _gui(enabled=True, billing_date=QDate(2026, 1, 31))

    _buckets, _title, _first_column_title, start_date, end_date, bucket_ranges = gui._build_analysis_buckets(
        "month",
        date(2026, 4, 1),
    )

    assert start_date == date(2026, 4, 30)
    assert end_date == date(2026, 5, 30)
    assert bucket_ranges[0] == (date(2026, 4, 30), date(2026, 4, 30))
    assert bucket_ranges[-1] == (date(2026, 5, 30), date(2026, 5, 30))
