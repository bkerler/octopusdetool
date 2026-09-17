from datetime import date, datetime

from octopusdetool.octopusdetool import (
    APP_TIMEZONE,
    READING_INTERVAL,
    TariffAgreement,
    normalize_datetime,
)
from octopusdetool.octopusdetool_gui import OctopusSmartMeterGUI


def _gui() -> OctopusSmartMeterGUI:
    return OctopusSmartMeterGUI.__new__(OctopusSmartMeterGUI)


def test_incomplete_day_detection_uses_local_day_boundaries() -> None:
    gui = _gui()
    local_day = date(2026, 4, 23)
    readings = []

    for index in range(96):
        start_local = datetime(
            local_day.year,
            local_day.month,
            local_day.day,
            tzinfo=APP_TIMEZONE,
        ) + index * READING_INTERVAL
        readings.append(
            {
                "start": normalize_datetime(start_local),
                "end": normalize_datetime(start_local + READING_INTERVAL),
                "consumption_kwh": 0.1,
            }
        )

    assert gui._get_incomplete_days(
        readings,
        datetime(2026, 4, 23),
        datetime(2026, 4, 23, 23, 59, 59),
    ) == []


def test_incomplete_day_detection_reports_local_day_with_missing_quarter_hours() -> None:
    gui = _gui()
    local_day = date(2026, 4, 23)
    readings = []

    for index in range(95):
        start_local = datetime(
            local_day.year,
            local_day.month,
            local_day.day,
            tzinfo=APP_TIMEZONE,
        ) + index * READING_INTERVAL
        readings.append(
            {
                "start": normalize_datetime(start_local),
                "end": normalize_datetime(start_local + READING_INTERVAL),
                "consumption_kwh": 0.1,
            }
        )

    assert gui._get_incomplete_days(
        readings,
        datetime(2026, 4, 23),
        datetime(2026, 4, 23, 23, 59, 59),
    ) == [local_day]


def test_agreement_filter_skips_days_before_active_agreement_start() -> None:
    gui = _gui()
    agreement = TariffAgreement(
        display_name="Intelligent Octopus 12",
        valid_from="2026-04-03T22:00:00+00:00",
        valid_to="2027-04-03T22:00:00+00:00",
        agreement_id="3409850",
    )

    assert gui._filter_days_for_agreement(
        [date(2026, 1, 28), date(2026, 4, 4), date(2027, 4, 4)],
        agreement,
    ) == [date(2026, 4, 4)]


def test_agreement_date_bounds_use_local_start_and_exclusive_end() -> None:
    gui = _gui()
    agreement = TariffAgreement(
        display_name="Intelligent Octopus 12",
        valid_from="2026-04-03T22:00:00+00:00",
        valid_to="2027-04-03T22:00:00+00:00",
        agreement_id="3409850",
    )

    assert gui._agreement_date_bounds(agreement) == (
        date(2026, 4, 4),
        date(2027, 4, 3),
    )
