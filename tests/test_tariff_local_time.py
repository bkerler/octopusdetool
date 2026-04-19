from datetime import datetime, timezone

from octopusdetool.octopusdetool import (
    TARIFF_INTELLIGENT_GO,
    classify_tariff_zone,
    get_tariff_rate_ct,
)


def test_go_tariff_uses_local_time_for_zone_classification() -> None:
    # 2026-01-15 04:00 UTC is 05:00 in Europe/Berlin (CET), which is still low tariff.
    reading_start = datetime(2026, 1, 15, 4, 0, tzinfo=timezone.utc)

    assert classify_tariff_zone(reading_start, TARIFF_INTELLIGENT_GO) == "low"


def test_go_tariff_switches_after_local_five_oclock_window() -> None:
    # 2026-01-15 05:00 UTC is 06:00 in Europe/Berlin (CET), which is standard tariff.
    reading_start = datetime(2026, 1, 15, 5, 0, tzinfo=timezone.utc)

    assert classify_tariff_zone(reading_start, TARIFF_INTELLIGENT_GO) == "standard"


def test_go_tariff_rate_uses_local_time_boundaries() -> None:
    low_rate = 15.92
    standard_rate = 29.13

    assert get_tariff_rate_ct(
        datetime(2026, 1, 15, 4, 0, tzinfo=timezone.utc),
        tariff_go_ct=low_rate,
        tariff_standard_ct=standard_rate,
        tariff_type=TARIFF_INTELLIGENT_GO,
    ) == low_rate

    assert get_tariff_rate_ct(
        datetime(2026, 1, 15, 5, 0, tzinfo=timezone.utc),
        tariff_go_ct=low_rate,
        tariff_standard_ct=standard_rate,
        tariff_type=TARIFF_INTELLIGENT_GO,
    ) == standard_rate
