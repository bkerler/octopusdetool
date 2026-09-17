import shutil
from datetime import date, datetime, timedelta

import openpyxl

from octopusdetool.analysis_view import DisplayBucket, TariffChartView
from octopusdetool.octopusdetool import (
    APP_TIMEZONE,
    OctopusGermanyClient,
    READING_FREQUENCY_TYPE,
    READING_INTERVAL,
    fill_excel_template,
    get_bundled_excel_template_path,
    normalize_datetime,
)
from octopusdetool.octopusdetool_gui import OctopusSmartMeterGUI


class _CheckBoxStub:
    def __init__(self, checked: bool = True):
        self.checked = checked

    def isChecked(self) -> bool:
        return self.checked


class _ResponseStub:
    status_code = 200
    text = ""

    def __init__(self, body):
        self.body = body

    def json(self):
        return self.body

    def raise_for_status(self):
        return None


def test_both_graphql_paths_request_native_intervals() -> None:
    client = OctopusGermanyClient("", "")
    client.token = "test"
    requests = []

    def fake_request(_query, variables):
        requests.append(variables)
        return {
            "property": {
                "measurements": {
                    "edges": [
                        {
                            "node": {
                                "value": "0.25",
                                "unit": "kWh",
                                "startAt": "2026-01-01T00:00:00Z",
                                "endAt": "2026-01-01T00:15:00Z",
                                "durationInSeconds": 900,
                            }
                        }
                    ],
                    "pageInfo": {"hasNextPage": False, "endCursor": "cursor"},
                }
            }
        }

    client._graphql_request = fake_request
    readings = client.get_consumption_graphql("property", fetch_all=True)
    readings.extend(
        client.get_smart_usage("property", "malo", date(2026, 1, 1))
    )

    assert len(readings) == 3
    assert all(
        request["utilityFilters"][0]["electricityFilters"]["readingFrequencyType"]
        == READING_FREQUENCY_TYPE
        for request in requests
    )
    assert all(reading["end"] - reading["start"] == READING_INTERVAL for reading in readings)


def test_graphql_rate_limit_retries_with_backoff() -> None:
    client = OctopusGermanyClient("", "")
    client.token = "test"
    rate_limit = {
        "errors": [
            {"message": "Too many requests.", "extensions": {"errorCode": "KT-CT-1199"}}
        ],
        "data": {"property": {"measurements": None}},
    }
    success = {"data": {"property": {"measurements": {"edges": []}}}}
    responses = iter([_ResponseStub(rate_limit), _ResponseStub(rate_limit), _ResponseStub(success)])
    retries = []

    client._post_with_retry = lambda **_kwargs: next(responses)
    client.rate_limit_callback = lambda delay, attempt, total: retries.append(
        (delay, attempt, total)
    )

    assert client._graphql_request("query", {}) == {"property": {"measurements": {"edges": []}}}
    assert retries == [(5, 1, 3), (10, 2, 3)]
    assert client.last_error_kind is None


def test_smart_usage_handles_graphql_null_measurements() -> None:
    client = OctopusGermanyClient("", "")
    client.token = "test"
    client._graphql_request = lambda *_args, **_kwargs: {"property": {"measurements": None}}

    assert client.get_smart_usage("property", "malo", date(2026, 1, 1)) == []


def test_graphql_pagination_is_not_limited_to_one_hundred_pages() -> None:
    client = OctopusGermanyClient("", "")
    client.token = "test"
    requests = []

    def fake_request(_query, variables):
        requests.append(variables)
        direction = variables["utilityFilters"][0]["electricityFilters"]["readingDirection"]
        cursor = variables.get("after")
        page = int(cursor.rsplit("-", 1)[1]) + 1 if cursor else 1
        start = datetime(2026, 1, 1) + timedelta(minutes=(page - 1) * 15)
        return {
            "property": {
                "measurements": {
                    "edges": [
                        {
                            "node": {
                                "value": "0.25",
                                "startAt": f"{start.isoformat()}Z",
                                "endAt": f"{(start + READING_INTERVAL).isoformat()}Z",
                            }
                        }
                    ],
                    "pageInfo": {
                        "hasNextPage": page < 101,
                        "endCursor": f"cursor-{page}",
                    },
                }
            }
        }

    client._graphql_request = fake_request
    readings = client.get_consumption_graphql("property", fetch_all=True)

    assert len(requests) == 202
    assert len(readings) == 202


def test_incomplete_day_detection_uses_fifteen_minute_intervals() -> None:
    gui = OctopusSmartMeterGUI.__new__(OctopusSmartMeterGUI)
    day = date(2026, 4, 23)
    readings = [
        {
            "start": normalize_datetime(
                datetime(day.year, day.month, day.day, tzinfo=APP_TIMEZONE)
                + index * READING_INTERVAL
            ),
            "end": normalize_datetime(
                datetime(day.year, day.month, day.day, tzinfo=APP_TIMEZONE)
                + (index + 1) * READING_INTERVAL
            ),
            "consumption_kwh": 0.1,
        }
        for index in range(96)
    ]

    assert gui._expected_intervals_for_day(day) == 96
    assert gui._get_incomplete_days(
        readings,
        datetime(2026, 4, 23),
        datetime(2026, 4, 23, 23, 59, 59),
    ) == []
    assert gui._get_incomplete_days(
        readings[:-1],
        datetime(2026, 4, 23),
        datetime(2026, 4, 23, 23, 59, 59),
    ) == [day]


def test_expected_interval_count_handles_daylight_saving_changes() -> None:
    gui = OctopusSmartMeterGUI.__new__(OctopusSmartMeterGUI)

    assert gui._expected_intervals_for_day(date(2026, 3, 29)) == 92
    assert gui._expected_intervals_for_day(date(2026, 10, 25)) == 100


def test_missing_entry_report_uses_fifteen_minute_steps() -> None:
    gui = OctopusSmartMeterGUI.__new__(OctopusSmartMeterGUI)
    gui.use_local_time_checkbox = _CheckBoxStub(False)
    starts = [
        datetime(2026, 1, 1, 0, 0),
        datetime(2026, 1, 1, 0, 30),
        datetime(2026, 1, 1, 0, 45),
    ]

    readings = [{"start": start} for start in starts]

    assert gui._list_missing_entry_timestamps(readings) == [datetime(2026, 1, 1, 0, 15)]


def test_daily_analysis_keeps_quarter_hour_bars_and_tooltips() -> None:
    gui = OctopusSmartMeterGUI.__new__(OctopusSmartMeterGUI)
    day = date(2026, 1, 1)
    start = datetime(2026, 1, 1, 10, 15)
    gui.existing_data = [
        {
            "start": start,
            "end": start + READING_INTERVAL,
            "consumption_kwh": 0.25,
        }
    ]
    gui.use_local_time_checkbox = _CheckBoxStub(False)
    gui.current_tariff_rates = []
    gui.current_tariff_type = "TWO_ZONES"
    gui.billing_month_checkbox = _CheckBoxStub(False)

    buckets, _title, first_column_title, _start_date, _end_date, _ranges = gui._build_analysis_buckets(
        "day",
        day,
    )

    assert len(buckets) == 96
    assert first_column_title == "Stunde"
    assert buckets[0].axis_label == "00"
    assert [bucket.axis_label for bucket in buckets[1:4]] == ["", "", ""]
    assert buckets[4].axis_label == "01"
    assert buckets[41].total_kwh == 0.25
    assert buckets[41].tooltip_label == "01.01.2026 10:15–10:30"


def test_daily_chart_hour_summary_aggregates_four_quarter_hours() -> None:
    chart = TariffChartView.__new__(TariffChartView)
    buckets = [
        DisplayBucket(
            axis_label=f"{hour:02d}" if quarter == 0 else "",
            tooltip_label=f"01.01.2026 {hour:02d}:{quarter * 15:02d}–00:00",
        )
        for hour in range(24)
        for quarter in range(4)
    ]
    for quarter, bucket in enumerate(buckets[40:44]):
        bucket.rate_values_kwh["GO"] = 0.1
        bucket.rate_values_kwh["STANDARD"] = 0.2
        bucket.generation_kwh = 0.05
        bucket.meter_reading_kwh = 10.0 + quarter
    chart._buckets = buckets
    captured = {}
    chart._show_bucket_tooltip = lambda bucket, *, title=None: captured.update(
        bucket=bucket,
        title=title,
    )

    chart._show_hour_summary(10)

    summary = captured["bucket"]
    assert captured["title"] == "01.01.2026 10:00–11:00"
    assert summary.rate_kwh("GO") == 0.4
    assert summary.rate_kwh("STANDARD") == 0.8
    assert summary.total_generation_kwh == 0.2
    assert summary.meter_reading_kwh == 13.0


def test_excel_export_sums_four_quarter_hour_readings_into_one_hour(tmp_path) -> None:
    template = tmp_path / "template.xlsx"
    output = tmp_path / "output.xlsx"
    shutil.copy2(get_bundled_excel_template_path(), template)

    workbook = openpyxl.load_workbook(template)
    workbook["Verbrauch"]["C9"] = None
    workbook.save(template)
    workbook.close()

    start = datetime(2026, 1, 1)
    readings = [
        {
            "start": start + index * READING_INTERVAL,
            "consumption_kwh": 0.25,
        }
        for index in range(4)
    ]

    assert fill_excel_template(readings, str(template), str(output))

    workbook = openpyxl.load_workbook(output, data_only=False)
    assert workbook["Verbrauch"]["C9"].value == 1.0
    workbook.close()
