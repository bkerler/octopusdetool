from datetime import datetime

from octopusdetool.octopusdetool import _extract_reading_direction, merge_readings


def test_extract_reading_direction_accepts_direct_graphql_shape() -> None:
    node = {
        "metaData": {
            "utilityFilters": {
                "readingDirection": "GENERATION",
            },
        },
    }

    assert _extract_reading_direction(node) == "GENERATION"


def test_extract_reading_direction_accepts_nested_graphql_shape() -> None:
    node = {
        "metaData": {
            "utilityFilters": {
                "electricityFilters": {
                    "readingDirection": "GENERATION",
                },
            },
        },
    }

    assert _extract_reading_direction(node) == "GENERATION"


def test_merge_keeps_consumption_and_generation_for_same_interval() -> None:
    start = datetime(2026, 4, 23, 10)
    end = datetime(2026, 4, 23, 11)

    readings = merge_readings(
        [
            {
                "start": start,
                "end": end,
                "direction": "CONSUMPTION",
                "energy_kwh": 0.4,
                "consumption_kwh": 0.4,
                "net_kwh": 0.4,
            },
            {
                "start": start,
                "end": end,
                "direction": "GENERATION",
                "energy_kwh": 0.3,
                "consumption_kwh": 0.3,
                "net_kwh": -0.3,
            },
        ]
    )

    assert len(readings) == 2
    assert round(sum(reading["net_kwh"] for reading in readings), 3) == 0.1
