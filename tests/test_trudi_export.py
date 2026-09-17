from datetime import datetime

from octopusdetool.octopusdetool import parse_trudi_export_readings
from octopusdetool.octopusdetool_gui import OctopusSmartMeterGUI


def test_parse_trudi_export_readings_scales_wh_values_to_kwh(tmp_path) -> None:
    export_path = tmp_path / "trudi_export.xml"
    export_path.write_text(
        """<?xml version="1.0" encoding="utf-8"?>
<UsagePoints xmlns="http://vde.de/AR_2418-6.xsd" xmlns:espi="http://naesb.org/espi">
  <UsagePoint>
    <MeterReading>
      <ReadingType>
        <espi:powerOfTenMultiplier>0</espi:powerOfTenMultiplier>
        <espi:uom>72</espi:uom>
        <scaler>-1</scaler>
        <obisCode>0100010800FF</obisCode>
      </ReadingType>
      <IntervalBlock>
        <IntervalReading>
          <espi:value>52462486</espi:value>
          <timePeriod>
            <start>2026-04-29T08:45:00+02:00</start>
          </timePeriod>
        </IntervalReading>
      </IntervalBlock>
    </MeterReading>
  </UsagePoint>
</UsagePoints>
""",
        encoding="utf-8",
    )

    readings = parse_trudi_export_readings(export_path)

    assert readings == [
        {
            "read_at": datetime(2026, 4, 29, 6, 45),
            "value": 5246.2486,
            "obis_code": "0100010800FF",
        }
    ]


def test_parse_trudi_export_readings_ignores_generation_obis(tmp_path) -> None:
    export_path = tmp_path / "trudi_export.xml"
    export_path.write_text(
        """<?xml version="1.0" encoding="utf-8"?>
<UsagePoints xmlns="http://vde.de/AR_2418-6.xsd" xmlns:espi="http://naesb.org/espi">
  <UsagePoint>
    <MeterReading>
      <ReadingType>
        <scaler>-1</scaler>
        <obisCode>0100020800FF</obisCode>
      </ReadingType>
      <IntervalBlock>
        <IntervalReading>
          <espi:value>1339711</espi:value>
          <timePeriod>
            <start>2026-04-29T08:45:00+02:00</start>
          </timePeriod>
        </IntervalReading>
      </IntervalBlock>
    </MeterReading>
  </UsagePoint>
</UsagePoints>
""",
        encoding="utf-8",
    )

    assert parse_trudi_export_readings(export_path) == []


def test_trudi_cache_uses_selected_account_folder(tmp_path, monkeypatch) -> None:
    gui = OctopusSmartMeterGUI.__new__(OctopusSmartMeterGUI)
    gui._demo_mode = False
    gui.selected_account_number = "A-123"

    def fake_account_cache_dir(account_number):
        return tmp_path / "accounts" / str(account_number)

    monkeypatch.setattr(
        "octopusdetool.octopusdetool_gui.get_account_cache_dir",
        fake_account_cache_dir,
    )

    assert gui._get_trudi_cache_path() == tmp_path / "accounts" / "A-123" / "trudi.yaml"


def test_trudi_cache_round_trip(tmp_path, monkeypatch) -> None:
    gui = OctopusSmartMeterGUI.__new__(OctopusSmartMeterGUI)
    gui._demo_mode = False
    gui.selected_account_number = "A-123"
    gui.trudi_source_path = tmp_path / "trudi_export.xml"
    gui.trudi_readings = [
        {
            "read_at": datetime(2026, 4, 29, 6, 45),
            "value": 5246.2486,
            "obis_code": "0100010800FF",
        }
    ]

    monkeypatch.setattr(
        "octopusdetool.octopusdetool_gui.get_account_cache_dir",
        lambda account_number: tmp_path / "accounts" / str(account_number),
    )

    gui._write_trudi_cache()
    gui._set_trudi_readings([], None)
    gui._read_trudi_cache()

    assert gui.trudi_readings_by_time == {datetime(2026, 4, 29, 6, 45): 5246.2486}
