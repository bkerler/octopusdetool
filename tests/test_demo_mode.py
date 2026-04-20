from octopusdetool.octopusdetool import (
    TARIFF_DYNAMIC,
    TARIFF_INTELLIGENT_12,
    TARIFF_INTELLIGENT_HEAT,
    TARIFF_TWO_ZONES,
    TARIFF_THREE_ZONES,
    get_demo_tariff_profile,
)


def test_demo_mode_two_zones_profile():
    display_name, settings, rates = get_demo_tariff_profile("2")

    assert display_name == TARIFF_INTELLIGENT_12
    assert settings.tariff_type == TARIFF_TWO_ZONES
    assert len(rates) == 2
    assert rates[0].rate_ct == 12.0
    assert rates[1].rate_ct == 29.0


def test_demo_mode_three_zones_profile():
    display_name, settings, rates = get_demo_tariff_profile("3")

    assert display_name == TARIFF_INTELLIGENT_HEAT
    assert settings.tariff_type == TARIFF_THREE_ZONES
    assert len(rates) == 3
    assert {rate.name for rate in rates} == {"Demo Low", "Demo Standard", "Demo High"}


def test_demo_mode_dynamic_profile():
    display_name, settings, rates = get_demo_tariff_profile("dynamically")

    assert display_name == TARIFF_DYNAMIC
    assert settings.tariff_type == TARIFF_DYNAMIC
    assert len(rates) == 4
    assert rates[0].windows[0] == ("00:00", "04:00")
