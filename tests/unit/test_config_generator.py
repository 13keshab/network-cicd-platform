from pathlib import Path

from src.sot_processor.config_generator import (
    build_device_configs,
    write_device_configs,
)
from src.sot_processor.validators import WORKBOOK


def test_configs_are_generated_for_all_enabled_devices():
    configs = build_device_configs(WORKBOOK)

    assert set(configs) == {"R1", "R2", "R3"}


def test_r2_interface_configuration_is_correct():
    configs = build_device_configs(WORKBOOK)
    r2_config = configs["R2"]

    assert "hostname R2" in r2_config
    assert "interface GigabitEthernet1/0" in r2_config
    assert "description TO-R1-STAGING" in r2_config
    assert "ip address 10.10.12.2 255.255.255.252" in r2_config
    assert " no shutdown" in r2_config


def test_routing_protocol_commands_are_not_generated():
    configs = build_device_configs(WORKBOOK)

    for config in configs.values():
        assert "router ospf" not in config.lower()
        assert "router eigrp" not in config.lower()
        assert "router rip" not in config.lower()


def test_config_files_are_written(tmp_path):
    configs = build_device_configs(WORKBOOK)

    paths = write_device_configs(configs, tmp_path)

    assert {path.name for path in paths} == {
        "R1.cfg", "R2.cfg", "R3.cfg"
    }

    for path in paths:
        assert path.is_file()
        assert path.read_text(encoding="utf-8").strip()


def test_invalid_workbook_is_rejected(tmp_path):
    from openpyxl import load_workbook

    invalid_workbook = tmp_path / "invalid.xlsx"
    workbook = load_workbook(WORKBOOK)
    workbook["Interfaces"]["D2"] = "999.999.999.999"
    workbook.save(invalid_workbook)
    workbook.close()

    try:
        build_device_configs(invalid_workbook)
    except ValueError as exc:
        assert "Source of Truth validation failed" in str(exc)
    else:
        raise AssertionError("Invalid workbook should not generate configs")
