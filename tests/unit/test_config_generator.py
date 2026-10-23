from openpyxl import load_workbook
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


def test_current_workbook_generates_ospf_for_r2_only():
    configs = build_device_configs(WORKBOOK)

    assert "router ospf 10" in configs["R2"]
    assert "network 10.10.12.0 0.0.0.3 area 0" in configs["R2"]

    # Shared router R1 and production router R3 remain excluded from OSPF.
    for device in ("R1", "R3"):
        assert "router ospf" not in configs[device].lower()

    assert "router eigrp 100" in configs["R2"]
    assert "router eigrp 100" in configs["R3"]
    assert "router rip" in configs["R2"].lower()
    assert "router rip" in configs["R3"].lower()
    assert "router eigrp" not in configs["R1"].lower()
    assert "router rip" not in configs["R1"].lower()


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



def test_enabled_ospf_generates_ios_commands_for_r2_only(tmp_path):
    workbook_path = tmp_path / "enabled_ospf.xlsx"
    workbook = load_workbook(WORKBOOK)
    sheet = workbook["Desired_Config"]

    for row in range(2, sheet.max_row + 1):
        if (
            sheet.cell(row, 1).value == "R2"
            and str(sheet.cell(row, 2).value or "").lower() == "routing"
            and str(sheet.cell(row, 3).value or "").lower() == "ospf"
        ):
            sheet.cell(row, 4).value = (
                "process_id=10; area=0; network=10.10.12.0/30"
            )
            sheet.cell(row, 5).value = True
            break

    workbook.save(workbook_path)
    workbook.close()

    configs = build_device_configs(workbook_path)

    assert "router ospf 10" in configs["R2"]
    assert "network 10.10.12.0 0.0.0.3 area 0" in configs["R2"]
    assert "router ospf" not in configs["R1"].lower()
    assert "router ospf" not in configs["R3"].lower()


def test_disabled_ospf_generates_no_routing_commands(tmp_path):
    # Create a separate workbook with OSPF disabled; do not change the real SoT.
    disabled_path = tmp_path / "ospf_disabled.xlsx"
    workbook = load_workbook(WORKBOOK)
    sheet = workbook["Desired_Config"]

    for row in range(2, sheet.max_row + 1):
        if (
            sheet.cell(row, 1).value == "R2"
            and str(sheet.cell(row, 2).value or "").strip().lower() == "routing"
            and str(sheet.cell(row, 3).value or "").strip().lower() == "ospf"
        ):
            sheet.cell(row, 5).value = False
            break
    else:
        workbook.close()
        raise AssertionError("Could not find the R2 OSPF row")

    workbook.save(disabled_path)
    workbook.close()

    configs = build_device_configs(disabled_path)

    for config in configs.values():
        assert "router ospf" not in config.lower()


def test_enabled_eigrp_generates_ios_commands(tmp_path):
    workbook_path = tmp_path / "enabled_eigrp.xlsx"
    workbook = load_workbook(WORKBOOK)
    sheet = workbook["Desired_Config"]

    for row in range(2, sheet.max_row + 1):
        if (
            sheet.cell(row, 1).value == "R2"
            and str(sheet.cell(row, 2).value or "").strip().lower()
            == "routing"
            and str(sheet.cell(row, 3).value or "").strip().lower()
            == "eigrp"
        ):
            sheet.cell(row, 4).value = (
                "as_number=100; network=10.10.12.0/30"
            )
            sheet.cell(row, 5).value = True
            break
    else:
        workbook.close()
        raise AssertionError("Could not find the R2 EIGRP row")

    workbook.save(workbook_path)
    workbook.close()

    configs = build_device_configs(workbook_path)

    assert "router eigrp 100" in configs["R2"]
    assert "network 10.10.12.0 0.0.0.3" in configs["R2"]
    assert "router eigrp" not in configs["R1"].lower()
    assert "router eigrp" in configs["R3"].lower()


def test_enabled_ripv2_generates_ios_commands(tmp_path):
    workbook_path = tmp_path / "enabled_rip.xlsx"
    workbook = load_workbook(WORKBOOK)
    sheet = workbook["Desired_Config"]

    for row in range(2, sheet.max_row + 1):
        if (
            sheet.cell(row, 1).value == "R2"
            and str(sheet.cell(row, 2).value or "").strip().lower()
            == "routing"
            and str(sheet.cell(row, 3).value or "").strip().lower()
            == "rip"
        ):
            sheet.cell(row, 4).value = (
                "version=2; network=10.0.0.0"
            )
            sheet.cell(row, 5).value = True
            break
    else:
        workbook.close()
        raise AssertionError("Could not find the R2 RIP row")

    workbook.save(workbook_path)
    workbook.close()

    configs = build_device_configs(workbook_path)

    assert "router rip" in configs["R2"].lower()
    assert " version 2" in configs["R2"].lower()
    assert " network 10.0.0.0" in configs["R2"].lower()
    assert "router rip" not in configs["R1"].lower()
    assert "router rip" in configs["R3"].lower()
