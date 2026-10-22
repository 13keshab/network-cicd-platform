
from openpyxl import load_workbook

from src.sot_processor.validators import WORKBOOK
from scripts.export_routing_intent import export_routing_intent


def test_current_workbook_enables_ospf_for_r2_only():
    intent = export_routing_intent(WORKBOOK)
    routing = intent["routing_intent"]

    assert set(routing) == {"R2", "R3"}

    assert routing["R2"]["ospf_enabled"] is True
    assert routing["R2"]["process_id"] == 10
    assert routing["R2"]["area"] == "0"
    assert routing["R2"]["networks"] == ["10.10.12.0/30"]

    assert routing["R3"]["ospf_enabled"] is False
    assert routing["R3"]["process_id"] is None
    assert routing["R3"]["area"] is None
    assert routing["R3"]["networks"] == []

    # R1 must never appear in deployment routing intent.
    assert "R1" not in routing


def test_enabled_ospf_is_exported_for_r2_only(tmp_path):
    path = tmp_path / "enabled_ospf.xlsx"
    workbook = load_workbook(WORKBOOK)
    sheet = workbook["Desired_Config"]

    for row in range(2, sheet.max_row + 1):
        if (
            sheet.cell(row, 1).value == "R2"
            and str(sheet.cell(row, 2).value or "").strip().lower()
            == "routing"
            and str(sheet.cell(row, 3).value or "").strip().lower()
            == "ospf"
        ):
            sheet.cell(row, 4).value = (
                "process_id=10; area=0; network=10.10.12.0/30"
            )
            sheet.cell(row, 5).value = True
            break
    else:
        workbook.close()
        raise AssertionError("Could not find the R2 OSPF configuration row")

    workbook.save(path)
    workbook.close()

    intent = export_routing_intent(path)
    routing = intent["routing_intent"]

    assert routing["R2"]["ospf_enabled"] is True
    assert routing["R2"]["process_id"] == 10
    assert routing["R2"]["area"] == "0"
    assert routing["R2"]["networks"] == ["10.10.12.0/30"]

    assert routing["R3"]["ospf_enabled"] is False
    assert routing["R3"]["process_id"] is None
    assert routing["R3"]["area"] is None
    assert routing["R3"]["networks"] == []

    assert "R1" not in routing
