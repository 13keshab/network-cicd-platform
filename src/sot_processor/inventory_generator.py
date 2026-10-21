"""Generate Ansible inventory YAML files from the Excel Source of Truth."""

import argparse
import sys
from pathlib import Path

import yaml
from openpyxl import load_workbook

# Make the project root importable when this file is run directly.
PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.sot_processor.validators import WORKBOOK, validate_workbook, read_rows


PROJECT_ROOT = Path(__file__).resolve().parents[2]
OUTPUT_DIR = PROJECT_ROOT / "generated" / "inventory"


def build_inventory(workbook_path=WORKBOOK):
    """Read a validated workbook and build grouped Ansible inventory data."""
    errors = validate_workbook(workbook_path)

    if errors:
        raise ValueError(
            "Source of Truth validation failed:\n- " + "\n- ".join(errors)
        )

    workbook = load_workbook(
        workbook_path,
        read_only=True,
        data_only=True,
    )

    try:
        devices = read_rows(workbook["Devices"])
    finally:
        workbook.close()

    inventory = {
        "all": {
            "children": {
                "shared": {"hosts": {}},
                "staging": {"hosts": {}},
                "production": {"hosts": {}},
            }
        }
    }

    for device in devices:
        # Skip devices explicitly disabled in the Source of Truth.
        if device.get("enabled") is not True and device.get("enabled") != 1:
            continue

        name = str(device["device_name"]).strip()
        environment = str(device["environment"]).strip().lower()
        address = str(device["management_ip"]).strip()

        host_vars = {
            "ansible_host": address,
            "ansible_connection": str(device["connection"]).strip(),
            "ansible_network_os": "cisco.ios.ios",
        }

        # R1 is shared transit infrastructure.
        if environment in {"shared", "central_transit"}:
            group = "shared"
        elif environment == "staging":
            group = "staging"
        elif environment in {"production", "prod"}:
            group = "production"
        else:
            raise ValueError(
                f"Unsupported environment {environment!r} for device {name}"
            )

        inventory["all"]["children"][group]["hosts"][name] = host_vars

    return inventory


def write_inventories(inventory, output_dir=OUTPUT_DIR):
    """Write combined and environment-specific YAML inventories."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    combined_path = output_dir / "all.yml"
    with combined_path.open("w", encoding="utf-8") as file:
        yaml.safe_dump(
            inventory,
            file,
            sort_keys=False,
            default_flow_style=False,
        )

    children = inventory["all"]["children"]

    for group in ("shared", "staging", "production"):
        environment_inventory = {
            "all": {
                "hosts": {},
                "children": {
                    group: children[group],
                },
            }
        }

        output_path = output_dir / f"{group}.yml"
        with output_path.open("w", encoding="utf-8") as file:
            yaml.safe_dump(
                environment_inventory,
                file,
                sort_keys=False,
                default_flow_style=False,
            )

    return combined_path


def main():
    parser = argparse.ArgumentParser(
        description="Generate Ansible inventories from the Excel Source of Truth."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=WORKBOOK,
        help="Path to the Source of Truth workbook.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=OUTPUT_DIR,
        help="Directory for generated inventory files.",
    )
    args = parser.parse_args()

    try:
        inventory = build_inventory(args.workbook)
        combined_path = write_inventories(inventory, args.output_dir)
    except (ValueError, OSError, KeyError) as exc:
        print(f"INVENTORY GENERATION FAILED: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc

    print("INVENTORY GENERATION PASSED")
    print(f"Combined inventory: {combined_path}")
    for filename in ("shared.yml", "staging.yml", "production.yml"):
        print(f"Generated: {args.output_dir / filename}")


if __name__ == "__main__":
    main()
