#!/usr/bin/env python3
"""Export validated OSPF intent from the Excel Source of Truth for Ansible."""

import argparse
import json
import sys
import ipaddress
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import load_workbook
from src.sot_processor.validators import (
    WORKBOOK,
    read_rows,
    validate_workbook,
)


ALLOWED_TARGETS = {"R2", "R3"}


def enabled(value):
    """Accept only explicit Boolean True or Excel numeric 1."""
    return value is True or (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and value == 1
    )


def export_routing_intent(workbook_path=WORKBOOK):
    """Return routing intent for permitted deployment targets."""
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
        rows = read_rows(workbook["Desired_Config"])
    finally:
        workbook.close()

    result = {
    "R2": {
        "ospf_enabled": False,
        "process_id": None,
        "area": None,
        "networks": [],
        "network_statements": [],
    },
    "R3": {
        "ospf_enabled": False,
        "process_id": None,
        "area": None,
        "networks": [],
        "network_statements": [],
    },
}

    for row in rows:
        if str(row.get("feature") or "").strip().lower() != "routing":
            continue

        device = str(row.get("device_name") or "").strip()
        protocol = str(row.get("setting") or "").strip().lower()

        # R1 is deliberately excluded from exported deployment intent.
        if device not in ALLOWED_TARGETS or protocol != "ospf":
            continue

        if not enabled(row.get("enabled")):
            continue

        settings = {}
        for item in str(row.get("value") or "").split(";"):
            if "=" not in item:
                raise ValueError(
                    f"Malformed OSPF setting for {device}: {item!r}"
                )
            key, value = item.split("=", 1)
            key = key.strip().lower()
            value = value.strip()

            if not key or not value or key in settings:
                raise ValueError(
                    f"Invalid or duplicate OSPF setting for {device}: {item!r}"
                )
            settings[key] = value

        try:
            process_id = int(settings["process_id"])
            networks = [
                item.strip()
                for item in settings["network"].split(",")
            ]
            area = settings["area"].strip()
            if not 1 <= process_id <= 65535 or not networks or any(
                not network for network in networks
            ) or not area:
                raise ValueError("invalid process ID, area, or network list")

            network_statements = []
            for network in networks:
                parsed_network = ipaddress.ip_network(
                    network, strict=False
                )
                wildcard = ipaddress.IPv4Address(
                    int(parsed_network.hostmask)
                ) if parsed_network.version == 4 else None

                if wildcard is None:
                    raise ValueError(
                        "Only IPv4 OSPF network statements are supported"
                    )

                network_statements.append(
                    f"network {parsed_network.network_address} "
                    f"{wildcard} area {area}"
                )
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(
                f"Invalid enabled OSPF settings for {device}"
            ) from exc

        result[device] = {
            "ospf_enabled": True,
            "process_id": process_id,
            "area": area,
            "networks": networks,
            "network_statements": network_statements,
        }

    return {"routing_intent": result}


def main():
    parser = argparse.ArgumentParser(
        description="Export validated routing intent as JSON."
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        default=WORKBOOK,
        help="Path to the Excel Source of Truth workbook.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=PROJECT_ROOT / "generated" / "routing_intent.json",
        help="Output JSON path.",
    )
    args = parser.parse_args()

    try:
        intent = export_routing_intent(args.workbook)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(
            json.dumps(intent, indent=2) + "\n",
            encoding="utf-8",
        )
    except (OSError, ValueError, KeyError, TypeError) as exc:
        print(f"ROUTING INTENT EXPORT FAILED: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc

    print(f"ROUTING INTENT EXPORT PASSED: {args.output}")
    for device, data in intent["routing_intent"].items():
        print(f"{device}: OSPF enabled = {data['ospf_enabled']}")


if __name__ == "__main__":
    main()
