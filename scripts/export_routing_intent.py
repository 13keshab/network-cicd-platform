#!/usr/bin/env python3
"""Export validated OSPF, EIGRP and RIPv2 intent from the Excel Source of Truth."""

import argparse
import ipaddress
import json
import sys
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
PROTOCOLS = {"ospf", "eigrp", "rip", "ripv2"}


def enabled(value):
    """Accept only explicit Boolean True or Excel numeric 1."""
    return value is True or (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and value == 1
    )


def parse_settings(row, device, protocol):
    """Parse semicolon-separated key=value settings."""
    settings = {}

    for item in str(row.get("value") or "").split(";"):
        if "=" not in item:
            raise ValueError(
                f"Malformed {protocol.upper()} setting for {device}: {item!r}"
            )

        key, value = item.split("=", 1)
        key, value = key.strip().lower(), value.strip()

        if not key or not value or key in settings:
            raise ValueError(
                f"Invalid or duplicate {protocol.upper()} setting "
                f"for {device}: {item!r}"
            )

        settings[key] = value

    return settings


def empty_protocol_intent():
    """Return a consistent disabled protocol structure."""
    return {
        "enabled": False,
        "networks": [],
    }


def empty_device_intent():
    """Return all supported protocols disabled by default."""
    return {
        "ospf": {
            **empty_protocol_intent(),
            "process_id": None,
            "area": None,
            "network_statements": [],
        },
        "eigrp": {
            **empty_protocol_intent(),
            "as_number": None,
            "network_statements": [],
        },
        "rip": {
            **empty_protocol_intent(),
            "version": 2,
            "network_statements": [],
        },
    }


def export_routing_intent(workbook_path=WORKBOOK):
    """Return validated routing intent for R2 and R3 only."""
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
        device: empty_device_intent()
        for device in sorted(ALLOWED_TARGETS)
    }

    for row in rows:
        if str(row.get("feature") or "").strip().lower() != "routing":
            continue

        device = str(row.get("device_name") or "").strip()
        protocol = str(row.get("setting") or "").strip().lower()

        # Never export deployment intent for R1 or unknown targets.
        if device not in ALLOWED_TARGETS or protocol not in PROTOCOLS:
            continue

        # Disabled protocols stay explicitly disabled in the JSON.
        if not enabled(row.get("enabled")):
            continue

        settings = parse_settings(row, device, protocol)

        if protocol == "ospf":
            required = {"process_id", "area", "network"}
        elif protocol == "eigrp":
            required = {"as_number", "network"}
        else:
            required = {"version", "network"}

        if set(settings) != required:
            missing = required - set(settings)
            unknown = set(settings) - required
            raise ValueError(
                f"Invalid {protocol.upper()} settings for {device}; "
                f"missing={sorted(missing)}, unknown={sorted(unknown)}"
            )

        networks = [
            item.strip()
            for item in settings["network"].split(",")
            if item.strip()
        ]
        if not networks:
            raise ValueError(
                f"No networks specified for enabled {protocol.upper()} on {device}"
            )

        if protocol == "ospf":
            process_id = int(settings["process_id"])
            area = settings["area"]

            if not 1 <= process_id <= 65535:
                raise ValueError(f"Invalid OSPF process ID on {device}")

            if "." in area:
                if ipaddress.ip_address(area).version != 4:
                    raise ValueError(f"Invalid OSPF area on {device}")
            elif not 0 <= int(area) <= 4294967295:
                raise ValueError(f"Invalid OSPF area on {device}")

            statements = []
            for network_text in networks:
                network = ipaddress.ip_network(network_text, strict=True)
                if network.version != 4:
                    raise ValueError("Only IPv4 OSPF networks are supported")

                wildcard = ipaddress.IPv4Address(int(network.hostmask))
                statements.append(
                    f"network {network.network_address} {wildcard} area {area}"
                )

            result[device]["ospf"] = {
                "enabled": True,
                "process_id": process_id,
                "area": area,
                "networks": networks,
                "network_statements": statements,
            }

        elif protocol == "eigrp":
            as_number = int(settings["as_number"])
            if not 1 <= as_number <= 65535:
                raise ValueError(f"Invalid EIGRP AS number on {device}")

            statements = []
            for network_text in networks:
                network = ipaddress.ip_network(network_text, strict=True)
                if network.version != 4:
                    raise ValueError("Only IPv4 EIGRP networks are supported")

                wildcard = ipaddress.IPv4Address(int(network.hostmask))
                statements.append(
                    f"network {network.network_address} {wildcard}"
                )

            result[device]["eigrp"] = {
                "enabled": True,
                "as_number": as_number,
                "networks": networks,
                "network_statements": statements,
            }

        else:
            version = int(settings["version"])
            if version != 2:
                raise ValueError(f"Only RIPv2 is supported on {device}")

            statements = []
            for network_text in networks:
                address = ipaddress.ip_address(network_text)
                if address.version != 4:
                    raise ValueError("Only IPv4 RIP networks are supported")

                first = int(str(address).split(".")[0])
                if not 1 <= first <= 223:
                    raise ValueError(
                        f"Invalid classful RIP network {network_text!r}"
                    )

                prefix = 8 if first <= 126 else 16 if first <= 191 else 24
                major_network = ipaddress.ip_network(
                    f"{address}/{prefix}",
                    strict=False,
                )

                if major_network.network_address != address:
                    raise ValueError(
                        f"RIP requires a classful major network address: "
                        f"{network_text!r}"
                    )

                statements.append(f"network {major_network.network_address}")

            result[device]["rip"] = {
                "enabled": True,
                "version": 2,
                "networks": networks,
                "network_statements": statements,
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
        print(f"{device}:")
        for protocol in ("ospf", "eigrp", "rip"):
            print(f"  {protocol.upper()} enabled = {data[protocol]['enabled']}")


if __name__ == "__main__":
    main()
