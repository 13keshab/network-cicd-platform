"""Generate Cisco IOS configurations from the Excel Source of Truth."""

import argparse
import ipaddress
import sys
from pathlib import Path

from openpyxl import load_workbook

PROJECT_ROOT = Path(__file__).resolve().parents[2]

if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.sot_processor.validators import (
    WORKBOOK,
    read_rows,
    validate_workbook,
)

DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "generated" / "configs"


def is_enabled(value):
    """Accept an explicit Boolean True or Excel numeric 1."""
    return value is True or (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and value == 1
    )


def parse_settings(value, device, protocol):
    """Parse semicolon-separated key=value settings."""
    settings = {}

    for item in str(value or "").split(";"):
        if "=" not in item:
            raise ValueError(
                f"Malformed {protocol} setting for {device}: {item!r}"
            )

        key, setting_value = item.split("=", 1)
        key = key.strip().lower()
        setting_value = setting_value.strip()

        if not key or not setting_value or key in settings:
            raise ValueError(
                f"Invalid or duplicate {protocol} setting "
                f"for {device}: {item!r}"
            )

        settings[key] = setting_value

    return settings


def build_routing_commands(device, desired_rows):
    """
    Generate OSPF, EIGRP, and RIPv2 commands.

    A protocol is generated only when its workbook row is explicitly enabled.
    """
    commands = []

    for row in desired_rows:
        row_device = str(row.get("device_name") or "").strip()
        feature = str(row.get("feature") or "").strip().lower()
        protocol = str(row.get("setting") or "").strip().lower()

        if row_device != device or feature != "routing":
            continue

        if protocol == "ripv2":
            protocol = "rip"

        if protocol not in {"ospf", "eigrp", "rip"}:
            continue

        if not is_enabled(row.get("enabled")):
            continue

        settings = parse_settings(
            row.get("value"), device, protocol
        )

        # ---------------- OSPF ----------------
        if protocol == "ospf":
            required = {"process_id", "area", "network"}

            if set(settings) != required:
                raise ValueError(
                    f"OSPF settings for {device} must be: "
                    f"{', '.join(sorted(required))}"
                )

            process_id = int(settings["process_id"])

            if not 1 <= process_id <= 65535:
                raise ValueError(
                    f"Invalid OSPF process ID: {process_id}"
                )

            area = settings["area"]

            if "." in area:
                area = str(ipaddress.IPv4Address(area))
            else:
                area_number = int(area)

                if not 0 <= area_number <= 4294967295:
                    raise ValueError(f"Invalid OSPF area: {area}")

                area = str(area_number)

            networks = [
                ipaddress.ip_network(item.strip(), strict=True)
                for item in settings["network"].split(",")
            ]

            if not networks or any(
                network.version != 4 for network in networks
            ):
                raise ValueError("OSPF requires IPv4 networks")

            commands.append(f"router ospf {process_id}")

            for network in networks:
                wildcard = ipaddress.IPv4Address(
                    int(network.hostmask)
                )

                commands.append(
                    f" network {network.network_address} "
                    f"{wildcard} area {area}"
                )

        # ---------------- EIGRP ----------------
        elif protocol == "eigrp":
            required = {"as_number", "network"}

            if set(settings) != required:
                raise ValueError(
                    f"EIGRP settings for {device} must be: "
                    f"{', '.join(sorted(required))}"
                )

            as_number = int(settings["as_number"])

            if not 1 <= as_number <= 65535:
                raise ValueError(
                    f"Invalid EIGRP AS number: {as_number}"
                )

            networks = [
                ipaddress.ip_network(item.strip(), strict=True)
                for item in settings["network"].split(",")
            ]

            if not networks or any(
                network.version != 4 for network in networks
            ):
                raise ValueError("EIGRP requires IPv4 networks")

            commands.append(f"router eigrp {as_number}")

            for network in networks:
                wildcard = ipaddress.IPv4Address(
                    int(network.hostmask)
                )

                commands.append(
                    f" network {network.network_address} {wildcard}"
                )

        # ---------------- RIPv2 ----------------
        elif protocol == "rip":
            required = {"version", "network"}

            if set(settings) != required:
                raise ValueError(
                    f"RIP settings for {device} must be: "
                    f"{', '.join(sorted(required))}"
                )

            if int(settings["version"]) != 2:
                raise ValueError("Only RIPv2 is supported")

            networks = [
                item.strip()
                for item in settings["network"].split(",")
            ]

            if not networks or any(not item for item in networks):
                raise ValueError("RIP network list cannot be empty")

            for network in networks:
                address = ipaddress.IPv4Address(network)
                octets = str(address).split(".")
                first_octet = int(octets[0])

                if first_octet < 128:
                    major_network = f"{first_octet}.0.0.0"
                elif first_octet < 192:
                    major_network = (
                        f"{first_octet}.{octets[1]}.0.0"
                    )
                elif first_octet < 224:
                    major_network = (
                        f"{first_octet}.{octets[1]}."
                        f"{octets[2]}.0"
                    )
                else:
                    raise ValueError(
                        f"Unsupported RIP network: {network}"
                    )

                if network != major_network:
                    raise ValueError(
                        f"RIP requires a classful major network: "
                        f"{network}"
                    )

            commands.append("router rip")
            commands.append(" version 2")

            for network in networks:
                commands.append(f" network {network}")

        commands.append("!")

    return commands


def build_device_configs(workbook_path=WORKBOOK):
    """Build configuration text for each enabled device in the workbook."""

    errors = validate_workbook(workbook_path)

    if errors:
        raise ValueError(
            "Source of Truth validation failed:\n- "
            + "\n- ".join(errors)
        )

    workbook = load_workbook(
        workbook_path,
        read_only=True,
        data_only=True,
    )

    try:
        devices = read_rows(workbook["Devices"])
        interfaces = read_rows(workbook["Interfaces"])
        desired = read_rows(workbook["Desired_Config"])
    finally:
        workbook.close()

    enabled_devices = {
        str(row["device_name"]).strip()
        for row in devices
        if is_enabled(row.get("enabled"))
    }

    # Collect enabled hostname requirements.
    hostnames = {}

    for row in desired:
        feature = str(row.get("feature") or "").strip().lower()

        if feature != "hostname" or not is_enabled(row.get("enabled")):
            continue

        device = str(row.get("device_name") or "").strip()
        hostname = str(row.get("value") or "").strip()

        if device in enabled_devices and hostname:
            hostnames[device] = hostname

    configs = {}

    for device in sorted(enabled_devices):
        lines = []

        hostname = hostnames.get(device)

        if hostname:
            lines.extend([
                f"hostname {hostname}",
                "!",
            ])

        device_interfaces = [
            row
            for row in interfaces
            if str(row.get("device_name") or "").strip() == device
        ]

        for row in sorted(
            device_interfaces,
            key=lambda item: str(item.get("interface") or ""),
        ):
            interface = str(row["interface"]).strip()
            description = str(row.get("description") or "").strip()
            address = str(row["ip_address"]).strip()
            prefix = int(row["prefix_length"])
            admin_state = (
                str(row.get("admin_state") or "").strip().lower()
            )

            parsed = ipaddress.ip_interface(f"{address}/{prefix}")
            subnet_mask = str(parsed.network.netmask)

            lines.append(f"interface {interface}")

            if description:
                lines.append(f" description {description}")

            lines.append(
                f" ip address {parsed.ip} {subnet_mask}"
            )

            if admin_state == "up":
                lines.append(" no shutdown")
            elif admin_state == "down":
                lines.append(" shutdown")
            else:
                raise ValueError(
                    f"Unsupported admin_state {admin_state!r} "
                    f"for {device}/{interface}"
                )

            lines.append("!")

        # Generate all explicitly enabled routing protocols.
        routing_commands = build_routing_commands(device, desired)

        if routing_commands:
            lines.extend(routing_commands)

        configs[device] = "\n".join(lines).rstrip() + "\n"

    return configs


def write_device_configs(configs, output_dir=DEFAULT_OUTPUT_DIR):
    """Write one reviewable configuration file per device."""

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    written_paths = []

    for device, config_text in sorted(configs.items()):
        output_path = output_dir / f"{device}.cfg"
        output_path.write_text(config_text, encoding="utf-8")
        written_paths.append(output_path)

    return written_paths


def main():
    parser = argparse.ArgumentParser(
        description="Generate Cisco IOS configurations from Excel."
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
        default=DEFAULT_OUTPUT_DIR,
        help="Directory for generated configuration files.",
    )

    args = parser.parse_args()

    try:
        configs = build_device_configs(args.workbook)
        paths = write_device_configs(configs, args.output_dir)
    except (ValueError, OSError, KeyError, TypeError) as exc:
        print(
            f"CONFIGURATION GENERATION FAILED: {exc}",
            file=sys.stderr,
        )
        raise SystemExit(1) from exc

    print("CONFIGURATION GENERATION PASSED")

    routing_generated = any(
        line.strip().lower().startswith(
            ("router ospf ", "router eigrp ", "router rip")
        )
        for config in configs.values()
        for line in config.splitlines()
    )

    if routing_generated:
        print("Routing protocol commands were generated.")
    else:
        print("No routing protocol commands were generated.")

    for path in paths:
        print(f"Generated: {path}")


if __name__ == "__main__":
    main()