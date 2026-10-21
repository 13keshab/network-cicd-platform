"""Generate reviewable Cisco IOS interface configurations from the workbook."""

import argparse
import ipaddress
import sys
from pathlib import Path

from openpyxl import load_workbook

PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.sot_processor.validators import WORKBOOK, read_rows, validate_workbook


DEFAULT_OUTPUT_DIR = PROJECT_ROOT / "generated" / "configs"


def build_device_configs(workbook_path=WORKBOOK):
    """Return a dictionary mapping device names to IOS configuration text."""
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
        interfaces = read_rows(workbook["Interfaces"])
        desired = read_rows(workbook["Desired_Config"])
    finally:
        workbook.close()

    enabled_devices = {
        str(row["device_name"]).strip()
        for row in devices
        if row.get("enabled") is True or row.get("enabled") == 1
    }

    # Only enabled hostname requirements are used.
    hostnames = {}
    for row in desired:
        if (
            str(row.get("feature") or "").strip().lower() == "hostname"
            and (row.get("enabled") is True or row.get("enabled") == 1)
        ):
            device = str(row.get("device_name") or "").strip()
            hostname = str(row.get("value") or "").strip()

            if device in enabled_devices and hostname:
                hostnames[device] = hostname

    configs = {}

    for device in sorted(enabled_devices):
        lines = []

        hostname = hostnames.get(device)
        if hostname:
            lines.extend([f"hostname {hostname}", "!"])

        device_interfaces = [
            row for row in interfaces
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
            admin_state = str(row.get("admin_state") or "").strip().lower()

            parsed = ipaddress.ip_interface(f"{address}/{prefix}")
            subnet_mask = str(parsed.network.netmask)

            lines.append(f"interface {interface}")
            if description:
                lines.append(f" description {description}")
            lines.append(f" ip address {parsed.ip} {subnet_mask}")

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

        # Routing commands are intentionally excluded in this stage.
        configs[device] = "\n".join(lines).rstrip() + "\n"

    return configs


def write_device_configs(configs, output_dir=DEFAULT_OUTPUT_DIR):
    """Write one reviewable IOS configuration file per device."""
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
        description="Generate reviewable IOS configurations from Excel."
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
        print(f"CONFIGURATION GENERATION FAILED: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc

    print("CONFIGURATION GENERATION PASSED")
    print("Routing protocol commands were not generated.")
    for path in paths:
        print(f"Generated: {path}")


if __name__ == "__main__":
    main()
