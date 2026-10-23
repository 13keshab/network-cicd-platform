#!/usr/bin/env bash
set -Eeuo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$PROJECT_ROOT"

if [[ -x "$PROJECT_ROOT/.venv/bin/python" ]]; then
    PYTHON="$PROJECT_ROOT/.venv/bin/python"
elif command -v python3 >/dev/null 2>&1; then
    PYTHON="$(command -v python3)"
elif command -v python >/dev/null 2>&1; then
    PYTHON="$(command -v python)"
else
    echo "ERROR: Python interpreter not found." >&2
    exit 1
fi

echo "Using Python: $PYTHON"
"$PYTHON" --version

echo
echo "=== 1. Validate Excel Source of Truth ==="
"$PYTHON" src/sot_processor/validators.py

echo
echo "=== 2. Generate Ansible inventories ==="
"$PYTHON" src/sot_processor/inventory_generator.py

echo
echo "=== 3. Generate desired configuration artifacts ==="
"$PYTHON" src/sot_processor/config_generator.py

echo
echo "=== 4. Export validated routing intent ==="
"$PYTHON" scripts/export_routing_intent.py

echo
echo "=== 5. Run unit tests ==="
"$PYTHON" -m pytest tests/unit -v

echo
echo "CI checks completed successfully."
