#!/usr/bin/env bash
set -Eeuo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$PROJECT_ROOT"

echo "=== 1. Validate Excel Source of Truth ==="
python src/sot_processor/validators.py

echo
echo "=== 2. Generate Ansible inventories ==="
python src/sot_processor/inventory_generator.py

echo
echo "=== 3. Generate desired configuration artifacts ==="
python src/sot_processor/config_generator.py

echo
echo "=== 4. Export validated routing intent ==="
python scripts/export_routing_intent.py

echo
echo "=== 5. Run unit tests ==="
python -m pytest tests/unit -v

echo
echo "CI CHECKS PASSED"