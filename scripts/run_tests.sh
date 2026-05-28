#!/bin/bash
# word-agent test runner
# Usage: bash scripts/run_tests.sh [pytest-args]
#
# Examples:
#   bash scripts/run_tests.sh                    # run all tests
#   bash scripts/run_tests.sh -v                 # verbose
#   bash scripts/run_tests.sh -k "font"          # only font tests
#   bash scripts/run_tests.sh --tb=short         # short tracebacks

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

cd "$PROJECT_DIR"

# Check pytest is installed
if ! python3 -m pytest --version > /dev/null 2>&1; then
    echo "pytest not found. Install with: pip install pytest>=7.0"
    exit 1
fi

# Run tests
python3 -m pytest test/ "$@"
