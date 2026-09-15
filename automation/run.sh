#!/usr/bin/env bash
# Convenience launcher — activates the venv and runs the automator.
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
# shellcheck disable=SC1091
source "$HERE/.venv/bin/activate"
exec python "$HERE/wlp_automator.py" "$@"
