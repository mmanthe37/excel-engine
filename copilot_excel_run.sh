#!/usr/bin/env bash
# Run Excel Engine from this repository in supervised watch mode.
# Usage:
#   ./copilot_excel_run.sh /path/to/workbook.xlsx /path/to/instructions.docx [extra excel-engine args...]
#
# Examples:
#   ./copilot_excel_run.sh "~/Desktop/Module3.xlsx" "~/Desktop/Module3_Instructions.docx"
#   ./copilot_excel_run.sh "./assignment.xlsx" "./instructions.txt" --phase 1 --output run.json

set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

usage() {
  cat <<'EOF'
Usage:
  ./copilot_excel_run.sh <workbook.xlsx> <instructions.(docx|rtfd|pdf|txt)> [extra run args...]

Runs the local repo version of Excel Engine with:
  - repo venv activation
  - editable install check (ensures this repo code is active)
  - real-time watch output (--watch)
  - post-run task-based verification (unless --skip-verify is passed)

Any extra args are passed through to:
  excel-engine run <workbook> <instructions> --watch ...

Special wrapper flag:
  --skip-verify   Skip the final `excel-engine verify --instructions ...` pass
EOF
}

if [[ $# -lt 2 ]]; then
  usage
  exit 1
fi

WORKBOOK="$1"
INSTRUCTIONS="$2"
shift 2

SKIP_VERIFY=0
RUN_ARGS=()
for arg in "$@"; do
  if [[ "$arg" == "--skip-verify" ]]; then
    SKIP_VERIFY=1
  else
    RUN_ARGS+=("$arg")
  fi
done

if [[ ! -f "$WORKBOOK" ]]; then
  echo "Error: workbook not found: $WORKBOOK" >&2
  exit 1
fi

if [[ ! -f "$INSTRUCTIONS" ]]; then
  echo "Error: instructions not found: $INSTRUCTIONS" >&2
  exit 1
fi

cd "$REPO_DIR"

if [[ ! -f ".venv/bin/activate" ]]; then
  echo "Error: missing $REPO_DIR/.venv" >&2
  echo "Create it first:" >&2
  echo "  python3 -m venv .venv && source .venv/bin/activate && pip install -e '.[all]'" >&2
  exit 1
fi

source ".venv/bin/activate"

# Ensure the active package points to this repo (not an older venv install elsewhere).
ACTIVE_PATH="$(python - <<'PY'
import excel_engine
print(excel_engine.__file__)
PY
)"
if [[ "$ACTIVE_PATH" != "$REPO_DIR/"* ]]; then
  echo "Re-linking excel_engine to local repo (editable install)..."
  python -m pip install -q -e .
fi

VERSION="$(python - <<'PY'
import excel_engine
print(excel_engine.__version__)
PY
)"
COMMIT="$(git rev-parse --short HEAD)"
BRANCH="$(git rev-parse --abbrev-ref HEAD)"

echo "Excel Engine activation"
echo "  Repo:    $REPO_DIR"
echo "  Branch:  $BRANCH"
echo "  Commit:  $COMMIT"
echo "  Version: $VERSION"
echo

set +e
if [[ ${#RUN_ARGS[@]} -gt 0 ]]; then
  python -m excel_engine.cli run "$WORKBOOK" "$INSTRUCTIONS" --watch "${RUN_ARGS[@]}"
else
  python -m excel_engine.cli run "$WORKBOOK" "$INSTRUCTIONS" --watch
fi
RUN_RC=$?
set -e

VERIFY_RC=0
if [[ $SKIP_VERIFY -eq 0 ]]; then
  echo
  echo "Post-run verification"
  set +e
  python -m excel_engine.cli verify "$WORKBOOK" --instructions "$INSTRUCTIONS"
  VERIFY_RC=$?
  set -e
fi

if [[ $RUN_RC -ne 0 || $VERIFY_RC -ne 0 ]]; then
  exit 1
fi
