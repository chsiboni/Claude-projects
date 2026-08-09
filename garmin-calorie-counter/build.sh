#!/usr/bin/env bash
# Build the watch app. Requires the Connect IQ SDK on PATH (monkeyc) and a
# developer key.
#
#   ./build.sh                 -> bin/CalSets.prg for marq2 (sideload build)
#   ./build.sh fenix7          -> build for a different device
#   ./build.sh marq2 release   -> bin/CalSets.iq for the Connect IQ Store
#
# Generate a key once, then keep it out of the repo:
#   openssl genrsa -out developer_key.pem 4096
#   openssl pkcs8 -topk8 -inform PEM -outform DER -in developer_key.pem \
#     -out developer_key.der -nocrypt
set -euo pipefail

DEVICE="${1:-marq2}"
MODE="${2:-debug}"
KEY="${DEVELOPER_KEY:-developer_key.der}"

cd "$(dirname "$0")"

if [ ! -f "$KEY" ]; then
  echo "Developer key not found at '$KEY'." >&2
  echo "Create one (see the header of this script) or set DEVELOPER_KEY." >&2
  exit 1
fi

if ! command -v monkeyc >/dev/null 2>&1; then
  echo "monkeyc not on PATH. Add the Connect IQ SDK's bin directory." >&2
  exit 1
fi

mkdir -p bin

if [ "$MODE" = "release" ]; then
  monkeyc -f monkey.jungle -o bin/CalSets.iq -y "$KEY" -e -r
  echo "Built bin/CalSets.iq (store package)"
else
  monkeyc -f monkey.jungle -o bin/CalSets.prg -y "$KEY" -d "$DEVICE"
  echo "Built bin/CalSets.prg for $DEVICE"
  echo "Copy it to GARMIN/APPS on the watch, or run: monkeydo bin/CalSets.prg $DEVICE"
fi
