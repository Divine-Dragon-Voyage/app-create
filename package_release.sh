#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
OUTPUT_DIR="${1:-$SCRIPT_DIR/dist}"
RELEASE_NAME="${2:-app-create}"

mkdir -p "$OUTPUT_DIR"

TMP_DIR="$(mktemp -d 2>/dev/null || mktemp -d -t app-create-pack)"
STAGING_DIR="$TMP_DIR/staging"
mkdir -p "$STAGING_DIR"

FILES=(
  "README.md"
  "package.json"
  "package-lock.json"
  "create_app.js"
  "bootstrap_windows.ps1"
  "deploy_windows.ps1"
  "release_url.txt"
  "setup_windows.cmd"
  "install_windows.cmd"
  "update_windows.cmd"
  "run_windows.cmd"
)

for file in "${FILES[@]}"; do
  if [[ -f "$SCRIPT_DIR/$file" ]]; then
    cp "$SCRIPT_DIR/$file" "$STAGING_DIR/$file"
  fi
done

DATE_TAG="$(date +%Y%m%d-%H%M%S)"
VERSION_ZIP="$OUTPUT_DIR/${RELEASE_NAME}-${DATE_TAG}.zip"
LATEST_ZIP="$OUTPUT_DIR/${RELEASE_NAME}-latest.zip"

rm -f "$VERSION_ZIP" "$LATEST_ZIP"
(cd "$STAGING_DIR" && zip -r "$VERSION_ZIP" . >/dev/null)
cp "$VERSION_ZIP" "$LATEST_ZIP"

rm -rf "$TMP_DIR"

echo "[OK] Release package created:"
echo "     $VERSION_ZIP"
echo "[OK] Latest package refreshed:"
echo "     $LATEST_ZIP"
