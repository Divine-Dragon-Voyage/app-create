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
  "WORKFLOW.md"
  "package.json"
  "package-lock.json"
  "create_app.js"
  "bootstrap_windows.ps1"
  "deploy_windows.ps1"
  "developer_url.txt"
  "release_url.txt"
  "启动.cmd"
  "tech_ops/README.md"
  "tech_ops/prepare_embedded_node.ps1"
  "tech_ops/release_windows.cmd"
  "tech_ops/release_mac_linux.sh"
  "user_ops/README.md"
  "user_ops/install_windows.cmd"
  "user_ops/update_windows.cmd"
  "user_ops/run_windows.cmd"
  "user_ops/launcher_windows.ps1"
  "使用说明.md"
)

for file in "${FILES[@]}"; do
  if [[ -f "$SCRIPT_DIR/$file" ]]; then
    mkdir -p "$STAGING_DIR/$(dirname "$file")"
    cp "$SCRIPT_DIR/$file" "$STAGING_DIR/$file"
  fi
done

DIRS=(
  "runtime"
)

for dir in "${DIRS[@]}"; do
  if [[ -d "$SCRIPT_DIR/$dir" ]]; then
    cp -R "$SCRIPT_DIR/$dir" "$STAGING_DIR/$dir"
  fi
done

# Enforce package strategy: do not ship embedded x86 archive by default.
rm -f "$STAGING_DIR"/runtime/node/node-v*-win-x86.zip 2>/dev/null || true

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
