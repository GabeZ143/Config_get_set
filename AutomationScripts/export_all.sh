#!/usr/bin/env bash
# =====================================================
# Camera Configuration Batch Export Script
# =====================================================
# Exports configuration from multiple cameras to XLSX.
# =====================================================

set -euo pipefail

# --- locate and load .env next to this script ---
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ENV_FILE="$SCRIPT_DIR/.env"

if [ -f "$ENV_FILE" ]; then
  # shellcheck disable=SC1090
  source "$ENV_FILE"
else
  echo "❌ .env not found at $ENV_FILE. Copy example.env to .env and fill values." >&2
  exit 1
fi

# --- require essentials ---
: "${CAM_USER:?Set CAM_USER in .env}"
: "${CAM_PASS:?Set CAM_PASS in .env}"
: "${CAPS_FILE:?Set CAPS_FILE in .env (filename only)}"
: "${CAMERA_IPS:?Set CAMERA_IPS (space-separated) in .env}"
: "${EXCEL_EXPORT_FILES:?Set EXCEL_EXPORT_FILES (space-separated) in .env}"

# --- resolve paths RELATIVE TO SCRIPT ---
CAPS_FILE="$SCRIPT_DIR/../Capabilities/$CAPS_FILE"

if [ ! -f "$CAPS_FILE" ]; then
  echo "❌ CAPS_FILE does not exist: $CAPS_FILE" >&2
  exit 1
fi

# Ensure Exports dir exists (relative to script)
EXPORTS_DIR="$SCRIPT_DIR/../Exports"
mkdir -p "$EXPORTS_DIR"

# --- parse lists ---
read -r -a CAMERA_IPS_ARR <<< "$CAMERA_IPS"
read -r -a EXCEL_EXPORT_FILES_ARR <<< "$EXCEL_EXPORT_FILES"

# Prepend ../Exports/ to each filename (now absolute-ish relative to script)
for i in "${!EXCEL_EXPORT_FILES_ARR[@]}"; do
  EXCEL_EXPORT_FILES_ARR[$i]="$EXPORTS_DIR/${EXCEL_EXPORT_FILES_ARR[$i]}"
done

# --- sanity checks ---
if [ "${#CAMERA_IPS_ARR[@]}" -eq 0 ]; then
  echo "❌ CAMERA_IPS is empty." >&2
  exit 1
fi
if [ "${#EXCEL_EXPORT_FILES_ARR[@]}" -eq 0 ]; then
  echo "❌ EXCEL_EXPORT_FILES is empty." >&2
  exit 1
fi
if [ "${#CAMERA_IPS_ARR[@]}" -ne "${#EXCEL_EXPORT_FILES_ARR[@]}" ]; then
  echo "❌ Count mismatch: ${#CAMERA_IPS_ARR[@]} CAMERA_IPS vs ${#EXCEL_EXPORT_FILES_ARR[@]} EXCEL_EXPORT_FILES." >&2
  exit 1
fi

PYTHON_BIN="python"
CONFIG_TOOL_SCRIPT="$SCRIPT_DIR/../camera_config_tool_full.py"

# =====================================================
# Loop through each camera (index-based)
# =====================================================
for i in "${!CAMERA_IPS_ARR[@]}"; do
  ip="${CAMERA_IPS_ARR[$i]}"
  out_xlsx="${EXCEL_EXPORT_FILES_ARR[$i]}"

  echo "-----------------------------------------"
  echo "Exporting configuration from camera: $ip"
  echo "Using caps: $CAPS_FILE"
  echo "Output: $out_xlsx"
  echo "-----------------------------------------"

  if "$PYTHON_BIN" "$CONFIG_TOOL_SCRIPT" export \
        --ip "$ip" \
        -u "$CAM_USER" \
        -p "$CAM_PASS" \
        --caps "$CAPS_FILE" \
        -o "$out_xlsx"; then
    echo "✅ Successfully exported config from $ip to $(basename "$out_xlsx")"
  else
    echo "❌ Failed to export config from $ip"
  fi
  echo
done

echo "All done!"
