#!/usr/bin/env bash
# =====================================================
# Camera Configuration Batch Apply Script
# =====================================================
# Runs the Camera Config Tool against multiple cameras
# using a specified capability JSON and Excel file.
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
: "${EXCEL_APPLY_FILE:?Set EXCEL_APPLY_FILE in .env}"
: "${CAPS_FILE:?Set CAPS_FILE in .env (filename only)}"
: "${CAMERA_IPS:?Set CAMERA_IPS (space-separated) in .env}"

# --- resolve paths RELATIVE TO SCRIPT ---
CAPS_FILE="$SCRIPT_DIR/../Capabilities/$CAPS_FILE"
EXCEL_APPLY_FILE="$SCRIPT_DIR/../ApplyFiles/$EXCEL_APPLY_FILE"

# Existence checks
if [ ! -f "$CAPS_FILE" ]; then
  echo "❌ CAPS_FILE does not exist: $CAPS_FILE" >&2
  exit 1
fi
if [ ! -f "$EXCEL_APPLY_FILE" ]; then
  echo "❌ EXCEL_APPLY_FILE does not exist: $EXCEL_APPLY_FILE" >&2
  exit 1
fi

# --- parse cameras list ---
read -r -a CAMERA_IPS_ARR <<< "$CAMERA_IPS"
if [ "${#CAMERA_IPS_ARR[@]}" -eq 0 ]; then
  echo "❌ CAMERA_IPS is empty." >&2
  exit 1
fi

PYTHON_BIN="python"
CONFIG_TOOL_SCRIPT="$SCRIPT_DIR/../camera_config_tool_full.py"

# =====================================================
# Loop through each camera
# =====================================================
for ip in "${CAMERA_IPS_ARR[@]}"; do
  echo "-----------------------------------------"
  echo "Applying configuration to camera: $ip"
  echo "Using: $CAPS_FILE + $EXCEL_APPLY_FILE"
  echo "-----------------------------------------"

  if "$PYTHON_BIN" "$CONFIG_TOOL_SCRIPT" apply \
      --ip "$ip" \
      -u "$CAM_USER" \
      -p "$CAM_PASS" \
      --caps "$CAPS_FILE" \
      -i "$EXCEL_APPLY_FILE"; then
    echo "✅ Successfully applied config to $ip"
  else
    echo "❌ Failed to apply config to $ip"
  fi
  echo
done

echo "All done!"
