#!/bin/bash
# =====================================================
# Camera Configuration Batch Apply Script
# =====================================================
# Runs the Camera Config Tool against multiple cameras
# using a specified capability JSON and Excel file.
# =====================================================

# === User-configurable variables ===
# Array of camera IPs
CAMERA_IPS=(
  "10.20.2.121"
  "10.20.2.122"
  "10.20.2.123"
  "10.20.2.124"
)

# Camera credentials
USERNAME="<username>"
PASSWORD="<password>"

# Config files
CAPS_FILE="../Capabilities/SMTP.json"     # capability JSON (can be any)
EXCEL_FILE="../smtp_config.xlsx"          # Excel file (matches the JSON)

# Python binary and tool script
PYTHON_BIN="python"
CONFIG_TOOL_SCRIPT="../camera_config_tool_full.py"

# =====================================================
# Loop through each camera
# =====================================================
for IP in "${CAMERA_IPS[@]}"; do
  echo "-----------------------------------------"
  echo "Applying configuration to camera: $IP"
  echo "Using: $CAPS_FILE + $EXCEL_FILE"
  echo "-----------------------------------------"

  $PYTHON_BIN "$CONFIG_TOOL_SCRIPT" apply \
    --ip "$IP" \
    -u "$USERNAME" \
    -p "$PASSWORD" \
    --caps "$CAPS_FILE" \
    -i "$EXCEL_FILE"

  if [ $? -eq 0 ]; then
    echo "✅ Successfully applied config to $IP"
  else
    echo "❌ Failed to apply config to $IP"
  fi

  echo
done

echo "All done!"
