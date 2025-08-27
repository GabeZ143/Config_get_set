# Camera Config Tool

A Python utility for automating **exporting**, **applying**, and **diffing** configuration values on IP cameras that expose CGI APIs.

## Overview

The tool is driven by a **capabilities JSON file**, which defines:

- Which API endpoints exist (e.g. `param.cgi`, `ptz.cgi`, `alarm.cgi`)
- Which `type=` sections are available (e.g. `AVStream`, `SMTP`, `motionAlarm`)
- Which keys are **gettable** and **settable**
- Optional per-section overrides (e.g. `setNorthPosition` requires `action=setNorthPosition` and no `type=`)

By editing this JSON, you can extend the tool to new devices or new endpoints without changing Python code.

## Features

- **Export** camera configs → Excel file in a **wide format**:
  - Each section (`AVStream`, `SMTP`, etc.) becomes a 3-column group (`Key`, `Value`, spacer).
  - Section headers in the first row, bold and underlined.
  - Rows below show key/value pairs; multiple instances (e.g. `streamID=0` and `streamID=1`) appear stacked in the same section.

- **Apply** configs ← Excel file:
  - Reads the wide format back and issues `action=set` CGI requests.
  - Only keys listed in `set_params` are sent (others skipped unless `--allow-unknown-keys`).

- **Diff** configs:
  - Compares live camera values to a baseline Excel.
  - Outputs differences in human-readable form.

- **Targets sheet** (optional):
  - Add a second sheet named `Targets` to your Excel to list which instances to fetch:

    ```text
    | Section     | cameraID | streamID | netCardId | IPProtoVer |
    |-------------|----------|----------|-----------|------------|
    | AVStream    | 1        | 0        |           |            |
    | AVStream    | 1        | 1        |           |            |
    | motionAlarm | 1        |          |           |            |
    | localNetwork|          |          | 1         | 1          |
    ```

  - Lets you avoid long `--selectors` arguments on the CLI.

## Usage

### Export

```bash
python camera_config_tool_full.py export   --ip 192.168.1.120 -u admin -p secret   --caps capabilities.json   --targets my_targets.xlsx   -o export.xlsx
```

### Apply

```bash
python camera_config_tool_full.py apply   --ip 192.168.1.120 -u admin -p secret   --caps capabilities.json   -i export.xlsx
```

### Diff

```bash
python camera_config_tool_full.py diff   --ip 192.168.1.120 -u admin -p secret   --caps capabilities.json   --targets my_targets.xlsx   -t baseline.xlsx
```

## Capabilities JSON

Two supported formats:

### Flat style

```json
{
  "AVStream": {
    "endpoint": "param.cgi",
    "get_params": ["cameraID","streamID","frameRate"],
    "set_params": ["cameraID","streamID","frameRate"]
  },
  "SMTP": {
    "endpoint": "param.cgi",
    "get_params": [],
    "set_params": ["server","port","user","password"]
  },
  "defaults": {
    "get": { "enabled": true, "action": "get", "requires_type": true },
    "set": { "enabled": true, "action": "set", "requires_type": true }
  }
}
```

### Nested style

```json
{
  "defaults": {
    "get": { "enabled": true, "action": "get", "requires_type": true },
    "set": { "enabled": true, "action": "set", "requires_type": true }
  },
  "sections": {
    "AVStream": {
      "endpoint": "param.cgi",
      "get_params": ["cameraID","streamID","frameRate"],
      "set_params": ["cameraID","streamID","frameRate"]
    },
    "SMTP": {
      "endpoint": "param.cgi",
      "get_params": [],
      "set_params": ["server","port","user","password"]
    }
  }
}
```

In both cases:
- `defaults` describe general GET/SET behavior.
- Individual sections can override defaults via a `modes` block (e.g. `setNorthPosition`).

## Notes

- Authentication (`userName`, `password`) is always sent as the **first two params**.
- Export skips sections where `get` mode is disabled.
- Apply skips keys not in `set_params` unless you pass `--allow-unknown-keys`.
- Works with any camera exposing `/cgi-bin/...` APIs — just update the capabilities JSON.

## Workflow Diagram

```
Capabilities JSON
        │
        ▼
   [ Export mode ]
        │
        ▼
   Excel (Wide format)
        │
   ┌────┴─────┐
   ▼          ▼
[ Apply ]   [ Diff ]
   │          │
   ▼          ▼
Camera   Baseline Excel
```
