# Zwift Overlay

Desktop app in Python that shows live training data in a compact overlay window for Zwift sessions.

## Current status

The app currently supports:

- Live power and heart rate from sensors (BLE and ANT+)
- W/kg calculation based on rider weight
- Rolling best power rows (preset windows + optional custom window)
- Session controls: Start, Delayed start, Pause, Reset, Stop
- Optional columns/rows in UI (session avg power, HR avg/max, speed avg, extra W/kg column)
- Sensor selection dialog with scan and assignment
- Auto reconnect checks for selected sensors after startup
- Contact window (Discord + Email flow)

## Run locally

From project root:

```powershell
py app.py
```

or launch without terminal:

```text
Start Zwift Overlay.vbs
```

Or download .exe file from Releases

## Sensor support

### Bluetooth (BLE)

Install:

```powershell
py -m pip install bleak
```

### ANT+

Install:

```powershell
py -m pip install openant
py -m pip install pyusb
```

## Settings and saved config

Config is saved automatically.

- Dev run (`py app.py`): `overlay_config.local.json` in project root
- EXE run (frozen build): `%APPDATA%\Zwift Overlay\overlay_config.json`

When running EXE, old config next to the EXE is migrated once if found.

## Build and releases (GitHub Actions)

Workflow file:

- `.github/workflows/release-windows.yml`

Behavior:

- Push to `main` -> build EXE and update release `Latest` (tag `latest`)
- Push to `stable` -> build EXE and update release `Stable` (tag `stable`)
- Same release is reused and updated (no new release per push)

The EXE is uploaded to GitHub Releases and should not live in source files.

## Changelog

Se [CHANGELOG.md](./CHANGELOG.md) för kortfattad historik över förbättringar, bugfixar och nya funktioner.

## Project structure

- `app.py` - app entrypoint
- `zwift_overlay/ui.py` - main UI and windows
- `zwift_overlay/config.py` - config model/load/save
- `zwift_overlay/sensors.py` - scan/discovery logic
- `zwift_overlay/sources/sensor_stub.py` - live telemetry source
- `zwift_overlay/stats.py` - aggregation and rolling averages
- `.github/workflows/release-windows.yml` - release build pipeline
- `.github/workflows/discord-on-push.yml` - Discord push notifications
