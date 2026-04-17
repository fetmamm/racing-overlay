from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path


CONFIG_PATH = Path(__file__).resolve().parent.parent / "overlay_config.json"


@dataclass(slots=True)
class SensorBinding:
    role: str
    name: str
    identifier: str
    transport: str


@dataclass(slots=True)
class AppConfig:
    rider_weight_kg: float = 100.0
    rider_weight_input: str = ""
    profile_name: str = ""
    profile_email: str = ""
    smtp_enabled: bool = False
    smtp_host: str = ""
    smtp_port: int = 587
    smtp_username: str = ""
    smtp_password: str = ""
    smtp_from_email: str = ""
    smtp_use_tls: bool = True
    always_on_top: bool = True
    power_display_seconds: int = 3
    wkg_decimals: int = 1
    delayed_start_seconds: int = 10
    ui_scale_percent: int = 100
    avg_power_windows_seconds: list[int] = field(default_factory=lambda: [300, 1200])
    custom_avg_power_seconds: int = 0
    show_custom_avg_power: bool = False
    show_session_avg_power: bool = True
    show_avg_hr: bool = True
    show_max_hr: bool = True
    show_avg_speed: bool = True
    show_adjusted_wkg_column: bool = False
    adjusted_wkg_percent: int = 90
    inactive_background: str = "#f3f3f3"
    sensors: dict[str, SensorBinding] = field(default_factory=dict)

    def get_sensor(self, role: str) -> SensorBinding | None:
        return self.sensors.get(role)

    def set_sensor(self, binding: SensorBinding) -> None:
        self.sensors[binding.role] = binding


def load_app_config() -> AppConfig:
    if not CONFIG_PATH.exists():
        return AppConfig()

    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception:
        return AppConfig()
    if not isinstance(data, dict):
        return AppConfig()

    def _safe_int(value: object, default: int) -> int:
        try:
            return int(value)
        except (TypeError, ValueError):
            return default

    def _safe_float(value: object, default: float) -> float:
        try:
            return float(value)
        except (TypeError, ValueError):
            return default

    try:
        adjusted_wkg_percent = int(data.get("adjusted_wkg_percent", 90))
    except (TypeError, ValueError):
        adjusted_wkg_percent = 90
    if adjusted_wkg_percent not in (90, 95):
        adjusted_wkg_percent = 90
    allowed_sensor_roles = {"power", "heart_rate"}
    sensors_payload = data.get("sensors", {})
    sensors: dict[str, SensorBinding] = {}
    if isinstance(sensors_payload, dict):
        for role, value in sensors_payload.items():
            if role not in allowed_sensor_roles or not isinstance(value, dict):
                continue
            name = value.get("name")
            identifier = value.get("identifier")
            transport = value.get("transport")
            if not all(isinstance(item, str) and item for item in (name, identifier, transport)):
                continue
            sensors[role] = SensorBinding(
                role=role,
                name=name,
                identifier=identifier,
                transport=transport,
            )

    avg_windows_payload = data.get("avg_power_windows_seconds", [300, 1200])
    avg_windows: list[int] = []
    if isinstance(avg_windows_payload, list):
        for raw_value in avg_windows_payload:
            value = _safe_int(raw_value, 0)
            if value > 0:
                avg_windows.append(value)
    if not avg_windows:
        avg_windows = [300, 1200]

    return AppConfig(
        rider_weight_kg=_safe_float(data.get("rider_weight_kg", 100.0), 100.0),
        rider_weight_input=str(data.get("rider_weight_input", "")),
        profile_name=str(data.get("profile_name", "")).strip(),
        profile_email=str(data.get("profile_email", "")).strip(),
        smtp_enabled=bool(data.get("smtp_enabled", False)),
        smtp_host=str(data.get("smtp_host", "")).strip(),
        smtp_port=max(1, _safe_int(data.get("smtp_port", 587), 587)),
        smtp_username=str(data.get("smtp_username", "")).strip(),
        smtp_password=str(data.get("smtp_password", "")),
        smtp_from_email=str(data.get("smtp_from_email", "")).strip(),
        smtp_use_tls=bool(data.get("smtp_use_tls", True)),
        always_on_top=bool(data.get("always_on_top", True)),
        power_display_seconds=max(1, _safe_int(data.get("power_display_seconds", 3), 3)),
        wkg_decimals=1 if _safe_int(data.get("wkg_decimals", 1), 1) not in (1, 2) else _safe_int(data.get("wkg_decimals", 1), 1),
        delayed_start_seconds=max(1, _safe_int(data.get("delayed_start_seconds", 10), 10)),
        ui_scale_percent=max(50, min(200, _safe_int(data.get("ui_scale_percent", 100), 100))),
        avg_power_windows_seconds=avg_windows,
        custom_avg_power_seconds=max(0, _safe_int(data.get("custom_avg_power_seconds", 0), 0)),
        show_custom_avg_power=bool(data.get("show_custom_avg_power", False)),
        show_session_avg_power=bool(data.get("show_session_avg_power", True)),
        show_avg_hr=bool(data.get("show_avg_hr", True)),
        show_max_hr=bool(data.get("show_max_hr", True)),
        show_avg_speed=bool(data.get("show_avg_speed", True)),
        show_adjusted_wkg_column=bool(data.get("show_adjusted_wkg_column", False)),
        adjusted_wkg_percent=adjusted_wkg_percent,
        inactive_background=str(data.get("inactive_background", "#f3f3f3")),
        sensors=sensors,
    )


def save_app_config(config: AppConfig) -> None:
    payload = {
        "rider_weight_kg": config.rider_weight_kg,
        "rider_weight_input": config.rider_weight_input,
        "profile_name": config.profile_name,
        "profile_email": config.profile_email,
        "smtp_enabled": config.smtp_enabled,
        "smtp_host": config.smtp_host,
        "smtp_port": config.smtp_port,
        "smtp_username": config.smtp_username,
        "smtp_password": config.smtp_password,
        "smtp_from_email": config.smtp_from_email,
        "smtp_use_tls": config.smtp_use_tls,
        "always_on_top": config.always_on_top,
        "power_display_seconds": config.power_display_seconds,
        "wkg_decimals": config.wkg_decimals,
        "delayed_start_seconds": config.delayed_start_seconds,
        "ui_scale_percent": config.ui_scale_percent,
        "avg_power_windows_seconds": config.avg_power_windows_seconds,
        "custom_avg_power_seconds": config.custom_avg_power_seconds,
        "show_custom_avg_power": config.show_custom_avg_power,
        "show_session_avg_power": config.show_session_avg_power,
        "show_avg_hr": config.show_avg_hr,
        "show_max_hr": config.show_max_hr,
        "show_avg_speed": config.show_avg_speed,
        "show_adjusted_wkg_column": config.show_adjusted_wkg_column,
        "adjusted_wkg_percent": config.adjusted_wkg_percent,
        "inactive_background": config.inactive_background,
        "sensors": {
            role: {
                "name": binding.name,
                "identifier": binding.identifier,
                "transport": binding.transport,
            }
            for role, binding in config.sensors.items()
        },
    }
    CONFIG_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")
