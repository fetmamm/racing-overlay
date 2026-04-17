from __future__ import annotations

from zwift_overlay.config import AppConfig
from zwift_overlay.sources.base import TelemetrySource
from zwift_overlay.sources.sensor_stub import SensorTelemetrySource


def create_telemetry_source(config: AppConfig) -> TelemetrySource:
    return SensorTelemetrySource(config)
