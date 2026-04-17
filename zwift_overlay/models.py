from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(slots=True)
class TelemetrySample:
    timestamp: datetime
    heart_rate: int | None = None
    speed_kph: float | None = None
    power_watts: int | None = None
    cadence_rpm: int | None = None


@dataclass(slots=True)
class SummaryStats:
    current_heart_rate: int | None
    current_speed_kph: float | None
    current_power_watts: int | None
    current_cadence_rpm: int | None
    elapsed_seconds: int
    average_heart_rate: float | None
    max_heart_rate: int | None
    average_power_watts: float | None
    rolling_power_5m: float | None
    rolling_power_20m: float | None
    average_cadence_rpm: float | None
    average_speed_kph: float | None
    sample_count: int
