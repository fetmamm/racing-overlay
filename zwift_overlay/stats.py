from __future__ import annotations

from collections import deque
from datetime import datetime, timedelta

from zwift_overlay.models import SummaryStats, TelemetrySample


class TelemetryAggregator:
    def __init__(self) -> None:
        self.samples: list[TelemetrySample] = []
        self._power_5m: deque[TelemetrySample] = deque()
        self._power_20m: deque[TelemetrySample] = deque()
    def add_sample(self, sample: TelemetrySample) -> SummaryStats:
        self.samples.append(sample)
        self._append_rolling_sample(self._power_5m, sample, timedelta(minutes=5))
        self._append_rolling_sample(self._power_20m, sample, timedelta(minutes=20))
        return self.summary()

    def summary(self) -> SummaryStats:
        latest = self.samples[-1] if self.samples else None
        elapsed_seconds = 0
        if self.samples:
            elapsed = self.samples[-1].timestamp - self.samples[0].timestamp
            elapsed_seconds = max(0, int(elapsed.total_seconds()))
        return SummaryStats(
            current_heart_rate=latest.heart_rate if latest else None,
            current_speed_kph=latest.speed_kph if latest else None,
            current_power_watts=latest.power_watts if latest else None,
            current_cadence_rpm=latest.cadence_rpm if latest else None,
            elapsed_seconds=elapsed_seconds,
            average_heart_rate=self._average("heart_rate"),
            max_heart_rate=self._max("heart_rate"),
            average_power_watts=self._average("power_watts"),
            rolling_power_5m=self._deque_average(self._power_5m, "power_watts"),
            rolling_power_20m=self._deque_average(self._power_20m, "power_watts"),
            average_cadence_rpm=self._average("cadence_rpm"),
            average_speed_kph=self._average("speed_kph"),
            sample_count=len(self.samples),
        )

    def rolling_average(self, attribute: str, seconds: int) -> float | None:
        if seconds <= 0 or not self.samples:
            return None
        latest_timestamp = self.samples[-1].timestamp
        cutoff = latest_timestamp - timedelta(seconds=seconds)
        values: list[float] = []
        for sample in reversed(self.samples):
            if sample.timestamp < cutoff:
                break
            value = getattr(sample, attribute)
            if value is not None:
                values.append(float(value))
        if not values:
            return None
        return sum(values) / len(values)

    def _append_rolling_sample(
        self,
        bucket: deque[TelemetrySample],
        sample: TelemetrySample,
        window: timedelta,
    ) -> None:
        bucket.append(sample)
        cutoff = sample.timestamp - window
        while bucket and bucket[0].timestamp < cutoff:
            bucket.popleft()

    def _average(self, attribute: str) -> float | None:
        values = [getattr(sample, attribute) for sample in self.samples]
        filtered = [value for value in values if value is not None]
        if not filtered:
            return None
        return sum(filtered) / len(filtered)

    def _deque_average(
        self,
        bucket: deque[TelemetrySample],
        attribute: str,
    ) -> float | None:
        filtered = [
            getattr(sample, attribute) for sample in bucket if getattr(sample, attribute) is not None
        ]
        if not filtered:
            return None
        return sum(filtered) / len(filtered)

    def _max(self, attribute: str) -> int | None:
        values = [getattr(sample, attribute) for sample in self.samples]
        filtered = [value for value in values if value is not None]
        if not filtered:
            return None
        return max(filtered)


def create_sample(
    heart_rate: int | None,
    speed_kph: float | None,
    power_watts: int | None,
    cadence_rpm: int | None,
    timestamp: datetime | None = None,
) -> TelemetrySample:
    return TelemetrySample(
        timestamp=timestamp or datetime.now(),
        heart_rate=heart_rate,
        speed_kph=speed_kph,
        power_watts=power_watts,
        cadence_rpm=cadence_rpm,
    )
