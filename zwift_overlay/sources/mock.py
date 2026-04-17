from __future__ import annotations

import math
import tkinter as tk
from datetime import datetime

from zwift_overlay.models import TelemetrySample
from zwift_overlay.sources.base import SampleCallback, TelemetrySource


class MockTelemetrySource(TelemetrySource):
    name = "demo"

    def __init__(self) -> None:
        self._callback: SampleCallback | None = None
        self._running = False
        self._tick = 0
        self._root: tk.Misc | None = None

    def start(self, callback: SampleCallback) -> None:
        if self._running:
            return
        self._callback = callback
        self._running = True
        self._root = tk._default_root
        self._schedule_next()

    def stop(self) -> None:
        self._running = False

    def _schedule_next(self) -> None:
        if not self._running or self._root is None:
            return

        self._tick += 1
        phase = self._tick / 8
        sample = TelemetrySample(
            timestamp=datetime.now(),
            heart_rate=int(138 + 10 * math.sin(phase)),
            speed_kph=34.0 + 2.5 * math.sin(phase / 2),
            power_watts=int(225 + 35 * math.sin(phase * 1.5)),
            cadence_rpm=int(88 + 6 * math.sin(phase * 1.2)),
        )
        if self._callback is not None:
            self._callback(sample)

        self._root.after(1000, self._schedule_next)
