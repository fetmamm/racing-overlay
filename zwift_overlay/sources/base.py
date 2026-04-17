from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Callable

from zwift_overlay.models import TelemetrySample


SampleCallback = Callable[[TelemetrySample], None]


class TelemetrySource(ABC):
    name = "okänd källa"

    @abstractmethod
    def start(self, callback: SampleCallback) -> None:
        raise NotImplementedError

    @abstractmethod
    def stop(self) -> None:
        raise NotImplementedError
