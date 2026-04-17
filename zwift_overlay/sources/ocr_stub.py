from __future__ import annotations

from zwift_overlay.sources.base import SampleCallback, TelemetrySource


class OcrTelemetrySource(TelemetrySource):
    name = "ocr"

    def start(self, callback: SampleCallback) -> None:
        raise NotImplementedError(
            "OCR-källan är förberedd men inte inkopplad ännu. "
            "Nästa steg är att läsa ett valt område på skärmen och tolka siffrorna."
        )

    def stop(self) -> None:
        return
