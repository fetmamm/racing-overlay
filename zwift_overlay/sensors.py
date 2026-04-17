from __future__ import annotations

import asyncio
from dataclasses import dataclass
import sys
import threading
import time


SENSOR_ROLES = {
    "power": "Power Meter",
    "heart_rate": "Heart Rate Monitor",
}

TRANSPORTS = {
    "ble": "Bluetooth LE",
    "ant": "ANT+",
}

BLE_SERVICE_TYPES = {
    "180d": "Heart Rate",
    "1816": "Cadence/Speed",
    "1818": "Power",
    "1826": "Fitness Machine",
}

BLE_NAME_HINTS = {
    "heart": "Heart Rate",
    "hr": "Heart Rate",
    "tickr": "Heart Rate",
    "h10": "Heart Rate",
    "pulse": "Heart Rate",
    "power": "Power",
    "assioma": "Power",
    "favero": "Power",
    "stages": "Power",
    "4iiii": "Power",
    "rally": "Power",
    "vector": "Power",
    "watt": "Power",
    "cad": "Cadence",
    "cadence": "Cadence",
    "rpm": "Cadence",
    "magene": "Cadence",
    "garmin": "Sensor",
    "wahoo": "Sensor",
}

KNOWN_ANT_USB_IDS = {
    (0x0FCF, 0x1008): "Garmin ANT+ USB-mottagare",
    (0x0FCF, 0x1009): "Dynastream ANTUSB-m",
    (0x0FCF, 0x1004): "Dynastream ANT2USB",
    (0x0FCF, 0x1003): "Dynastream ANT USB stick",
    (0x0FCF, 0x1006): "Dynastream ANT USB2 stick",
}

BLE_SCAN_SECONDS = 12.0
ANT_DONGLE_DETECT_TIMEOUT_SECONDS = 6.0
ANT_SENSOR_SCAN_SECONDS = 10.0


@dataclass(slots=True)
class DiscoveredSensor:
    name: str
    identifier: str
    transport: str
    sensor_type: str = "Unknown"
    details: str = ""
    native_device: object | None = None


class SensorDiscoveryError(RuntimeError):
    pass


class SensorScanCancelledError(SensorDiscoveryError):
    pass


class SensorDiscoveryService:
    def __init__(self) -> None:
        self._availability_cache: dict[str, tuple[float, tuple[bool, str]]] = {}

    def check_transport_available(self, transport: str, force_refresh: bool = False) -> tuple[bool, str]:
        now = time.monotonic()
        cache_entry = self._availability_cache.get(transport)
        ttl_seconds = 4.0 if transport == "ble" else 1.5
        if not force_refresh and cache_entry is not None:
            cached_at, cached_result = cache_entry
            if now - cached_at <= ttl_seconds:
                return cached_result

        if transport == "ble":
            result = self._check_ble_available()
        elif transport == "ant":
            result = self._check_ant_available()
        else:
            result = (False, f"Unknown transport: {transport}")

        self._availability_cache[transport] = (now, result)
        return result

    def scan(
        self,
        transport: str,
        stop_event: threading.Event | None = None,
        scan_seconds: float | None = None,
        allow_empty: bool = False,
    ) -> list[DiscoveredSensor]:
        if transport == "ble":
            return self._scan_ble(stop_event, scan_seconds=scan_seconds, allow_empty=allow_empty)
        if transport == "ant":
            return self._scan_ant(stop_event, scan_seconds=scan_seconds, allow_empty=allow_empty)
        raise SensorDiscoveryError(f"Unknown transport: {transport}")

    def _check_ble_available(self) -> tuple[bool, str]:
        self._prepare_windows_ble_environment()
        try:
            from bleak import BleakScanner
        except ImportError:
            return False, "BLE unavailable: missing 'bleak' package."

        try:
            asyncio.run(BleakScanner.discover(timeout=0.35, return_adv=False))
        except Exception as exc:
            return False, f"BLE unavailable: {exc}"
        return True, "BLE available"

    def _check_ant_available(self) -> tuple[bool, str]:
        try:
            import openant  # noqa: F401
        except ImportError:
            return False, "ANT+ unavailable: missing 'openant' package."

        try:
            dongles = self._detect_ant_dongles_with_retry(
                timeout_seconds=ANT_DONGLE_DETECT_TIMEOUT_SECONDS,
                poll_interval=0.4,
            )
        except SensorDiscoveryError as exc:
            return False, f"ANT+ unavailable: {exc}"

        if dongles:
            return True, "ANT+ dongle detected"

        if self._probe_ant_node_access():
            return True, "ANT+ adapter reachable"

        return False, "ANT+ unavailable: no ANT+ USB dongle detected."

    @staticmethod
    def _probe_ant_node_access() -> bool:
        try:
            from openant.devices import ANTPLUS_NETWORK_KEY
            from openant.easy.node import Node
        except Exception:
            return False

        node = None
        try:
            node = Node()
            node.set_network_key(0x00, ANTPLUS_NETWORK_KEY)
            return True
        except Exception:
            return False
        finally:
            if node is not None:
                try:
                    node.stop()
                except Exception:
                    pass

    def _scan_ble(
        self,
        stop_event: threading.Event | None = None,
        scan_seconds: float | None = None,
        allow_empty: bool = False,
    ) -> list[DiscoveredSensor]:
        self._prepare_windows_ble_environment()
        try:
            from bleak import BleakScanner
        except ImportError as exc:
            raise SensorDiscoveryError(
                "Bluetooth scanning requires the 'bleak' package. Install with: py -m pip install bleak"
            ) from exc

        try:
            effective_scan_seconds = BLE_SCAN_SECONDS if scan_seconds is None else max(0.4, float(scan_seconds))
            discovered = asyncio.run(
                self._discover_ble_devices(
                    BleakScanner,
                    stop_event,
                    scan_seconds=effective_scan_seconds,
                )
            )
        except SensorScanCancelledError:
            raise
        except Exception as exc:  # pragma: no cover - beror pa lokal Bluetooth-stack
            raise SensorDiscoveryError(f"Bluetooth scan failed: {exc}") from exc

        discovered.sort(key=lambda item: (item.sensor_type == "Unknown", item.name.lower(), item.identifier))
        if discovered:
            return discovered

        if stop_event is not None and stop_event.is_set():
            raise SensorScanCancelledError("Search stopped.")
        if allow_empty:
            return []

        raise SensorDiscoveryError(
            "No BLE devices were found. Make sure the sensor is awake, "
            "Bluetooth is enabled in Windows, and the sensor is not already locked to Zwift."
        )

    @staticmethod
    def _prepare_windows_ble_environment() -> None:
        if not sys.platform.startswith("win"):
            return

        try:
            from bleak.backends.winrt.util import uninitialize_sta
        except Exception:
            return

        try:
            uninitialize_sta()
        except Exception:
            return

    async def _discover_ble_devices(
        self,
        scanner_class: type,
        stop_event: threading.Event | None = None,
        scan_seconds: float = BLE_SCAN_SECONDS,
    ) -> list[DiscoveredSensor]:
        devices: dict[str, DiscoveredSensor] = {}
        scanner = scanner_class()
        await scanner.start()
        started_at = time.monotonic()
        discovered: dict[str, tuple[object, object]] = {}
        try:
            while time.monotonic() - started_at < scan_seconds:
                if stop_event is not None and stop_event.is_set():
                    raise SensorScanCancelledError("Search stopped.")
                await asyncio.sleep(0.2)
            discovered = getattr(scanner, "discovered_devices_and_advertisement_data", {}) or {}
        finally:
            await scanner.stop()

        for identifier, pair in discovered.items():
            device, advertisement_data = pair
            resolved_identifier = getattr(device, "address", None) or identifier
            name = getattr(device, "name", None) or getattr(advertisement_data, "local_name", None)
            if not name:
                name = "Unknown BLE device"

            sensor_type = self._classify_ble_device(name, advertisement_data)
            details = self._describe_ble_device(advertisement_data)
            devices[resolved_identifier] = DiscoveredSensor(
                name=name,
                identifier=resolved_identifier,
                transport="ble",
                sensor_type=sensor_type,
                details=details,
                native_device=device,
            )

        return list(devices.values())

    def _classify_ble_device(self, name: str, advertisement_data: object) -> str:
        uuids = getattr(advertisement_data, "service_uuids", None) or []
        normalized = [self._short_ble_uuid(uuid) for uuid in uuids]
        for short_uuid in normalized:
            if short_uuid in BLE_SERVICE_TYPES:
                return BLE_SERVICE_TYPES[short_uuid]

        lower_name = name.lower()
        for hint, sensor_type in BLE_NAME_HINTS.items():
            if hint in lower_name:
                return sensor_type

        manufacturer_data = getattr(advertisement_data, "manufacturer_data", None) or {}
        if manufacturer_data:
            return "Unknown sensor"
        return "Unknown"

    def _describe_ble_device(self, advertisement_data: object) -> str:
        service_uuids = getattr(advertisement_data, "service_uuids", None) or []
        rssi = getattr(advertisement_data, "rssi", None)

        type_parts = []
        for uuid in service_uuids:
            label = BLE_SERVICE_TYPES.get(self._short_ble_uuid(uuid))
            if label and label not in type_parts:
                type_parts.append(label)

        parts: list[str] = []
        if type_parts:
            parts.append("/".join(type_parts))
        if rssi is not None:
            parts.append(f"RSSI {rssi}")
        if not parts:
            return "BLE advertisement detected"
        return " | ".join(parts)

    @staticmethod
    def _short_ble_uuid(uuid: str) -> str:
        normalized = uuid.lower().replace("-", "")
        if normalized.startswith("0000") and len(normalized) >= 8:
            return normalized[4:8]
        return normalized[:4]

    def _scan_ant(
        self,
        stop_event: threading.Event | None = None,
        scan_seconds: float | None = None,
        allow_empty: bool = False,
    ) -> list[DiscoveredSensor]:
        if stop_event is not None and stop_event.is_set():
            raise SensorScanCancelledError("Search stopped.")
        try:
            import openant  # noqa: F401
        except ImportError as exc:
            raise SensorDiscoveryError(
                "ANT+ scanning requires the 'openant' package and an ANT+ dongle. "
                "Install with: py -m pip install openant"
            ) from exc

        dongles = self._detect_ant_dongles_with_retry(
            timeout_seconds=ANT_DONGLE_DETECT_TIMEOUT_SECONDS,
            poll_interval=0.4,
            stop_event=stop_event,
        )
        if stop_event is not None and stop_event.is_set():
            raise SensorScanCancelledError("Search stopped.")
        if not dongles and not self._probe_ant_node_access():
            raise SensorDiscoveryError(
                "No ANT+ dongle was found in Windows. Check that the dongle is plugged in "
                "and not already in use by another application."
            )

        effective_scan_seconds = ANT_SENSOR_SCAN_SECONDS if scan_seconds is None else max(0.6, float(scan_seconds))
        sensors = self._scan_ant_plus_sensors(stop_event=stop_event, scan_seconds=effective_scan_seconds)
        if stop_event is not None and stop_event.is_set():
            raise SensorScanCancelledError("Search stopped.")
        if sensors:
            return sensors
        if allow_empty:
            return []

        raise SensorDiscoveryError(
            "ANT+ dongle found, but no ANT+ sensors were discovered. "
            "Wake up the sensor and make sure it is not already connected to another app."
        )

    def _scan_ant_plus_sensors(
        self,
        stop_event: threading.Event | None = None,
        scan_seconds: float = 3.0,
    ) -> list[DiscoveredSensor]:
        try:
            from openant.devices import ANTPLUS_NETWORK_KEY
            from openant.devices.common import DeviceType
            from openant.devices.scanner import Scanner
            from openant.easy.node import Node
        except ImportError as exc:
            raise SensorDiscoveryError(
                "ANT+ scanning requires openant scanner modules. Reinstall with: py -m pip install --upgrade openant"
            ) from exc

        found_by_id: dict[str, DiscoveredSensor] = {}
        worker_error: list[Exception] = []
        node: object | None = None
        scanner: object | None = None

        def _on_found(device_tuple: tuple[int, int, int]) -> None:
            try:
                device_id, device_type, transmission_type = device_tuple
                device_type_name = DeviceType(device_type).name
            except Exception:
                device_id, device_type, transmission_type = device_tuple
                device_type_name = "Unknown"

            sensor_type = self._ant_sensor_type_label(device_type_name)
            identifier = f"ANT:{device_id}:{device_type}:{transmission_type}"
            display_name = f"ANT+ {sensor_type} #{device_id}"
            details = f"Device ID {device_id} | Type {device_type_name} ({device_type}) | Tx {transmission_type}"
            found_by_id[identifier] = DiscoveredSensor(
                name=display_name,
                identifier=identifier,
                transport="ant",
                sensor_type=sensor_type,
                details=details,
            )

        def _worker() -> None:
            nonlocal node, scanner
            try:
                node = Node()
                node.set_network_key(0x00, ANTPLUS_NETWORK_KEY)
                scanner = Scanner(node, device_id=0, device_type=0)
                scanner.on_found = _on_found
                node.start()
            except Exception as exc:  # pragma: no cover - depends on local ANT stack
                worker_error.append(exc)

        thread = threading.Thread(target=_worker, daemon=True)
        thread.start()

        started = time.monotonic()
        while thread.is_alive() and time.monotonic() - started < scan_seconds:
            if stop_event is not None and stop_event.is_set():
                break
            time.sleep(0.1)

        try:
            if scanner is not None:
                scanner.close_channel()
        except Exception:
            pass
        try:
            if node is not None:
                node.stop()
        except Exception:
            pass
        thread.join(timeout=1.5)

        if stop_event is not None and stop_event.is_set():
            raise SensorScanCancelledError("Search stopped.")
        if worker_error:
            raise SensorDiscoveryError(f"ANT+ scan failed: {worker_error[0]}")

        sensors = list(found_by_id.values())
        sensors.sort(key=lambda item: (item.sensor_type == "Unknown", item.name.lower(), item.identifier))
        return sensors

    @staticmethod
    def _ant_sensor_type_label(device_type_name: str) -> str:
        mapping = {
            "PowerMeter": "Power",
            "HeartRate": "Heart Rate",
            "BikeCadence": "Cadence",
            "BikeSpeedCadence": "Cadence/Speed",
            "BikeSpeed": "Speed",
            "FitnessEquipment": "Fitness Machine",
        }
        return mapping.get(device_type_name, device_type_name or "Unknown")

    def _detect_ant_dongles(self, stop_event: threading.Event | None = None) -> list[DiscoveredSensor]:
        try:
            import usb.core
        except ImportError as exc:
            raise SensorDiscoveryError(
                "ANT+ scanning found openant but is missing 'pyusb' for USB detection. "
                "Install with: py -m pip install pyusb"
            ) from exc

        dongles: list[DiscoveredSensor] = []
        for (vendor_id, product_id), label in KNOWN_ANT_USB_IDS.items():
            if stop_event is not None and stop_event.is_set():
                raise SensorScanCancelledError("Search stopped.")
            matches = usb.core.find(find_all=True, idVendor=vendor_id, idProduct=product_id)
            for index, _device in enumerate(matches):
                dongles.append(
                    DiscoveredSensor(
                        name=label,
                        identifier=f"USB {vendor_id:04X}:{product_id:04X} #{index + 1}",
                        transport="ant",
                        sensor_type="ANT+ dongle",
                        details="Dongle detected. ANT+ sensor scanning is the next step.",
                    )
                )
        return dongles

    def _detect_ant_dongles_with_retry(
        self,
        timeout_seconds: float,
        poll_interval: float,
        stop_event: threading.Event | None = None,
    ) -> list[DiscoveredSensor]:
        started = time.monotonic()
        while True:
            if stop_event is not None and stop_event.is_set():
                raise SensorScanCancelledError("Search stopped.")
            dongles = self._detect_ant_dongles(stop_event=stop_event)
            if dongles:
                return dongles
            if time.monotonic() - started >= timeout_seconds:
                return []
            time.sleep(max(0.1, poll_interval))
