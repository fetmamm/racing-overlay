from __future__ import annotations

import asyncio
from dataclasses import dataclass
from datetime import datetime
import threading
import time
from typing import Any

from zwift_overlay.config import AppConfig, SensorBinding
from zwift_overlay.models import TelemetrySample
from zwift_overlay.sources.base import SampleCallback, TelemetrySource


HEART_RATE_MEASUREMENT_UUID = "00002a37-0000-1000-8000-00805f9b34fb"
CYCLING_POWER_MEASUREMENT_UUID = "00002a63-0000-1000-8000-00805f9b34fb"
CSC_MEASUREMENT_UUID = "00002a5b-0000-1000-8000-00805f9b34fb"


@dataclass(slots=True)
class CrankState:
    cumulative_revs: int | None = None
    event_time: int | None = None


class SensorTelemetrySource(TelemetrySource):
    name = "sensor"

    def __init__(self, config: AppConfig) -> None:
        self.config = config
        self._callback: SampleCallback | None = None
        self._stop_event = threading.Event()
        self._latest_lock = threading.Lock()

        self._ble_thread: threading.Thread | None = None
        self._ble_loop: asyncio.AbstractEventLoop | None = None

        self._ant_thread: threading.Thread | None = None
        self._ant_node: Any | None = None
        self._ant_devices: list[Any] = []

        self._latest_values: dict[str, int | float | None] = {
            "heart_rate": None,
            "power_watts": None,
            "cadence_rpm": None,
        }
        self._last_update_monotonic: dict[str, float] = {
            "heart_rate": 0.0,
            "power_watts": 0.0,
            "cadence_rpm": 0.0,
        }
        self._power_crank_state = CrankState()
        self._csc_crank_state = CrankState()

    def start(self, callback: SampleCallback) -> None:
        if not self.config.sensors:
            # Allow session start without sensors (for timer/speed-only workflows).
            self._callback = callback
            self._stop_event.clear()
            return

        ble_bindings = [binding for binding in self.config.sensors.values() if binding.transport == "ble"]
        ant_bindings = [binding for binding in self.config.sensors.values() if binding.transport == "ant"]

        if not ble_bindings and not ant_bindings:
            # Nothing supported to start right now, but do not block session start.
            self._callback = callback
            self._stop_event.clear()
            return

        if self._ble_thread is not None and not self._ble_thread.is_alive():
            self._ble_thread = None
        if self._ant_thread is not None and not self._ant_thread.is_alive():
            self._ant_thread = None

        start_ble = bool(ble_bindings) and (self._ble_thread is None or not self._ble_thread.is_alive())
        start_ant = bool(ant_bindings) and (self._ant_thread is None or not self._ant_thread.is_alive())

        self._callback = callback
        if start_ble or start_ant:
            self._stop_event.clear()
        if any(value is not None for value in self._latest_values.values()):
            self._emit_sample()

        if start_ble:
            self._ble_thread = threading.Thread(
                target=self._ble_thread_main,
                args=(ble_bindings,),
                daemon=True,
            )
            self._ble_thread.start()

        if start_ant:
            self._ant_thread = threading.Thread(
                target=self._ant_thread_main,
                args=(ant_bindings,),
                daemon=True,
            )
            self._ant_thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._ble_loop is not None:
            self._ble_loop.call_soon_threadsafe(lambda: None)
        if self._ant_node is not None:
            try:
                self._ant_node.stop()
            except Exception:
                pass

        if self._ble_thread is not None and self._ble_thread.is_alive():
            self._ble_thread.join(timeout=0.8)
            if not self._ble_thread.is_alive():
                self._ble_thread = None
        else:
            self._ble_thread = None
        if self._ant_thread is not None and self._ant_thread.is_alive():
            self._ant_thread.join(timeout=0.8)
            if not self._ant_thread.is_alive():
                self._ant_thread = None
        else:
            self._ant_thread = None

        if self._ble_thread is None:
            self._ble_loop = None
        if self._ant_thread is None:
            self._ant_node = None
            self._ant_devices = []

    def _ble_thread_main(self, bindings: list[SensorBinding]) -> None:
        try:
            asyncio.run(self._async_ble_main(bindings))
        except Exception:
            return

    async def _async_ble_main(self, bindings: list[SensorBinding]) -> None:
        self._ble_loop = asyncio.get_running_loop()
        grouped: dict[str, list[SensorBinding]] = {}
        for binding in bindings:
            grouped.setdefault(binding.identifier, []).append(binding)

        tasks = [asyncio.create_task(self._run_ble_device(identifier, device_bindings)) for identifier, device_bindings in grouped.items()]
        try:
            while not self._stop_event.is_set():
                if tasks and all(task.done() for task in tasks):
                    break
                await asyncio.sleep(0.2)
        finally:
            for task in tasks:
                task.cancel()
            await asyncio.gather(*tasks, return_exceptions=True)

    async def _run_ble_device(self, identifier: str, bindings: list[SensorBinding]) -> None:
        try:
            from bleak import BleakClient
        except ImportError as exc:
            raise NotImplementedError(
                "Bluetooth live reading requires the 'bleak' package. Install with: py -m pip install bleak"
            ) from exc

        roles = {binding.role for binding in bindings}
        monitored_keys = self._monitored_value_keys_for_roles(roles)
        while not self._stop_event.is_set():
            try:
                async with BleakClient(identifier) as client:
                    available = await self._subscribe_for_ble_bindings(client, bindings)
                    connection_started_at = time.monotonic()
                    hr_fallback_task: asyncio.Task[None] | None = None
                    power_fallback_task: asyncio.Task[None] | None = None
                    if "heart_rate" in roles and HEART_RATE_MEASUREMENT_UUID in available:
                        hr_fallback_task = asyncio.create_task(self._ble_hr_fallback_loop(client))
                    if "power" in roles and CYCLING_POWER_MEASUREMENT_UUID in available:
                        power_fallback_task = asyncio.create_task(self._ble_power_fallback_loop(client))
                    try:
                        while not self._stop_event.is_set():
                            for value_key in monitored_keys:
                                if self._is_stream_stale(
                                    key=value_key,
                                    connected_at=connection_started_at,
                                    stale_seconds=10.0,
                                    bootstrap_seconds=15.0,
                                ):
                                    raise RuntimeError(f"{value_key} stream became stale; reconnecting.")
                            await asyncio.sleep(0.5)
                    finally:
                        if hr_fallback_task is not None:
                            hr_fallback_task.cancel()
                            await asyncio.gather(hr_fallback_task, return_exceptions=True)
                        if power_fallback_task is not None:
                            power_fallback_task.cancel()
                            await asyncio.gather(power_fallback_task, return_exceptions=True)
            except asyncio.CancelledError:
                raise
            except NotImplementedError:
                return
            except Exception:
                if self._stop_event.is_set():
                    return
                await asyncio.sleep(1.0)

    async def _subscribe_for_ble_bindings(self, client: object, bindings: list[SensorBinding]) -> set[str]:
        roles = {binding.role for binding in bindings}
        available = {service.uuid.lower() for service in client.services} | {
            characteristic.uuid.lower()
            for service in client.services
            for characteristic in service.characteristics
        }

        if "heart_rate" in roles and HEART_RATE_MEASUREMENT_UUID in available:
            await client.start_notify(HEART_RATE_MEASUREMENT_UUID, self._handle_ble_heart_rate)

        needs_power_stream = "power" in roles or ("cadence" in roles and CYCLING_POWER_MEASUREMENT_UUID in available)
        if needs_power_stream and CYCLING_POWER_MEASUREMENT_UUID in available:
            allow_cadence_from_power = "cadence" in roles
            await client.start_notify(
                CYCLING_POWER_MEASUREMENT_UUID,
                lambda sender, data: self._handle_ble_power(sender, data, allow_cadence_from_power),
            )

        if "cadence" in roles and CSC_MEASUREMENT_UUID in available:
            await client.start_notify(CSC_MEASUREMENT_UUID, self._handle_ble_cadence)

        if "heart_rate" in roles and HEART_RATE_MEASUREMENT_UUID not in available:
            raise NotImplementedError("The selected heart rate sensor does not expose Heart Rate Measurement.")
        if "power" in roles and CYCLING_POWER_MEASUREMENT_UUID not in available:
            raise NotImplementedError("The selected power meter does not expose Cycling Power Measurement.")
        if "cadence" in roles and CSC_MEASUREMENT_UUID not in available and CYCLING_POWER_MEASUREMENT_UUID not in available:
            raise NotImplementedError("The selected cadence sensor does not expose cadence data via CSC or Cycling Power.")
        return available

    def _ant_thread_main(self, bindings: list[SensorBinding]) -> None:
        try:
            from openant.devices import ANTPLUS_NETWORK_KEY
            from openant.devices.bike_speed_cadence import BikeCadence, BikeSpeedCadence
            from openant.devices.common import DeviceType
            from openant.devices.heart_rate import HeartRate
            from openant.devices.power_meter import PowerMeter
            from openant.easy.node import Node
        except ImportError as exc:
            raise NotImplementedError(
                "ANT+ live reading requires the 'openant' package. Install with: py -m pip install openant"
            ) from exc

        grouped: dict[str, list[SensorBinding]] = {}
        for binding in bindings:
            grouped.setdefault(binding.identifier, []).append(binding)
        all_roles = {binding.role for binding in bindings}
        monitored_keys = self._monitored_value_keys_for_roles(all_roles)

        while not self._stop_event.is_set():
            node: Any | None = None
            devices: list[Any] = []
            watchdog_thread: threading.Thread | None = None
            connection_started_at = time.monotonic()
            try:
                node = Node()
                self._ant_node = node
                node.set_network_key(0x00, ANTPLUS_NETWORK_KEY)

                for identifier, device_bindings in grouped.items():
                    parsed = self._parse_ant_identifier(identifier)
                    if parsed is None:
                        raise NotImplementedError(
                            f"Invalid ANT+ identifier '{identifier}'. Re-scan and reassign the ANT+ sensor."
                        )
                    device_id, device_type, transmission_type = parsed
                    roles = {binding.role for binding in device_bindings}
                    device = self._create_ant_device(
                        node=node,
                        device_id=device_id,
                        device_type=device_type,
                        transmission_type=transmission_type,
                        roles=roles,
                        device_type_enum=DeviceType,
                        heart_rate_cls=HeartRate,
                        power_meter_cls=PowerMeter,
                        bike_cadence_cls=BikeCadence,
                        bike_speed_cadence_cls=BikeSpeedCadence,
                    )
                    device.on_device_data = (
                        lambda _page, page_name, data, roles=roles: self._handle_ant_device_data(roles, page_name, data)
                    )
                    devices.append(device)

                self._ant_devices = devices

                def _watchdog() -> None:
                    while not self._stop_event.is_set():
                        for value_key in monitored_keys:
                            if self._is_stream_stale(
                                key=value_key,
                                connected_at=connection_started_at,
                                stale_seconds=12.0,
                                bootstrap_seconds=18.0,
                            ):
                                try:
                                    node.stop()
                                except Exception:
                                    pass
                                return
                        time.sleep(1.0)

                watchdog_thread = threading.Thread(target=_watchdog, daemon=True)
                watchdog_thread.start()
                node.start()
            except Exception:
                if self._stop_event.is_set():
                    return
            finally:
                for device in devices:
                    try:
                        device.close_channel()
                    except Exception:
                        pass
                if node is not None:
                    try:
                        node.stop()
                    except Exception:
                        pass
                if watchdog_thread is not None and watchdog_thread.is_alive():
                    watchdog_thread.join(timeout=0.2)
                self._ant_node = None

            if self._stop_event.is_set():
                return
            time.sleep(1.0)

    @staticmethod
    def _parse_ant_identifier(identifier: str) -> tuple[int, int, int] | None:
        parts = identifier.split(":")
        if len(parts) != 4:
            return None
        if parts[0].upper() != "ANT":
            return None
        try:
            return int(parts[1]), int(parts[2]), int(parts[3])
        except ValueError:
            return None

    @staticmethod
    def _create_ant_device(
        node: Any,
        device_id: int,
        device_type: int,
        transmission_type: int,
        roles: set[str],
        device_type_enum: Any,
        heart_rate_cls: Any,
        power_meter_cls: Any,
        bike_cadence_cls: Any,
        bike_speed_cadence_cls: Any,
    ) -> Any:
        dt = device_type_enum(device_type)
        if dt == device_type_enum.HeartRate:
            return heart_rate_cls(node, device_id=device_id, trans_type=transmission_type)
        if dt == device_type_enum.PowerMeter:
            return power_meter_cls(node, device_id=device_id, trans_type=transmission_type)
        if dt == device_type_enum.BikeCadence:
            return bike_cadence_cls(node, device_id=device_id, trans_type=transmission_type)
        if dt == device_type_enum.BikeSpeedCadence:
            return bike_speed_cadence_cls(node, device_id=device_id, trans_type=transmission_type)
        if "power" in roles:
            return power_meter_cls(node, device_id=device_id, trans_type=transmission_type)
        if "heart_rate" in roles:
            return heart_rate_cls(node, device_id=device_id, trans_type=transmission_type)
        return bike_cadence_cls(node, device_id=device_id, trans_type=transmission_type)

    def _handle_ant_device_data(self, roles: set[str], page_name: str, data: object) -> None:
        updated = False
        if "heart_rate" in roles:
            heart_rate = getattr(data, "heart_rate", None)
            if isinstance(heart_rate, (int, float)) and int(heart_rate) > 0:
                self._set_latest_value("heart_rate", int(heart_rate))
                updated = True

        if "power" in roles:
            power = getattr(data, "instantaneous_power", None)
            if not isinstance(power, (int, float)):
                power = getattr(data, "average_power", None)
            if isinstance(power, (int, float)) and int(power) >= 0:
                self._set_latest_value("power_watts", int(power))
                updated = True

        if "cadence" in roles:
            cadence: int | float | None = None
            if hasattr(data, "calculated_cadence"):
                cadence = getattr(data, "calculated_cadence", None)
                if cadence is None and hasattr(data, "cadence"):
                    cadence = getattr(data, "cadence")
            elif hasattr(data, "cadence"):
                cadence = getattr(data, "cadence")
            if isinstance(cadence, (int, float)) and 0 <= float(cadence) <= 250:
                self._set_latest_value("cadence_rpm", int(round(float(cadence))))
                updated = True

        if updated or page_name in {"standard_power", "standard_torque", "heart_rate", "bike_cadence"}:
            self._emit_sample()

    def _handle_ble_heart_rate(self, _sender: object, data: bytearray) -> None:
        heart_rate = self._parse_ble_heart_rate_data(data)
        if heart_rate is None:
            return
        self._set_latest_value("heart_rate", heart_rate)
        self._emit_sample()

    def _handle_ble_power(
        self,
        _sender: object,
        data: bytearray,
        allow_cadence_from_power: bool,
    ) -> None:
        if len(data) < 4:
            return

        flags = int.from_bytes(data[0:2], "little")
        power = int.from_bytes(data[2:4], "little", signed=True)
        self._set_latest_value("power_watts", power)

        offset = 4
        if flags & (1 << 0):
            offset += 1
        if flags & (1 << 2):
            offset += 2
        if flags & (1 << 4):
            offset += 6
        if allow_cadence_from_power and flags & (1 << 5) and len(data) >= offset + 4:
            crank_revs = int.from_bytes(data[offset : offset + 2], "little")
            event_time = int.from_bytes(data[offset + 2 : offset + 4], "little")
            cadence = self._compute_cadence(self._power_crank_state, crank_revs, event_time)
            if cadence is not None:
                self._set_latest_value("cadence_rpm", cadence)

        self._emit_sample()

    def _handle_ble_cadence(self, _sender: object, data: bytearray) -> None:
        if len(data) < 1:
            return

        flags = data[0]
        offset = 1
        if flags & 0x01:
            offset += 6
        if flags & 0x02 and len(data) >= offset + 4:
            crank_revs = int.from_bytes(data[offset : offset + 2], "little")
            event_time = int.from_bytes(data[offset + 2 : offset + 4], "little")
            cadence = self._compute_cadence(self._csc_crank_state, crank_revs, event_time)
            if cadence is not None:
                self._set_latest_value("cadence_rpm", cadence)
                self._emit_sample()

    def _compute_cadence(
        self,
        state: CrankState,
        cumulative_revs: int,
        event_time: int,
    ) -> int | None:
        previous_revs = state.cumulative_revs
        previous_time = state.event_time
        state.cumulative_revs = cumulative_revs
        state.event_time = event_time

        if previous_revs is None or previous_time is None:
            return None

        delta_revs = (cumulative_revs - previous_revs) % 65536
        delta_time = (event_time - previous_time) % 65536
        if delta_revs <= 0 or delta_time <= 0:
            return None

        cadence = round((delta_revs * 60 * 1024) / delta_time)
        if cadence < 0 or cadence > 250:
            return None
        return cadence

    def _set_latest_value(self, key: str, value: int | float | None) -> None:
        with self._latest_lock:
            self._latest_values[key] = value
            if value is not None:
                self._last_update_monotonic[key] = time.monotonic()

    def _value_age_seconds(self, key: str) -> float:
        with self._latest_lock:
            updated_at = self._last_update_monotonic.get(key, 0.0)
        if updated_at <= 0:
            return float("inf")
        return max(0.0, time.monotonic() - updated_at)

    def _is_stream_stale(
        self,
        key: str,
        connected_at: float,
        stale_seconds: float,
        bootstrap_seconds: float,
    ) -> bool:
        age = self._value_age_seconds(key)
        if age == float("inf"):
            # No value received yet. Allow a bootstrap window before forcing reconnect.
            return (time.monotonic() - connected_at) > bootstrap_seconds
        return age > stale_seconds

    async def _ble_hr_fallback_loop(self, client: object) -> None:
        while not self._stop_event.is_set():
            await asyncio.sleep(1.0)
            if self._value_age_seconds("heart_rate") <= 2.5:
                continue
            try:
                raw = await client.read_gatt_char(HEART_RATE_MEASUREMENT_UUID)
            except Exception:
                continue
            heart_rate = self._parse_ble_heart_rate_data(raw)
            if heart_rate is None:
                continue
            self._set_latest_value("heart_rate", heart_rate)
            self._emit_sample()

    async def _ble_power_fallback_loop(self, client: object) -> None:
        while not self._stop_event.is_set():
            await asyncio.sleep(1.0)
            if self._value_age_seconds("power_watts") <= 2.5:
                continue
            try:
                raw = await client.read_gatt_char(CYCLING_POWER_MEASUREMENT_UUID)
            except Exception:
                continue
            power = self._parse_ble_power_data(raw)
            if power is None:
                continue
            self._set_latest_value("power_watts", power)
            self._emit_sample()

    @staticmethod
    def _parse_ble_heart_rate_data(data: bytearray | bytes) -> int | None:
        if len(data) < 2:
            return None
        flags = data[0]
        is_uint16 = flags & 0x01
        if is_uint16:
            if len(data) < 3:
                return None
            return int.from_bytes(data[1:3], "little")
        return int(data[1])

    @staticmethod
    def _parse_ble_power_data(data: bytearray | bytes) -> int | None:
        if len(data) < 4:
            return None
        return int.from_bytes(data[2:4], "little", signed=True)

    @staticmethod
    def _monitored_value_keys_for_roles(roles: set[str]) -> list[str]:
        keys: list[str] = []
        if "heart_rate" in roles:
            keys.append("heart_rate")
        if "power" in roles:
            keys.append("power_watts")
        if "cadence" in roles:
            keys.append("cadence_rpm")
        return keys

    def _emit_sample(self) -> None:
        if self._callback is None:
            return
        with self._latest_lock:
            heart_rate = self._coerce_int(self._latest_values["heart_rate"])
            power_watts = self._coerce_int(self._latest_values["power_watts"])
            cadence_rpm = self._coerce_int(self._latest_values["cadence_rpm"])
        self._callback(
            TelemetrySample(
                timestamp=datetime.now(),
                heart_rate=heart_rate,
                power_watts=power_watts,
                cadence_rpm=cadence_rpm,
            )
        )

    @staticmethod
    def _coerce_int(value: int | float | None) -> int | None:
        if value is None:
            return None
        return int(value)
