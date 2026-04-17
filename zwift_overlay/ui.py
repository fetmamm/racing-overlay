from __future__ import annotations

import tkinter as tk
import tkinter.font as tkfont
from tkinter import colorchooser, messagebox, ttk
from typing import Callable
from datetime import datetime
import threading
import time
import traceback
import webbrowser
import sys
import os
import smtplib
import ssl
from email.message import EmailMessage
from urllib.parse import quote

from zwift_overlay.config import AppConfig, SensorBinding, load_app_config, save_app_config
from zwift_overlay.models import SummaryStats, TelemetrySample
from zwift_overlay.sensors import (
    SENSOR_ROLES,
    TRANSPORTS,
    DiscoveredSensor,
    SensorDiscoveryError,
    SensorScanCancelledError,
    SensorDiscoveryService,
)
from zwift_overlay.source_factory import create_telemetry_source
from zwift_overlay.sources.base import TelemetrySource
from zwift_overlay.stats import TelemetryAggregator
from zwift_overlay.version import APP_VERSION

AVG_POWER_PRESET_SECONDS = [10, 30, 60, 120, 180, 240, 300, 600, 900, 1200, 1800, 2700, 3600]
DISCORD_SERVER_URL = "https://discord.gg/3ARGhyAPSZ"
CONTACT_EMAIL = "jesperr.svensson@gmail.com"
HEART_RATE_MEASUREMENT_UUID = "00002a37-0000-1000-8000-00805f9b34fb"
CYCLING_POWER_MEASUREMENT_UUID = "00002a63-0000-1000-8000-00805f9b34fb"
# App-owner managed mail service config (end users should not need to fill SMTP details).
APP_SMTP_HOST = os.getenv("ZWIFT_OVERLAY_SMTP_HOST", "").strip()
try:
    APP_SMTP_PORT = int(os.getenv("ZWIFT_OVERLAY_SMTP_PORT", "587") or "587")
except ValueError:
    APP_SMTP_PORT = 587
APP_SMTP_USERNAME = os.getenv("ZWIFT_OVERLAY_SMTP_USERNAME", "").strip()
APP_SMTP_PASSWORD = os.getenv("ZWIFT_OVERLAY_SMTP_PASSWORD", "")
APP_SMTP_FROM_EMAIL = os.getenv("ZWIFT_OVERLAY_SMTP_FROM_EMAIL", "").strip()
APP_SMTP_USE_TLS = os.getenv("ZWIFT_OVERLAY_SMTP_USE_TLS", "1").strip() not in {"0", "false", "False"}


class OverlayApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.report_callback_exception = self._handle_tk_exception
        self.root.title(f"Zwift Overlay {APP_VERSION}")
        self.base_normal_size = (360, 520)
        self.base_compact_size = (360, 360)
        self.base_normal_minsize = (340, 500)
        self.base_compact_minsize = (340, 340)
        self.ui_scale_factor = 1.0
        self.style = ttk.Style()

        self.aggregator = TelemetryAggregator()
        self.config = load_app_config()
        self.source: TelemetrySource = create_telemetry_source(self.config)
        self.discovery_service = SensorDiscoveryService()
        self.labels: dict[str, tk.StringVar] = {}
        self.is_session_running = False
        self.weight_var = tk.StringVar(value=self.config.rider_weight_input)
        self.topmost_var = tk.BooleanVar(value=self.config.always_on_top)
        self.sensor_status_vars: dict[str, tk.StringVar] = {}
        self.sensor_status_dots: dict[str, int] = {}
        self.status_lights: dict[str, int] = {}
        self.sensor_activity: dict[str, str] = {}
        self.sensor_last_seen_names: dict[str, str] = {}
        self.last_scan_devices_by_transport: dict[str, dict[str, DiscoveredSensor]] = {}
        self.session_started_at: datetime | None = None
        self.accumulated_elapsed_seconds = 0
        self.elapsed_timer_id: str | None = None
        self.current_session_state = "stopped"
        self.refresh_in_progress = False
        self.delayed_start_timer_id: str | None = None
        self.delayed_start_remaining_seconds = 0
        self.summary_render_timer_id: str | None = None
        self.last_summary_render_monotonic = 0.0
        self.pending_summary: SummaryStats | None = None
        self.best_avg_power_by_window: dict[int, float] = {}
        self.sensor_live_value_hints: dict[str, str] = {}
        self.source_stop_token = 0
        self.startup_reconnect_after_id: str | None = None
        self.startup_reconnect_attempt = 0
        self.startup_reconnect_max_attempts = 2
        self.startup_reconnect_interval_ms = 2000
        self.base_tk_scaling = float(self.root.tk.call("tk", "scaling"))
        self.base_font_sizes = {
            "title": 14,
            "value": 11,
            "status": 8,
            "default": 9,
            "header": 10,
            "button": 8,
            "delay_title": 14,
            "delay_overlay": 22,
        }
        self.font_title = tkfont.Font(family="Segoe UI", size=self.base_font_sizes["title"], weight="bold")
        self.font_value = tkfont.Font(family="Consolas", size=self.base_font_sizes["value"], weight="bold")
        self.font_status = tkfont.Font(family="Segoe UI", size=self.base_font_sizes["status"])
        self.font_default = tkfont.Font(family="Segoe UI", size=self.base_font_sizes["default"])
        self.font_header = tkfont.Font(family="Segoe UI", size=self.base_font_sizes["header"], weight="bold")
        self.font_button = tkfont.Font(family="Segoe UI", size=self.base_font_sizes["button"])
        self.font_delay_title = tkfont.Font(
            family="Segoe UI",
            size=self.base_font_sizes["delay_title"],
            weight="bold",
        )
        self.font_delay_overlay = tkfont.Font(
            family="Segoe UI",
            size=self.base_font_sizes["delay_overlay"],
            weight="bold",
        )

        self._apply_ui_scale()
        self._build_ui()
        self._render_summary(self.aggregator.summary())
        self._set_session_state("stopped")
        self._toggle_topmost()
        self._handle_sensor_update()
        self._schedule_startup_auto_reconnect()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        self._configure_styles()

        self.main_frame = ttk.Frame(self.root, padding=12, style="Overlay.TFrame")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.title_row = ttk.Frame(self.main_frame, style="Overlay.TFrame")
        self.title_row.pack(fill=tk.X)
        self.title_label = ttk.Label(
            self.title_row,
            text=f"Zwift Overlay {APP_VERSION}",
            font=self.font_title,
            style="OverlayTitle.TLabel",
        )
        self.title_label.pack(side=tk.LEFT, anchor=tk.W)
        self.settings_button = ttk.Button(
            self.title_row,
            text="\u2699",
            width=3,
            command=self.open_settings_window,
        )
        self.settings_button.pack(side=tk.RIGHT, anchor=tk.E)
        self.contact_button = ttk.Button(
            self.title_row,
            text="\u2709",
            width=3,
            command=self.open_contact_window,
        )
        self.contact_button.pack(side=tk.RIGHT, anchor=tk.E, padx=(0, 6))

        self._build_status_lights(self.main_frame)

        self.metrics_frame = ttk.Frame(self.main_frame, style="Overlay.TFrame")
        self.metrics_frame.pack(fill=tk.X)
        self.metric_rows: dict[str, ttk.Frame] = {}
        self.avg_power_row_vars: dict[int, tuple[tk.StringVar, tk.StringVar, tk.StringVar]] = {}
        self.avg_power_row_frames: dict[int, ttk.Frame] = {}
        self.avg_power_adjusted_labels: dict[int, ttk.Label] = {}

        for label, key in [
            ("Time", "elapsed"),
        ]:
            row = ttk.Frame(self.metrics_frame, style="Overlay.TFrame")
            row.pack(fill=tk.X, pady=2)

            ttk.Label(row, text=label, width=14, style="OverlayMetric.TLabel").pack(side=tk.LEFT)
            value = tk.StringVar(value="-")
            self.labels[key] = value
            ttk.Label(row, textvariable=value, font=self.font_value, style="OverlayValue.TLabel").pack(side=tk.LEFT)

        self.metrics_header_row = ttk.Frame(self.metrics_frame, style="Overlay.TFrame")
        self.metrics_header_row.pack(fill=tk.X, pady=(2, 2))
        ttk.Label(self.metrics_header_row, text="", width=14, style="OverlayMetric.TLabel").pack(side=tk.LEFT)
        metrics_header_group = ttk.Frame(self.metrics_header_row, style="Overlay.TFrame")
        metrics_header_group.pack(side=tk.LEFT)
        ttk.Label(metrics_header_group, text="Power", width=8, style="OverlayHeader.TLabel").pack(side=tk.LEFT)
        ttk.Label(metrics_header_group, text="W/kg", width=8, style="OverlayHeader.TLabel").pack(side=tk.LEFT, padx=(6, 0))
        self.adjusted_wkg_header_var = tk.StringVar(value="")
        self.adjusted_wkg_header_label = ttk.Label(
            metrics_header_group,
            textvariable=self.adjusted_wkg_header_var,
            width=10,
            style="OverlayHeader.TLabel",
        )
        self.adjusted_wkg_header_label.pack(side=tk.LEFT, padx=(6, 0))

        effect_row = ttk.Frame(self.metrics_frame, style="Overlay.TFrame")
        effect_row.pack(fill=tk.X, pady=2)
        self.power_label_var = tk.StringVar(value="Power")
        ttk.Label(effect_row, textvariable=self.power_label_var, width=14, style="OverlayMetric.TLabel").pack(side=tk.LEFT)

        power_group = ttk.Frame(effect_row, style="Overlay.TFrame")
        power_group.pack(side=tk.LEFT)

        power_value = tk.StringVar(value="-")
        self.labels["current_power"] = power_value
        ttk.Label(
            power_group,
            textvariable=power_value,
            font=self.font_value,
            width=8,
            style="OverlayValue.TLabel",
        ).pack(side=tk.LEFT)

        wkg_value = tk.StringVar(value="-")
        self.labels["current_wkg"] = wkg_value
        ttk.Label(
            power_group,
            textvariable=wkg_value,
            font=self.font_value,
            width=8,
            style="OverlayValue.TLabel",
        ).pack(side=tk.LEFT, padx=(6, 0))
        current_adjusted_wkg_value = tk.StringVar(value="-")
        self.labels["current_adjusted_wkg"] = current_adjusted_wkg_value
        self.current_adjusted_wkg_label = ttk.Label(
            power_group,
            textvariable=current_adjusted_wkg_value,
            font=self.font_value,
            width=9,
            style="OverlayValue.TLabel",
        )
        self.current_adjusted_wkg_label.pack(side=tk.LEFT, padx=(6, 0))

        self.avg_power_rows_container = ttk.Frame(self.metrics_frame, style="Overlay.TFrame")
        self.avg_power_rows_container.pack(fill=tk.X)
        self._build_avg_power_rows()

        session_avg_row = ttk.Frame(self.metrics_frame, style="Overlay.TFrame")
        session_avg_row.pack(fill=tk.X, pady=2)
        self.metric_rows["session_avg_power"] = session_avg_row
        ttk.Label(session_avg_row, text="Power (AVG)", width=14, style="OverlayMetric.TLabel").pack(side=tk.LEFT)

        session_avg_group = ttk.Frame(session_avg_row, style="Overlay.TFrame")
        session_avg_group.pack(side=tk.LEFT)

        avg_power_value = tk.StringVar(value="-")
        self.labels["avg_power"] = avg_power_value
        ttk.Label(
            session_avg_group,
            textvariable=avg_power_value,
            font=self.font_value,
            width=8,
            style="OverlayValue.TLabel",
        ).pack(side=tk.LEFT)

        avg_power_wkg_value = tk.StringVar(value="-")
        self.labels["avg_power_wkg"] = avg_power_wkg_value
        ttk.Label(
            session_avg_group,
            textvariable=avg_power_wkg_value,
            font=self.font_value,
            width=8,
            style="OverlayValue.TLabel",
        ).pack(side=tk.LEFT, padx=(6, 0))
        avg_power_adjusted_wkg_value = tk.StringVar(value="-")
        self.labels["avg_power_adjusted_wkg"] = avg_power_adjusted_wkg_value
        self.avg_power_adjusted_wkg_label = ttk.Label(
            session_avg_group,
            textvariable=avg_power_adjusted_wkg_value,
            font=self.font_value,
            width=9,
            style="OverlayValue.TLabel",
        )
        self.avg_power_adjusted_wkg_label.pack(side=tk.LEFT, padx=(6, 0))

        hr_row = ttk.Frame(self.metrics_frame, style="Overlay.TFrame")
        hr_row.pack(fill=tk.X, pady=2)
        self.metric_rows["hr_summary"] = hr_row
        ttk.Label(hr_row, text="HR/AVG/MAX", width=14, style="OverlayMetric.TLabel").pack(side=tk.LEFT)

        hr_group = ttk.Frame(hr_row, style="Overlay.TFrame")
        hr_group.pack(side=tk.LEFT)
        current_hr = tk.StringVar(value="-")
        avg_hr = tk.StringVar(value="-")
        max_hr = tk.StringVar(value="-")
        self.labels["current_hr"] = current_hr
        self.labels["avg_hr"] = avg_hr
        self.labels["max_hr"] = max_hr
        ttk.Label(
            hr_group,
            textvariable=current_hr,
            font=self.font_value,
            width=8,
            style="OverlayValue.TLabel",
        ).pack(side=tk.LEFT)
        ttk.Label(
            hr_group,
            textvariable=avg_hr,
            font=self.font_value,
            width=8,
            style="OverlayValue.TLabel",
        ).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(
            hr_group,
            textvariable=max_hr,
            font=self.font_value,
            width=9,
            style="OverlayValue.TLabel",
        ).pack(side=tk.LEFT, padx=(6, 0))

        speed_row = ttk.Frame(self.metrics_frame, style="Overlay.TFrame")
        speed_row.pack(fill=tk.X, pady=2)
        self.metric_rows["avg_speed"] = speed_row
        ttk.Label(speed_row, text="Speed / AVG", width=14, style="OverlayMetric.TLabel").pack(side=tk.LEFT)

        speed_group = ttk.Frame(speed_row, style="Overlay.TFrame")
        speed_group.pack(side=tk.LEFT)
        current_speed = tk.StringVar(value="-")
        avg_speed = tk.StringVar(value="-")
        self.labels["current_speed"] = current_speed
        self.labels["avg_speed"] = avg_speed
        ttk.Label(
            speed_group,
            textvariable=current_speed,
            font=self.font_value,
            width=8,
            style="OverlayValue.TLabel",
        ).pack(side=tk.LEFT)
        ttk.Label(
            speed_group,
            textvariable=avg_speed,
            font=self.font_value,
            width=8,
            style="OverlayValue.TLabel",
        ).pack(side=tk.LEFT, padx=(6, 0))

        self._sync_power_label()
        self._sync_adjusted_wkg_ui()
        self._apply_metric_visibility()

        self.controls = ttk.Frame(self.main_frame, style="Overlay.TFrame")
        self.controls.pack(fill=tk.X, pady=(14, 0))
        self.controls_top = ttk.Frame(self.controls, style="Overlay.TFrame")
        self.controls_top.pack(fill=tk.X)
        self.controls_bottom = ttk.Frame(self.controls, style="Overlay.TFrame")
        self.controls_bottom.pack(fill=tk.X, pady=(4, 0))

        self.start_button = ttk.Button(self.controls_top, text="Start", command=self.start)
        self.start_button.pack(side=tk.LEFT)
        self.delayed_start_button = ttk.Button(self.controls_top, text="Delayed start", command=self.start_delayed)
        self.delayed_start_button.pack(side=tk.LEFT, padx=(8, 0))
        self._refresh_delayed_start_button_label()
        self.pause_button = ttk.Button(self.controls_bottom, text="Pause", command=self.pause)
        self.pause_button.pack(side=tk.LEFT)
        self.reset_button = ttk.Button(self.controls_bottom, text="Reset", command=self.reset)
        self.reset_button.pack(side=tk.LEFT, padx=(8, 0))
        self.stop_button = ttk.Button(self.controls_bottom, text="Stop", command=self.stop)
        self.stop_button.pack(side=tk.LEFT, padx=(8, 0))

        self.controls_separator = ttk.Separator(self.main_frame, orient=tk.HORIZONTAL)
        self.controls_separator.pack(fill=tk.X, pady=(10, 8))

        self.config_row = ttk.Frame(self.main_frame, style="Overlay.TFrame")
        self.config_row.pack(fill=tk.X, pady=(0, 0))
        ttk.Button(self.config_row, text="Sensors", command=self.open_sensor_window).pack(side=tk.LEFT)
        self.refresh_button = ttk.Button(
            self.config_row,
            text="Refresh sensors",
            command=self.refresh_selected_sensors,
        )
        self.refresh_button.pack(side=tk.LEFT, padx=(8, 0))

        self.sensor_info_frame = ttk.Frame(self.main_frame, style="Overlay.TFrame")
        self.sensor_info_frame.pack(fill=tk.X, pady=(8, 0))
        self._build_sensor_status_rows(self.sensor_info_frame)

        self.status_var = tk.StringVar(value="Waiting for session")
        self.status_label = ttk.Label(
            self.main_frame,
            textvariable=self.status_var,
            font=self.font_status,
            wraplength=320,
            justify=tk.LEFT,
            style="OverlayStatus.TLabel",
        )
        self.status_label.pack(
            anchor=tk.W,
            fill=tk.X,
            pady=(12, 0),
        )
        self.footer_separator = ttk.Separator(self.main_frame, orient=tk.HORIZONTAL)
        self.footer_separator.pack(fill=tk.X, pady=(8, 6))
        self.footer_label = ttk.Label(
            self.main_frame,
            text="@slangens 2026",
            font=self.font_status,
            style="OverlayStatus.TLabel",
            justify=tk.RIGHT,
        )
        self.footer_label.pack(anchor=tk.E, fill=tk.X, pady=(6, 0))

        self.delay_overlay_var = tk.StringVar(value="")
        self.delay_overlay_frame = ttk.Frame(self.main_frame, style="Overlay.TFrame", padding=(16, 10))
        self.delay_overlay_title_label = ttk.Label(
            self.delay_overlay_frame,
            text="Start in:",
            font=self.font_delay_title,
            style="OverlayMetric.TLabel",
        )
        self.delay_overlay_title_label.pack()
        self.delay_overlay_label = ttk.Label(
            self.delay_overlay_frame,
            textvariable=self.delay_overlay_var,
            font=self.font_delay_overlay,
            style="OverlayTitle.TLabel",
        )
        self.delay_overlay_label.pack()

    def _build_status_lights(self, parent: ttk.Frame) -> None:
        self.header_status_row = ttk.Frame(parent, style="Overlay.TFrame")
        self.header_status_row.pack(fill=tk.X, pady=(8, 4))

        self.light_canvas = tk.Canvas(
            self.header_status_row,
            width=98,
            height=18,
            highlightthickness=0,
            bd=0,
            bg=self.config.inactive_background,
        )
        self.light_canvas.pack(side=tk.LEFT)

        for x, state, _color in [
            (8, "running", "#39b36b"),
            (44, "paused", "#d9b43b"),
            (80, "stopped", "#d95c5c"),
        ]:
            light = self.light_canvas.create_oval(x, 3, x + 14, 17, fill="#d8d8d8", outline="")
            self.status_lights[state] = light

    def _build_sensor_status_rows(self, parent: ttk.Frame) -> None:
        display_labels = {
            "power": "Power",
            "heart_rate": "Heart Rate",
        }
        for role in ("power", "heart_rate"):
            row = ttk.Frame(parent, style="Overlay.TFrame")
            row.pack(fill=tk.X, pady=1)

            canvas = tk.Canvas(
                row,
                width=14,
                height=14,
                highlightthickness=0,
                bd=0,
                bg=self.config.inactive_background,
            )
            canvas.pack(side=tk.LEFT, pady=(1, 0))
            dot = canvas.create_oval(2, 2, 12, 12, fill="#d8d8d8", outline="")
            self.sensor_status_dots[role] = dot
            setattr(self, f"_sensor_canvas_{role}", canvas)

            value = tk.StringVar(value=f"{display_labels[role]}: (none selected)")
            self.sensor_status_vars[role] = value
            ttk.Label(
                row,
                textvariable=value,
                font=self.font_status,
                justify=tk.LEFT,
                wraplength=300,
                style="OverlayStatus.TLabel",
            ).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6, 0))

    def _configure_styles(self) -> None:
        background = self.config.inactive_background
        button_pad_x = self._scaled_int(4)
        button_pad_y = self._scaled_int(2)
        self.root.configure(bg=background)
        self.style.configure("Overlay.TFrame", background=background)
        self.style.configure("Overlay.TLabel", background=background, foreground="#111111", font=self.font_default)
        self.style.configure("OverlayTitle.TLabel", background=background, foreground="#111111", font=self.font_title)
        self.style.configure("OverlayMetric.TLabel", background=background, foreground="#111111", font=self.font_default)
        self.style.configure("OverlayHeader.TLabel", background=background, foreground="#111111", font=self.font_header)
        self.style.configure("OverlayValue.TLabel", background=background, foreground="#111111", font=self.font_value)
        self.style.configure("OverlayStatus.TLabel", background=background, foreground="#222222", font=self.font_status)
        self.style.configure("TButton", font=self.font_button, padding=(button_pad_x, button_pad_y))
        self.style.configure("TCheckbutton", font=self.font_default)
        self.style.configure("TRadiobutton", font=self.font_default)
        self.style.configure("TCombobox", font=self.font_default)

    def _toggle_topmost(self) -> None:
        self.root.attributes("-topmost", self.topmost_var.get())
        self._apply_window_appearance()

    def _sync_power_label(self) -> None:
        window_seconds = max(1, int(self.config.power_display_seconds))
        self.power_label_var.set(f"Power ({window_seconds}s)")

    def _sync_adjusted_wkg_ui(self) -> None:
        percent = 95 if int(self.config.adjusted_wkg_percent) == 95 else 90
        show_column = bool(self.config.show_adjusted_wkg_column)
        if hasattr(self, "adjusted_wkg_header_var"):
            self.adjusted_wkg_header_var.set(f"{percent}% W/kg")
        if hasattr(self, "adjusted_wkg_header_label"):
            self._set_optional_column_widget(self.adjusted_wkg_header_label, show_column)
        if hasattr(self, "current_adjusted_wkg_label"):
            self._set_optional_column_widget(self.current_adjusted_wkg_label, show_column)
        if hasattr(self, "avg_power_adjusted_wkg_label"):
            self._set_optional_column_widget(self.avg_power_adjusted_wkg_label, show_column)
        for label in self.avg_power_adjusted_labels.values():
            self._set_optional_column_widget(label, show_column)

    @staticmethod
    def _set_optional_column_widget(widget: tk.Misc, show: bool) -> None:
        if show and not widget.winfo_manager():
            widget.pack(side=tk.LEFT, padx=(6, 0))
        if not show and widget.winfo_manager():
            widget.pack_forget()

    def _selected_avg_power_windows(self) -> list[int]:
        windows = {max(1, int(value)) for value in self.config.avg_power_windows_seconds if int(value) > 0}
        if self.config.show_custom_avg_power and self.config.custom_avg_power_seconds > 0:
            windows.add(int(self.config.custom_avg_power_seconds))
        return sorted(windows)

    def _build_avg_power_rows(self) -> None:
        for row in self.avg_power_row_frames.values():
            row.destroy()
        self.avg_power_row_vars = {}
        self.avg_power_row_frames = {}
        self.avg_power_adjusted_labels = {}

        for seconds in self._selected_avg_power_windows():
            row = ttk.Frame(self.avg_power_rows_container, style="Overlay.TFrame")
            row.pack(fill=tk.X, pady=2)
            label = self._format_duration_label(seconds)
            label = f"{label} (best)"
            ttk.Label(
                row,
                text=label,
                width=14,
                style="OverlayMetric.TLabel",
            ).pack(side=tk.LEFT)

            value_group = ttk.Frame(row, style="Overlay.TFrame")
            value_group.pack(side=tk.LEFT)

            power_value = tk.StringVar(value="-")
            wkg_value = tk.StringVar(value="-")
            adjusted_wkg_value = tk.StringVar(value="-")
            ttk.Label(
                value_group,
                textvariable=power_value,
                font=self.font_value,
                width=8,
                style="OverlayValue.TLabel",
            ).pack(side=tk.LEFT)
            ttk.Label(
                value_group,
                textvariable=wkg_value,
                font=self.font_value,
                width=8,
                style="OverlayValue.TLabel",
            ).pack(side=tk.LEFT, padx=(6, 0))
            adjusted_label = ttk.Label(
                value_group,
                textvariable=adjusted_wkg_value,
                font=self.font_value,
                width=9,
                style="OverlayValue.TLabel",
            )
            adjusted_label.pack(side=tk.LEFT, padx=(6, 0))
            self.avg_power_row_frames[seconds] = row
            self.avg_power_row_vars[seconds] = (power_value, wkg_value, adjusted_wkg_value)
            self.avg_power_adjusted_labels[seconds] = adjusted_label

        self._sync_adjusted_wkg_ui()

    @staticmethod
    def _format_duration_label(seconds: int) -> str:
        if seconds % 60 == 0:
            return f"{seconds // 60}min"
        return f"{seconds}s"

    def _apply_metric_visibility(self) -> None:
        visibility = {
            "session_avg_power": self.config.show_session_avg_power,
            "hr_summary": self.config.show_avg_hr,
            "avg_speed": self.config.show_avg_speed,
        }
        for row_key, row in self.metric_rows.items():
            should_show = visibility.get(row_key, True)
            if should_show and not row.winfo_manager():
                row.pack(fill=tk.X, pady=2)
            if not should_show and row.winfo_manager():
                row.pack_forget()

    def _set_session_state(self, state: str) -> None:
        self.current_session_state = state
        colors = {
            "running": "#39b36b",
            "paused": "#d9b43b",
            "stopped": "#d95c5c",
        }
        inactive = "#d8d8d8"
        for key, item_id in self.status_lights.items():
            color = colors[key] if key == state else inactive
            self.light_canvas.itemconfig(item_id, fill=color)
        self._update_control_buttons(state)
        self._apply_window_appearance()

    def _update_control_buttons(self, state: str) -> None:
        delayed_active = self.delayed_start_timer_id is not None or self.delayed_start_remaining_seconds > 0
        if state == "running":
            self.start_button.config(state=tk.DISABLED)
            self.pause_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.NORMAL)
            self.reset_button.config(state=tk.NORMAL)
            self.delayed_start_button.config(state=tk.DISABLED)
            if self.delayed_start_button.winfo_manager():
                self.delayed_start_button.pack_forget()
            return
        if state == "paused":
            if not self.delayed_start_button.winfo_manager():
                self.delayed_start_button.pack(side=tk.LEFT, padx=(8, 0), after=self.start_button)
            self.start_button.config(state=tk.NORMAL)
            self.pause_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.reset_button.config(state=tk.NORMAL)
            self.delayed_start_button.config(state=tk.NORMAL)
            return

        if not self.delayed_start_button.winfo_manager():
            self.delayed_start_button.pack(side=tk.LEFT, padx=(8, 0), after=self.start_button)
        self.start_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL if delayed_active else tk.DISABLED)
        self.reset_button.config(state=tk.NORMAL)
        self.delayed_start_button.config(state=tk.NORMAL)

    def start(self) -> None:
        self._cancel_delayed_start()
        self._cancel_startup_auto_reconnect()
        if self.is_session_running:
            return
        if not self._ensure_weight_before_start():
            return
        self.is_session_running = True
        self.session_started_at = datetime.now()
        self.status_var.set(f"Running source: {self.source.name}")
        try:
            self.source.start(self._threadsafe_handle_sample)
            self._set_session_state("running")
            self._schedule_elapsed_tick()
        except NotImplementedError as exc:
            self.is_session_running = False
            self.session_started_at = None
            self.status_var.set(str(exc))
            self._set_session_state("stopped")
        except Exception as exc:
            self.is_session_running = False
            self.session_started_at = None
            self.status_var.set(f"Start failed: {exc}")
            self._set_session_state("stopped")

    def pause(self) -> None:
        self._cancel_delayed_start()
        if self.is_session_running and self.session_started_at is not None:
            elapsed = datetime.now() - self.session_started_at
            self.accumulated_elapsed_seconds += max(0, int(elapsed.total_seconds()))
        self.is_session_running = False
        self.session_started_at = None
        self._cancel_elapsed_tick()
        self._set_session_state("paused")
        self._stop_source_async("Pause")
        self.status_var.set("Paused")
        self._render_summary(self.aggregator.summary())

    def reset(self) -> None:
        self._cancel_delayed_start()
        was_running = self.current_session_state == "running"
        if self.summary_render_timer_id is not None:
            self.root.after_cancel(self.summary_render_timer_id)
            self.summary_render_timer_id = None
        self.pending_summary = None
        self.aggregator = TelemetryAggregator()
        self.best_avg_power_by_window = {}
        self.accumulated_elapsed_seconds = 0
        if was_running:
            # Keep sensor streams running and just reset session counters/statistics.
            self.is_session_running = True
            self.session_started_at = datetime.now()
            self._set_session_state("running")
            self._schedule_elapsed_tick()
            self._render_summary(self.aggregator.summary())
            self.status_var.set("Reset")
            return
        self.is_session_running = False
        self.session_started_at = None
        self._cancel_elapsed_tick()
        self._set_session_state("stopped")
        self._render_summary(self.aggregator.summary())
        self._stop_source_async("Reset")
        self.status_var.set("Reset")

    def stop(self) -> None:
        self._cancel_delayed_start()
        self.is_session_running = False
        self.session_started_at = None
        if self.summary_render_timer_id is not None:
            self.root.after_cancel(self.summary_render_timer_id)
            self.summary_render_timer_id = None
        self.pending_summary = None
        self._cancel_elapsed_tick()
        self._set_session_state("stopped")
        self._stop_source_async("Stop")
        self.accumulated_elapsed_seconds = 0
        self.aggregator = TelemetryAggregator()
        self.best_avg_power_by_window = {}
        self._render_summary(self.aggregator.summary())
        self.status_var.set("Stopped")

    def _restart_after_reset(self) -> None:
        if self.current_session_state != "stopped" or self.is_session_running:
            return
        self.start()

    def _stop_source_async(self, action: str, on_complete: Callable[[], None] | None = None) -> None:
        self.source_stop_token += 1
        token = self.source_stop_token
        source = self.source

        def _stop_worker() -> None:
            stop_error: Exception | None = None
            try:
                source.stop()
            except Exception as exc:
                stop_error = exc
            try:
                self.root.after(0, lambda: self._on_stop_source_finished(token, action, stop_error, on_complete))
            except tk.TclError:
                return

        threading.Thread(target=_stop_worker, daemon=True).start()

    def _on_stop_source_finished(
        self,
        token: int,
        action: str,
        stop_error: Exception | None,
        on_complete: Callable[[], None] | None,
    ) -> None:
        if token != self.source_stop_token:
            return
        if stop_error is not None and self.current_session_state != "running":
            self.status_var.set(f"{action} warning: {stop_error}")
        if on_complete is not None:
            on_complete()

    def _stop_source_safely(self, action: str, timeout_seconds: float = 1.2) -> bool:
        stop_error: list[Exception] = []

        def _stop_worker() -> None:
            try:
                self.source.stop()
            except Exception as exc:
                stop_error.append(exc)

        thread = threading.Thread(target=_stop_worker, daemon=True)
        thread.start()
        thread.join(timeout=timeout_seconds)
        if thread.is_alive():
            self.status_var.set(f"{action} in progress...")
            return False
        if stop_error:
            self.status_var.set(f"{action} warning: {stop_error[0]}")
        return True

    def _schedule_restart_after_reset(self, attempts: int, interval_ms: int) -> None:
        if attempts <= 0 or self.is_session_running:
            return
        self.start()
        if self.is_session_running:
            return
        self.root.after(interval_ms, lambda: self._schedule_restart_after_reset(attempts - 1, interval_ms))

    def start_delayed(self) -> None:
        if self.is_session_running or self.delayed_start_timer_id is not None:
            return
        if not self._ensure_weight_before_start():
            return
        self.delayed_start_remaining_seconds = max(1, int(self.config.delayed_start_seconds))
        self.delayed_start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self._show_delay_overlay(self.delayed_start_remaining_seconds)
        self._tick_delayed_start()

    def _ensure_weight_before_start(self) -> bool:
        if self.weight_var.get().strip():
            return True
        self.status_var.set("Weight is empty. Open Settings to set rider weight.")
        self._show_missing_weight_dialog()
        return False

    def _show_missing_weight_dialog(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Missing weight")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            frame,
            text="Weight is empty.\nOpen Settings to set your rider weight before starting?",
            justify=tk.LEFT,
        ).pack(anchor=tk.W)

        buttons = ttk.Frame(frame)
        buttons.pack(anchor=tk.E, pady=(10, 0))
        result = {"open_settings": False}

        def _close(open_settings: bool) -> None:
            result["open_settings"] = open_settings
            dialog.destroy()

        ttk.Button(buttons, text="Settings", command=lambda: _close(True)).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Cancel", command=lambda: _close(False)).pack(side=tk.LEFT, padx=(8, 0))

        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()
        pos_x = root_x + max(0, (root_w - width) // 2)
        pos_y = root_y + max(0, (root_h - height) // 2)
        dialog.geometry(f"+{pos_x}+{pos_y}")

        dialog.wait_window()
        if result["open_settings"]:
            self.open_settings_window()

    def _tick_delayed_start(self) -> None:
        if self.delayed_start_remaining_seconds <= 0:
            self.delayed_start_timer_id = None
            self._refresh_delayed_start_button_label()
            self._hide_delay_overlay()
            self.start()
            return
        self._show_delay_overlay(self.delayed_start_remaining_seconds)
        self.delayed_start_remaining_seconds -= 1
        self.delayed_start_timer_id = self.root.after(1000, self._tick_delayed_start)

    def _cancel_delayed_start(self) -> None:
        if self.delayed_start_timer_id is not None:
            self.root.after_cancel(self.delayed_start_timer_id)
            self.delayed_start_timer_id = None
        self.delayed_start_remaining_seconds = 0
        self._hide_delay_overlay()
        if hasattr(self, "delayed_start_button"):
            self._refresh_delayed_start_button_label()
            next_state = tk.DISABLED if self.current_session_state == "running" else tk.NORMAL
            self.delayed_start_button.config(state=next_state)

    def _show_delay_overlay(self, seconds: int) -> None:
        if not hasattr(self, "delay_overlay_frame"):
            return
        self.delay_overlay_var.set(f"{seconds}s")
        self.delay_overlay_frame.configure(relief=tk.SOLID, borderwidth=1)
        self.delay_overlay_frame.place(relx=0.5, rely=0.38, anchor=tk.CENTER)
        self.delay_overlay_frame.lift()

    def _hide_delay_overlay(self) -> None:
        if not hasattr(self, "delay_overlay_frame"):
            return
        self.delay_overlay_var.set("")
        self.delay_overlay_frame.place_forget()

    def _refresh_delayed_start_button_label(self) -> None:
        if not hasattr(self, "delayed_start_button"):
            return
        selected = max(1, int(self.config.delayed_start_seconds))
        self.delayed_start_button.config(text=f"Delayed start ({selected}s)")

    def _handle_sample(self, sample: TelemetrySample) -> None:
        if not self.is_session_running:
            return
        summary = self.aggregator.add_sample(sample)
        self._update_best_avg_power_windows()
        self._queue_summary_render(summary)

    def _update_best_avg_power_windows(self) -> None:
        for seconds in self._selected_avg_power_windows():
            rolling_power = self.aggregator.rolling_average("power_watts", seconds)
            if rolling_power is None:
                continue
            previous_best = self.best_avg_power_by_window.get(seconds)
            if previous_best is None or rolling_power > previous_best:
                self.best_avg_power_by_window[seconds] = rolling_power

    def _threadsafe_handle_sample(self, sample: TelemetrySample) -> None:
        self.root.after(0, lambda: self._handle_sample(sample))

    def _queue_summary_render(self, summary: SummaryStats) -> None:
        self.pending_summary = summary
        # Keep metric updates synced with elapsed timer tick (1 Hz).

    def _flush_summary_render(self) -> None:
        self.summary_render_timer_id = None
        summary = self.pending_summary
        if summary is None:
            return
        self.pending_summary = None
        self.last_summary_render_monotonic = time.monotonic()
        self._render_summary(summary)

    def _render_summary(self, summary: SummaryStats) -> None:
        power_window_seconds = max(1, int(self.config.power_display_seconds))
        displayed_power = self.aggregator.rolling_average("power_watts", power_window_seconds)
        self.labels["current_power"].set(self._format_power(displayed_power))
        self.labels["current_wkg"].set(self._format_wkg(displayed_power))
        self.labels["current_adjusted_wkg"].set(self._format_adjusted_wkg(displayed_power))
        self.labels["elapsed"].set(self._format_elapsed(self._current_elapsed_seconds()))

        for seconds, value_vars in self.avg_power_row_vars.items():
            power_value, wkg_value, adjusted_wkg_value = value_vars
            best_power = self.best_avg_power_by_window.get(seconds)
            power_value.set(self._format_power(best_power))
            wkg_value.set(self._format_wkg(best_power))
            adjusted_wkg_value.set(self._format_adjusted_wkg(best_power))

        self.labels["avg_power"].set(self._format_power(summary.average_power_watts))
        self.labels["avg_power_wkg"].set(self._format_wkg(summary.average_power_watts))
        self.labels["avg_power_adjusted_wkg"].set(self._format_adjusted_wkg(summary.average_power_watts))
        self.labels["current_hr"].set(self._format_int(summary.current_heart_rate, ""))
        self.labels["avg_hr"].set(self._format_int(self._rounded(summary.average_heart_rate), ""))
        self.labels["max_hr"].set(self._format_int(summary.max_heart_rate, ""))
        self.labels["current_speed"].set(self._format_float(summary.current_speed_kph, ""))
        self.labels["avg_speed"].set(self._format_float(summary.average_speed_kph, ""))
        self._update_sensor_live_value_hints(summary)

    def _update_sensor_live_value_hints(self, summary: SummaryStats) -> None:
        if not self.is_session_running:
            self.sensor_live_value_hints = {}
            return
        hints: dict[str, str] = {}
        for role, binding in self.config.sensors.items():
            if role == "heart_rate" and summary.current_heart_rate is not None:
                hints[binding.identifier] = f"{summary.current_heart_rate} bpm"
            elif role == "power" and summary.current_power_watts is not None:
                hints[binding.identifier] = f"{int(round(summary.current_power_watts))} W"
        self.sensor_live_value_hints = hints

    @staticmethod
    def _format_int(value: int | None, suffix: str) -> str:
        if value is None:
            return "-"
        if not suffix:
            return f"{value}"
        return f"{value} {suffix}"

    @staticmethod
    def _format_float(value: float | None, suffix: str) -> str:
        if value is None:
            return "-"
        if not suffix:
            return f"{value:.1f}"
        return f"{value:.1f} {suffix}"

    @staticmethod
    def _format_power(value: float | None) -> str:
        if value is None:
            return "-"
        return f"{int(round(value))} W"

    @staticmethod
    def _rounded(value: float | None) -> int | None:
        if value is None:
            return None
        return int(round(value))

    def _format_wkg(self, power_watts: int | float | None) -> str:
        if power_watts is None:
            return "-"

        try:
            weight_kg = float(self.weight_var.get().replace(",", "."))
        except ValueError:
            return "Enter weight"

        if weight_kg <= 0:
            return "Invalid weight"
        decimals = 1 if int(self.config.wkg_decimals) != 2 else 2
        return f"{float(power_watts) / weight_kg:.{decimals}f}".replace(".", ",")

    def _format_adjusted_wkg(self, power_watts: int | float | None) -> str:
        if power_watts is None:
            return "-"
        percent = 95 if int(self.config.adjusted_wkg_percent) == 95 else 90
        adjusted_power = float(power_watts) * (percent / 100.0)
        return self._format_wkg(adjusted_power)

    @staticmethod
    def _format_elapsed(elapsed_seconds: int) -> str:
        hours, remainder = divmod(elapsed_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        if hours > 0:
            return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        return f"{minutes:02d}:{seconds:02d}"

    def open_sensor_window(self) -> None:
        try:
            SensorConfigWindow(
                self.root,
                self.config,
                self.discovery_service,
                self._handle_sensor_update,
                self._handle_scan_results,
                self._sensor_state_for_role,
                self._sensor_live_value_hints,
            )
        except Exception as exc:
            self.status_var.set(f"Could not open Sensors: {exc}")
            messagebox.showerror("Sensors", f"Could not open Sensors window.\n{exc}", parent=self.root)

    def _sensor_state_for_role(self, role: str) -> str:
        return self.sensor_activity.get(role, "unverified")

    def _sensor_live_value_hints(self) -> dict[str, str]:
        return dict(self.sensor_live_value_hints)

    def open_settings_window(self) -> None:
        try:
            SettingsWindow(self.root, self.config, self._handle_settings_update)
        except Exception as exc:
            self.status_var.set(f"Could not open Settings: {exc}")
            messagebox.showerror("Settings", f"Could not open Settings window.\n{exc}", parent=self.root)

    def open_contact_window(self) -> None:
        try:
            ContactWindow(self.root, self.config)
        except Exception as exc:
            self.status_var.set(f"Could not open Contact: {exc}")
            messagebox.showerror("Contact", f"Could not open Contact window.\n{exc}", parent=self.root)

    def _handle_tk_exception(self, exc_type: type[BaseException], exc_value: BaseException, exc_tb: object) -> None:
        trace_text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        self.status_var.set(f"UI error handled: {exc_value}")
        try:
            messagebox.showerror(
                "Unexpected error",
                f"An error occurred, but the app kept running.\n\n{exc_value}",
                parent=self.root,
            )
        except Exception:
            pass
        print(trace_text)

    def _handle_sensor_update(self) -> None:
        if not self.is_session_running and self.current_session_state != "running":
            self.source = create_telemetry_source(self.config)
        selected = []
        display_labels = {
            "power": "Power",
            "heart_rate": "Heart Rate",
        }
        for role, label in SENSOR_ROLES.items():
            binding = self.config.sensors.get(role)
            display_label = display_labels.get(role, label)
            if binding is None:
                self.sensor_activity.pop(role, None)
                self.sensor_status_vars[role].set(f"{display_label}: (none selected)")
                self._set_sensor_dot(role, "none")
                continue
            scanned_devices = self.last_scan_devices_by_transport.get(binding.transport, {})
            if binding.identifier in scanned_devices:
                self.sensor_activity[role] = "active"
            elif role not in self.sensor_activity:
                self.sensor_activity[role] = "unverified"
            selected.append(f"{display_label}: {binding.name}")
            status = self.sensor_activity.get(role, "unverified")
            self.sensor_status_vars[role].set(
                f"{display_label}: {binding.name} ({self._format_sensor_state(status)})"
            )
            self._set_sensor_dot(role, status)

        if selected and not self.is_session_running:
            self.status_var.set("Ready to start with selected sensors")
        save_app_config(self.config)
        self._apply_window_appearance()

    def _handle_settings_update(self) -> None:
        self.topmost_var.set(self.config.always_on_top)
        self.weight_var.set(self.config.rider_weight_input)
        self._apply_ui_scale()
        self._sync_power_label()
        self._sync_adjusted_wkg_ui()
        self._refresh_delayed_start_button_label()
        self._build_avg_power_rows()
        self._apply_metric_visibility()
        save_app_config(self.config)
        self._toggle_topmost()
        self._render_summary(self.aggregator.summary())
        self._apply_window_appearance()

    def refresh_selected_sensors(self) -> None:
        if self.refresh_in_progress:
            return
        if not self.config.sensors:
            self.status_var.set("No sensors selected.")
            return
        if self._all_selected_sensors_active():
            self.status_var.set("All selected sensors are already active.")
            return
        self.refresh_in_progress = True
        self.refresh_button.config(state=tk.DISABLED)
        self.status_var.set("Refreshing selected sensors...")
        thread = threading.Thread(target=self._refresh_selected_sensors_worker, daemon=True)
        thread.start()

    def _refresh_selected_sensors_worker(self) -> None:
        devices_by_transport: dict[str, list[DiscoveredSensor]] = {}
        error_message: str | None = None
        transports = sorted({binding.transport for binding in self.config.sensors.values()})
        try:
            for transport in transports:
                devices_by_transport[transport] = self.discovery_service.scan(transport)
        except SensorScanCancelledError:
            error_message = "Refresh stopped."
        except SensorDiscoveryError as exc:
            error_message = str(exc)
        except Exception as exc:
            error_message = f"Refresh failed: {exc}"
        self.root.after(0, lambda: self._finish_sensor_refresh(devices_by_transport, error_message))

    def _finish_sensor_refresh(
        self,
        devices_by_transport: dict[str, list[DiscoveredSensor]],
        error_message: str | None,
    ) -> None:
        self.refresh_in_progress = False
        self.refresh_button.config(state=tk.NORMAL)
        for transport, devices in devices_by_transport.items():
            self._handle_scan_results(transport, devices)
        if self._all_selected_sensors_active():
            self._cancel_startup_auto_reconnect()
        if error_message is not None:
            self.status_var.set(error_message)
            return
        self.status_var.set("Sensor refresh complete.")

    def _schedule_startup_auto_reconnect(self) -> None:
        if not self.config.sensors:
            return
        self._cancel_startup_auto_reconnect(clear_attempts=False)
        self.startup_reconnect_attempt = 0
        self.startup_reconnect_after_id = self.root.after(2000, self._run_startup_auto_reconnect)

    def _run_startup_auto_reconnect(self) -> None:
        self.startup_reconnect_after_id = None
        if self.current_session_state == "running":
            return
        if not self.config.sensors:
            return
        if self._all_selected_sensors_active():
            self._cancel_startup_auto_reconnect()
            return
        if self.startup_reconnect_attempt >= self.startup_reconnect_max_attempts:
            self.status_var.set("Auto reconnect stopped. Sensors not active yet.")
            self._cancel_startup_auto_reconnect()
            return
        if not self.refresh_in_progress:
            self.startup_reconnect_attempt += 1
            self.refresh_selected_sensors()
        self.startup_reconnect_after_id = self.root.after(
            self.startup_reconnect_interval_ms,
            self._run_startup_auto_reconnect,
        )

    def _cancel_startup_auto_reconnect(self, clear_attempts: bool = True) -> None:
        if self.startup_reconnect_after_id is not None:
            try:
                self.root.after_cancel(self.startup_reconnect_after_id)
            except tk.TclError:
                pass
            self.startup_reconnect_after_id = None
        if clear_attempts:
            self.startup_reconnect_attempt = 0

    def _all_selected_sensors_active(self) -> bool:
        if not self.config.sensors:
            return False
        for role in self.config.sensors:
            if self.sensor_activity.get(role) != "active":
                return False
        return True

    def _handle_scan_results(self, transport: str, devices: list[DiscoveredSensor]) -> None:
        devices_by_id = {device.identifier: device for device in devices if device.transport == transport}
        self.last_scan_devices_by_transport[transport] = devices_by_id
        for role, binding in self.config.sensors.items():
            if binding.transport != transport:
                continue
            if binding.identifier in devices_by_id:
                self.sensor_activity[role] = "active"
                self.sensor_last_seen_names[role] = devices_by_id[binding.identifier].name
            else:
                self.sensor_activity[role] = "missing"
        self._handle_sensor_update()

    @staticmethod
    def _format_sensor_state(state: str) -> str:
        labels = {
            "active": "active now",
            "missing": "not found",
            "unknown": "unknown status",
            "unverified": "not verified in this session",
        }
        return labels.get(state, state)

    def _set_sensor_dot(self, role: str, state: str) -> None:
        colors = {
            "active": "#39b36b",
            "missing": "#d95c5c",
            "unverified": "#d9b43b",
            "unknown": "#d8d8d8",
            "none": "#d8d8d8",
        }
        canvas = getattr(self, f"_sensor_canvas_{role}", None)
        dot = self.sensor_status_dots.get(role)
        if canvas is None or dot is None:
            return
        canvas.itemconfig(dot, fill=colors.get(state, "#d8d8d8"))

    def _on_close(self) -> None:
        self._cancel_delayed_start()
        self._cancel_startup_auto_reconnect()
        if self.summary_render_timer_id is not None:
            try:
                self.root.after_cancel(self.summary_render_timer_id)
            except tk.TclError:
                pass
            self.summary_render_timer_id = None
        self.pending_summary = None
        self.is_session_running = False
        self.session_started_at = None
        self._cancel_elapsed_tick()
        self._stop_source_with_timeout()
        try:
            save_app_config(self.config)
        finally:
            self.root.quit()
            self.root.destroy()

    def _stop_source_with_timeout(self, timeout_seconds: float = 1.5) -> None:
        stop_error: list[Exception] = []

        def _stop_worker() -> None:
            try:
                self.source.stop()
            except Exception as exc:
                stop_error.append(exc)

        thread = threading.Thread(target=_stop_worker, daemon=True)
        thread.start()
        thread.join(timeout=timeout_seconds)
        if thread.is_alive():
            self.status_var.set("Stopping sensors took too long. Closing anyway.")
            return
        if stop_error:
            self.status_var.set(f"Close warning: {stop_error[0]}")

    def _apply_window_appearance(self) -> None:
        compact = self.current_session_state == "running"
        background = self.config.inactive_background
        self.root.attributes("-alpha", 1.0)
        self.root.configure(bg=background)
        self.style.configure("Overlay.TFrame", background=background)
        self.style.configure("Overlay.TLabel", background=background, foreground="#111111")
        self.style.configure("OverlayTitle.TLabel", background=background, foreground="#111111")
        self.style.configure("OverlayMetric.TLabel", background=background, foreground="#111111")
        self.style.configure("OverlayHeader.TLabel", background=background, foreground="#111111")
        self.style.configure("OverlayValue.TLabel", background=background, foreground="#111111")
        self.style.configure("OverlayStatus.TLabel", background=background, foreground="#222222")
        self.light_canvas.configure(bg=background)
        for role in ("power", "heart_rate"):
            canvas = getattr(self, f"_sensor_canvas_{role}", None)
            if canvas is not None:
                canvas.configure(bg=background)

        if compact:
            for widget in (
                self.title_row,
                self.config_row,
                self.sensor_info_frame,
                self.status_label,
                self.footer_separator,
                self.footer_label,
            ):
                if widget.winfo_manager():
                    widget.pack_forget()
        else:
            self._restore_normal_layout()

        self._fit_root_to_content(compact)

    def _restore_normal_layout(self) -> None:
        for widget in (
            self.title_row,
            self.header_status_row,
            self.metrics_frame,
            self.controls,
            self.controls_separator,
            self.config_row,
            self.sensor_info_frame,
            self.status_label,
            self.footer_separator,
            self.footer_label,
        ):
            if widget.winfo_manager():
                widget.pack_forget()

        self.title_row.pack(fill=tk.X)
        self.header_status_row.pack(fill=tk.X, pady=(8, 4))
        self.metrics_frame.pack(fill=tk.X)
        self.controls.pack(fill=tk.X, pady=(14, 0))
        self.controls_separator.pack(fill=tk.X, pady=(10, 8))
        self.config_row.pack(fill=tk.X, pady=(0, 0))
        self.sensor_info_frame.pack(fill=tk.X, pady=(8, 0))
        self.status_label.pack(anchor=tk.W, fill=tk.X, pady=(12, 0))
        self.footer_separator.pack(fill=tk.X, pady=(8, 6))
        self.footer_label.pack(anchor=tk.E, fill=tk.X, pady=(6, 0))

    def _calculate_compact_height(self) -> int:
        self.root.update_idletasks()
        row_heights = 0
        for widget in self.metrics_frame.winfo_children():
            if widget.winfo_manager():
                row_heights += widget.winfo_reqheight() + 4
        header_height = self.header_status_row.winfo_reqheight() + 8
        controls_height = self.controls.winfo_reqheight() + 14
        frame_padding = 32
        base_height = max(
            self._scaled_int(self.base_compact_minsize[1]),
            header_height + row_heights + controls_height + frame_padding,
        )
        return base_height

    def _fit_root_to_content(self, compact: bool) -> None:
        self.root.update_idletasks()
        if compact:
            base_width = self._scaled_int(self.base_compact_size[0])
            base_height = self._calculate_compact_height()
            min_width = self._scaled_int(self.base_compact_minsize[0])
            min_height = self._scaled_int(self.base_compact_minsize[1])
        else:
            base_width = self._scaled_int(self.base_normal_size[0])
            base_height = self._scaled_int(self.base_normal_size[1])
            min_width = self._scaled_int(self.base_normal_minsize[0])
            min_height = self._scaled_int(self.base_normal_minsize[1])

        if self.config.show_adjusted_wkg_column:
            base_width = max(base_width, self._scaled_int(392))
            min_width = max(min_width, self._scaled_int(360))

        req_width = self.main_frame.winfo_reqwidth() + 24
        req_height = self.main_frame.winfo_reqheight() + 24
        screen_width = max(min_width, self.root.winfo_screenwidth() - 80)
        screen_height = max(min_height, self.root.winfo_screenheight() - 80)

        width = max(min_width, base_width, req_width)
        height = max(min_height, base_height, req_height)
        width = min(width, screen_width)
        height = min(height, screen_height)

        self.root.geometry(f"{width}x{height}")
        self.root.minsize(min_width, min_height)

    def _apply_ui_scale(self) -> None:
        scale_percent = max(50, min(150, int(self.config.ui_scale_percent)))
        self.config.ui_scale_percent = scale_percent
        self.ui_scale_factor = scale_percent / 100
        self.root.tk.call("tk", "scaling", self.base_tk_scaling * self.ui_scale_factor)
        self.font_title.configure(size=self._scaled_int(self.base_font_sizes["title"]))
        self.font_value.configure(size=self._scaled_int(self.base_font_sizes["value"]))
        self.font_status.configure(size=self._scaled_int(self.base_font_sizes["status"]))
        self.font_default.configure(size=self._scaled_int(self.base_font_sizes["default"]))
        self.font_header.configure(size=self._scaled_int(self.base_font_sizes["header"]))
        self.font_button.configure(size=self._scaled_int(self.base_font_sizes["button"]))
        self.font_delay_title.configure(size=self._scaled_int(self.base_font_sizes["delay_title"]))
        self.font_delay_overlay.configure(size=self._scaled_int(self.base_font_sizes["delay_overlay"]))
        if hasattr(self, "style"):
            self._configure_styles()

    def _scaled_int(self, value: int) -> int:
        return max(1, int(round(value * self.ui_scale_factor)))

    def _current_elapsed_seconds(self) -> int:
        if self.is_session_running and self.session_started_at is not None:
            current = datetime.now() - self.session_started_at
            return self.accumulated_elapsed_seconds + max(0, int(current.total_seconds()))
        return self.accumulated_elapsed_seconds

    def _schedule_elapsed_tick(self) -> None:
        self._cancel_elapsed_tick()
        self.elapsed_timer_id = self.root.after(1000, self._tick_elapsed)

    def _tick_elapsed(self) -> None:
        if self.pending_summary is not None:
            self._flush_summary_render()
        else:
            self._render_summary(self.aggregator.summary())
        if self.is_session_running:
            self.elapsed_timer_id = self.root.after(1000, self._tick_elapsed)
        else:
            self.elapsed_timer_id = None

    def _cancel_elapsed_tick(self) -> None:
        if self.elapsed_timer_id is not None:
            self.root.after_cancel(self.elapsed_timer_id)
            self.elapsed_timer_id = None

    def run(self) -> None:
        self.root.mainloop()


class SensorConfigWindow:
    def __init__(
        self,
        root: tk.Misc,
        config: AppConfig,
        discovery_service: SensorDiscoveryService,
        on_save: Callable[[], None],
        on_scan_complete: Callable[[str, list[DiscoveredSensor]], None],
        get_sensor_state: Callable[[str], str],
        get_live_value_hints: Callable[[], dict[str, str]],
    ) -> None:
        self.config = config
        self.discovery_service = discovery_service
        self.on_save = on_save
        self.on_scan_complete = on_scan_complete
        self.get_sensor_state = get_sensor_state
        self.get_live_value_hints = get_live_value_hints
        self.devices: list[DiscoveredSensor] = []
        self.last_scan_devices_by_transport: dict[str, dict[str, DiscoveredSensor]] = {}
        self.status_var = tk.StringVar(value="Choose transport and scan for sensors.")
        self.scan_in_progress = False
        self.scan_started_at = 0.0
        self.scan_stop_event: threading.Event | None = None
        self.scan_thread: threading.Thread | None = None
        self.transport_status: dict[str, bool] = {"ble": False, "ant": False}
        self.transport_status_dots: dict[str, int] = {}
        self.transport_status_canvases: dict[str, tk.Canvas] = {}
        self.transport_status_vars: dict[str, tk.StringVar] = {}
        self.transport_status_in_progress = False
        self.transport_status_after_id: str | None = None
        self.transport_status_refresh_ms = 5000
        self.live_values_by_identifier: dict[str, str] = {}
        self.live_value_in_progress = False
        self.live_value_after_id: str | None = None
        self.live_value_refresh_ms = 1000
        self.live_value_stop_event = threading.Event()
        self.live_value_cache: dict[str, tuple[str, float]] = {}
        self.live_value_cache_ttl_seconds = 10.0
        self.live_probe_next_index = 0
        self.live_probe_batch_size = 3

        self.window = tk.Toplevel(root)
        self.window.title("Sensors")
        self.window.geometry("760x640")
        self.window.minsize(620, 520)
        self.window.transient(root)
        self.window.grab_set()
        self.window.protocol("WM_DELETE_WINDOW", self._on_close)

        self._build_ui()
        self._fit_window_to_content()
        self._render_bindings()
        self._refresh_transport_status()
        self._refresh_live_values()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.window, padding=14)
        frame.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(frame)
        header.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(header, text="Sensors", font=("Segoe UI", 11, "bold")).pack(anchor=tk.W)
        ttk.Label(
            header,
            text="Double-click a discovered sensor to assign it to Power or Heart Rate.",
            font=("Segoe UI", 8),
        ).pack(anchor=tk.W, pady=(2, 0))

        transport_card = ttk.LabelFrame(frame, text="Connections", padding=10)
        transport_card.pack(fill=tk.X)
        transport_status_row = ttk.Frame(transport_card)
        transport_status_row.pack(fill=tk.X, pady=(0, 8))
        for transport, label in (("ble", "BLE"), ("ant", "ANT+")):
            group = ttk.Frame(transport_status_row)
            group.pack(anchor=tk.W, pady=(0, 4))
            canvas = tk.Canvas(group, width=14, height=14, highlightthickness=0, bd=0)
            canvas.pack(side=tk.LEFT)
            dot = canvas.create_oval(2, 2, 12, 12, fill="#d8d8d8", outline="")
            self.transport_status_dots[transport] = dot
            self.transport_status_canvases[transport] = canvas
            text_var = tk.StringVar(value=f"{label} (checking...)")
            self.transport_status_vars[transport] = text_var
            ttk.Label(group, textvariable=text_var).pack(side=tk.LEFT, padx=(4, 0))

        transport_row = ttk.Frame(transport_card)
        transport_row.pack(fill=tk.X)
        self.scan_button = ttk.Button(
            transport_row,
            text="Scan for sensors",
            command=self.scan,
        )
        self.scan_button.pack(side=tk.LEFT)
        self.refresh_button = ttk.Button(
            transport_row,
            text="Refresh sensors",
            command=self.refresh_selected_sensors,
        )
        self.refresh_button.pack(side=tk.LEFT, padx=(8, 0))
        self.stop_search_button = ttk.Button(
            transport_row,
            text="Stop search",
            command=self.stop_search,
            state=tk.DISABLED,
        )
        self.stop_search_button.pack(side=tk.LEFT, padx=(8, 0))

        actions_row = ttk.Frame(transport_card)
        actions_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(actions_row, text="Disconnect all", command=self.disconnect_all).pack(side=tk.LEFT)

        columns = ("transport", "type", "name", "value", "id", "details")
        devices_card = ttk.LabelFrame(frame, text="Discovered sensors", padding=10)
        devices_card.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        self.tree = ttk.Treeview(devices_card, columns=columns, show="headings", height=12)
        self.tree.heading("transport", text="Transport")
        self.tree.heading("type", text="Type")
        self.tree.heading("name", text="Sensor")
        self.tree.heading("value", text="Value")
        self.tree.heading("id", text="Identifier")
        self.tree.heading("details", text="Details")
        self.tree.column("transport", width=80, anchor=tk.W)
        self.tree.column("type", width=100, anchor=tk.W)
        self.tree.column("name", width=150, anchor=tk.W)
        self.tree.column("value", width=100, anchor=tk.W)
        self.tree.column("id", width=170, anchor=tk.W)
        self.tree.column("details", width=180, anchor=tk.W)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree.bind("<Double-1>", self._on_sensor_double_click)
        self.tree.bind("<MouseWheel>", self._on_tree_mousewheel)
        self.tree.bind("<Button-4>", self._on_tree_mousewheel)
        self.tree.bind("<Button-5>", self._on_tree_mousewheel)
        tree_scroll = ttk.Scrollbar(devices_card, orient=tk.VERTICAL, command=self.tree.yview)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        bindings_frame = ttk.LabelFrame(frame, text="Selected sensors", padding=10)
        bindings_frame.pack(fill=tk.X, pady=(10, 0))
        self.binding_vars: dict[str, tk.StringVar] = {}
        self.binding_status_dots: dict[str, int] = {}
        self.binding_dot_canvases: dict[str, tk.Canvas] = {}
        display_labels = {
            "power": "Power",
            "heart_rate": "Heart Rate",
        }
        for role in ("power", "heart_rate"):
            row = ttk.Frame(bindings_frame)
            row.pack(fill=tk.X, pady=2)
            canvas = tk.Canvas(row, width=14, height=14, highlightthickness=0, bd=0)
            canvas.pack(side=tk.LEFT, pady=(1, 0))
            dot = canvas.create_oval(2, 2, 12, 12, fill="#d8d8d8", outline="")
            self.binding_status_dots[role] = dot
            self.binding_dot_canvases[role] = canvas

            ttk.Label(row, text=f"{display_labels[role]}:", width=14).pack(side=tk.LEFT, padx=(6, 0))
            var = tk.StringVar(value="(none selected)")
            self.binding_vars[role] = var
            ttk.Label(row, textvariable=var).pack(side=tk.LEFT)

        ttk.Label(
            frame,
            textvariable=self.status_var,
            font=("Segoe UI", 8),
            wraplength=640,
            justify=tk.LEFT,
        ).pack(
            anchor=tk.W,
            fill=tk.X,
            pady=(10, 0),
        )

        note = (
            "Speed comes from Zwift on screen and will be a separate OCR source. "
            "This window only handles power and heart rate."
        )
        ttk.Label(frame, text=note, wraplength=640, justify=tk.LEFT, font=("Segoe UI", 8)).pack(
            anchor=tk.W,
            fill=tk.X,
            pady=(8, 0),
        )

        footer_row = ttk.Frame(frame)
        footer_row.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(footer_row, text="Close", command=self._on_close).pack(side=tk.RIGHT)

    def _fit_window_to_content(self) -> None:
        self.window.update_idletasks()
        screen_w = max(640, self.window.winfo_screenwidth() - 80)
        screen_h = max(520, self.window.winfo_screenheight() - 80)
        req_w = self.window.winfo_reqwidth() + 8
        req_h = self.window.winfo_reqheight() + 8
        width = max(620, min(screen_w, req_w))
        height = max(520, min(screen_h, req_h))
        self.window.geometry(f"{width}x{height}")

    def _ensure_height_fits_content(self) -> None:
        if not self.window.winfo_exists():
            return
        self.window.update_idletasks()
        current_width = max(620, self.window.winfo_width())
        current_height = max(520, self.window.winfo_height())
        screen_h = max(520, self.window.winfo_screenheight() - 80)
        req_h = self.window.winfo_reqheight() + 8
        target_height = max(520, min(screen_h, req_h))
        # Resize in both directions so the window can shrink back when content gets shorter.
        if abs(target_height - current_height) >= 6:
            self.window.geometry(f"{current_width}x{target_height}")

    def scan(self) -> None:
        if self.scan_in_progress:
            return
        # Start immediately in background to keep UI responsive.
        self._start_scan(["ble", "ant"], "Scanning for devices...")

    def refresh_selected_sensors(self) -> None:
        if self.scan_in_progress:
            return
        if not self.config.sensors:
            self.status_var.set("No sensors selected.")
            return
        if self._all_selected_sensors_active():
            self.status_var.set("All selected sensors are already active.")
            return
        transports = sorted({binding.transport for binding in self.config.sensors.values()})
        # Start immediately in background to keep UI responsive.
        self._start_scan(transports, "Refreshing selected sensors...")

    def stop_search(self) -> None:
        if not self.scan_in_progress or self.scan_stop_event is None:
            return
        self.scan_stop_event.set()
        self.status_var.set("Stopping search...")

    def _start_scan(self, transports: list[str], status_text: str) -> None:
        self.scan_in_progress = True
        self.scan_started_at = time.perf_counter()
        self.scan_stop_event = threading.Event()
        self.scan_button.config(state=tk.DISABLED)
        self.refresh_button.config(state=tk.DISABLED)
        self.stop_search_button.config(state=tk.NORMAL)
        self.status_var.set(status_text)
        self.scan_thread = threading.Thread(
            target=self._scan_worker,
            args=(transports,),
            daemon=True,
        )
        self.scan_thread.start()

    def _scan_worker(self, transports: list[str]) -> None:
        devices_by_transport: dict[str, list[DiscoveredSensor]] = {}
        error_message: str | None = None
        cancelled = False
        try:
            for transport in transports:
                if self.scan_stop_event is not None and self.scan_stop_event.is_set():
                    raise SensorScanCancelledError("Search stopped.")
                if transport == "ble":
                    # Progressive BLE scan: update list every ~2s while still scanning.
                    total_seconds = 12.0
                    chunk_seconds = 2.0
                    found_by_id: dict[str, DiscoveredSensor] = {}
                    started = time.monotonic()
                    while time.monotonic() - started < total_seconds:
                        if self.scan_stop_event is not None and self.scan_stop_event.is_set():
                            raise SensorScanCancelledError("Search stopped.")
                        remaining = max(0.0, total_seconds - (time.monotonic() - started))
                        if remaining <= 0:
                            break
                        discovered_chunk = self.discovery_service.scan(
                            transport,
                            stop_event=self.scan_stop_event,
                            scan_seconds=min(chunk_seconds, remaining),
                            allow_empty=True,
                        )
                        for device in discovered_chunk:
                            found_by_id[device.identifier] = device
                        merged_chunk = list(found_by_id.values())
                        try:
                            self.window.after(
                                0,
                                lambda transport=transport, discovered=merged_chunk: self._apply_partial_scan_result(
                                    transport,
                                    discovered,
                                ),
                            )
                        except tk.TclError:
                            return
                    discovered = list(found_by_id.values())
                else:
                    discovered = self.discovery_service.scan(
                        transport,
                        stop_event=self.scan_stop_event,
                        allow_empty=True,
                    )
                    try:
                        self.window.after(
                            0,
                            lambda transport=transport, discovered=discovered: self._apply_partial_scan_result(
                                transport,
                                discovered,
                            ),
                        )
                    except tk.TclError:
                        return

                devices_by_transport[transport] = discovered
        except SensorScanCancelledError:
            cancelled = True
            error_message = "Search stopped."
        except SensorDiscoveryError as exc:
            error_message = str(exc)
        except Exception as exc:
            error_message = f"Unexpected error: {exc}"

        try:
            self.window.after(
                0,
                lambda: self._finish_scan_batch(devices_by_transport, error_message, cancelled),
            )
        except tk.TclError:
            return

    def _apply_partial_scan_result(self, transport: str, devices: list[DiscoveredSensor]) -> None:
        if not self.window.winfo_exists() or not self.scan_in_progress:
            return
        self.last_scan_devices_by_transport[transport] = {
            device.identifier: device for device in devices if device.transport == transport
        }
        self.on_scan_complete(transport, devices)

        merged_devices: list[DiscoveredSensor] = []
        for candidate in ("ble", "ant"):
            known = self.last_scan_devices_by_transport.get(candidate, {})
            merged_devices.extend(known.values())
        self.devices = merged_devices
        self._render_devices()
        self._render_bindings()

        label = "BLE" if transport == "ble" else "ANT+"
        total = len(self.devices)
        self.status_var.set(f"Found {len(devices)} via {label}. Total so far: {total}.")
        self._refresh_live_values()
        self._ensure_height_fits_content()

    def _finish_scan_batch(
        self,
        devices_by_transport: dict[str, list[DiscoveredSensor]],
        error_message: str | None,
        cancelled: bool,
    ) -> None:
        if not self.window.winfo_exists():
            return
        elapsed = max(0.0, time.perf_counter() - self.scan_started_at)
        self.scan_in_progress = False
        self.scan_stop_event = None
        self.scan_thread = None
        self.scan_button.config(state=tk.NORMAL)
        self.refresh_button.config(state=tk.NORMAL)
        self.stop_search_button.config(state=tk.DISABLED)
        merged_devices: list[DiscoveredSensor] = []
        for transport in ("ble", "ant"):
            merged_devices.extend(devices_by_transport.get(transport, []))
        self.devices = merged_devices

        for transport, devices in devices_by_transport.items():
            self.last_scan_devices_by_transport[transport] = {
                device.identifier: device for device in devices if device.transport == transport
            }
            self.on_scan_complete(transport, devices)

        self._render_devices()
        self._render_bindings()
        self._refresh_live_values()

        if cancelled:
            self.status_var.set(f"Search stopped. Time: {elapsed:.1f}s.")
            self._ensure_height_fits_content()
            return

        if error_message is not None:
            self.status_var.set(f"{error_message} Scan time: {elapsed:.1f}s.")
            self._refresh_transport_status()
            self._ensure_height_fits_content()
            return

        if len(devices_by_transport) == 1:
            transport = next(iter(devices_by_transport))
            label = TRANSPORTS.get(transport, transport)
            count = len(devices_by_transport[transport])
            self.status_var.set(f"Found {count} devices via {label}. Scan time: {elapsed:.1f}s.")
            self._ensure_height_fits_content()
            return

        total = sum(len(devices) for devices in devices_by_transport.values())
        self.status_var.set(
            f"Found {total} devices across {len(devices_by_transport)} transports. Scan time: {elapsed:.1f}s."
        )
        self._refresh_transport_status()
        self._ensure_height_fits_content()

    def _render_devices(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        sorted_devices = sorted(
            self.devices,
            key=lambda device: (
                0 if device.transport == "ble" else 1,
                device.name.lower(),
                device.identifier,
            ),
        )
        self.devices = sorted_devices
        for index, device in enumerate(sorted_devices):
            self.tree.insert(
                "",
                tk.END,
                iid=str(index),
                values=(
                    self._transport_short_label(device.transport),
                    device.sensor_type,
                    device.name,
                    self.live_values_by_identifier.get(device.identifier, "-"),
                    device.identifier,
                    device.details,
                ),
            )

    def _on_tree_mousewheel(self, event: tk.Event[tk.Misc]) -> str:
        first, last = self.tree.yview()
        if first <= 0.0 and last >= 1.0:
            return "break"
        delta = 0
        if hasattr(event, "delta") and event.delta:
            delta = -int(event.delta / 120)
        elif getattr(event, "num", None) == 4:
            delta = -1
        elif getattr(event, "num", None) == 5:
            delta = 1
        if delta != 0:
            self.tree.yview_scroll(delta, "units")
        return "break"

    def _refresh_live_values(self) -> None:
        if not self.window.winfo_exists():
            return
        if self.live_value_in_progress:
            self._schedule_live_value_refresh()
            return
        if not self.devices:
            self._schedule_live_value_refresh()
            return
        self.live_value_in_progress = True
        thread = threading.Thread(target=self._live_value_worker, daemon=True)
        thread.start()

    def _live_value_worker(self) -> None:
        try:
            values_by_identifier: dict[str, str] = {}
            devices_snapshot = list(self.devices)
            try:
                hint_values = self.get_live_value_hints()
            except Exception:
                hint_values = {}

            now = time.monotonic()
            role_hint_values: dict[str, str] = {}
            for role, binding in self.config.sensors.items():
                hinted_for_binding = hint_values.get(binding.identifier)
                if hinted_for_binding:
                    role_hint_values[role] = hinted_for_binding

            for device in devices_snapshot:
                hinted = hint_values.get(device.identifier)
                if not hinted:
                    device_name = (device.name or "").strip().lower()
                    for binding in self.config.sensors.values():
                        binding_name = (binding.name or "").strip().lower()
                        if (
                            binding.transport == device.transport
                            and binding_name
                            and binding_name == device_name
                        ):
                            hinted = hint_values.get(binding.identifier)
                            if hinted:
                                break
                if not hinted:
                    suggested_role = self._suggested_role_for_device(device)
                    if suggested_role is not None:
                        hinted = role_hint_values.get(suggested_role)
                if hinted:
                    values_by_identifier[device.identifier] = hinted
                    self.live_value_cache[device.identifier] = (hinted, now)
                    continue
                cached = self.live_value_cache.get(device.identifier)
                if cached is not None and now - cached[1] <= self.live_value_cache_ttl_seconds:
                    values_by_identifier[device.identifier] = cached[0]
                else:
                    values_by_identifier[device.identifier] = "-"

            probe_candidates = [device for device in devices_snapshot if device.identifier not in hint_values]
            if probe_candidates:
                probe_candidates.sort(
                    key=lambda device: (
                        0
                        if "heart" in (device.sensor_type or "").lower()
                        else (1 if "power" in (device.sensor_type or "").lower() else 2)
                    )
                )
                start = self.live_probe_next_index % len(probe_candidates)
                end = start + min(self.live_probe_batch_size, len(probe_candidates))
                batch = probe_candidates[start:end]
                if end > len(probe_candidates):
                    batch.extend(probe_candidates[: end - len(probe_candidates)])
                self.live_probe_next_index = (start + len(batch)) % len(probe_candidates)
                for device in batch:
                    if self.live_value_stop_event.is_set():
                        return
                    value = self._probe_live_value(device)
                    if value != "-":
                        values_by_identifier[device.identifier] = value
                        self.live_value_cache[device.identifier] = (value, time.monotonic())
            try:
                self.window.after(0, lambda: self._apply_live_values(values_by_identifier))
            except tk.TclError:
                self.live_value_in_progress = False
        except Exception:
            self.live_value_in_progress = False

    def _apply_live_values(self, values_by_identifier: dict[str, str]) -> None:
        if not self.window.winfo_exists():
            self.live_value_in_progress = False
            return
        self.live_values_by_identifier.update(values_by_identifier)
        self._render_devices()
        self.live_value_in_progress = False
        self._schedule_live_value_refresh()

    def _schedule_live_value_refresh(self) -> None:
        if not self.window.winfo_exists():
            return
        if self.live_value_after_id is not None:
            try:
                self.window.after_cancel(self.live_value_after_id)
            except tk.TclError:
                pass
            self.live_value_after_id = None
        self.live_value_after_id = self.window.after(self.live_value_refresh_ms, self._refresh_live_values)

    def _probe_live_value(self, device: DiscoveredSensor) -> str:
        if device.transport == "ble":
            return self._probe_ble_live_value(device)
        return "-"

    @staticmethod
    def _prepare_windows_ble_runtime() -> None:
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

    def _probe_ble_live_value(self, device: DiscoveredSensor) -> str:
        try:
            import asyncio
            from bleak import BleakClient, BleakScanner
        except Exception:
            return "-"

        self._prepare_windows_ble_runtime()

        async def _read_once() -> str:
            def _format_value(uuid: str, payload: bytearray | bytes) -> str | None:
                if uuid == HEART_RATE_MEASUREMENT_UUID:
                    heart_rate = self._parse_ble_heart_rate(payload)
                    if heart_rate is not None:
                        return f"{heart_rate} bpm"
                if uuid == CYCLING_POWER_MEASUREMENT_UUID:
                    power_watts = self._parse_ble_power(payload)
                    if power_watts is not None:
                        return f"{power_watts} W"
                return None

            async def _try_notify(client: object, uuid: str) -> str | None:
                loop = asyncio.get_running_loop()
                first_packet: asyncio.Future[bytearray | bytes] = loop.create_future()

                def _on_notify(_sender: object, data: bytearray) -> None:
                    if not first_packet.done():
                        first_packet.set_result(bytes(data))

                try:
                    await client.start_notify(uuid, _on_notify)
                except Exception:
                    return None

                try:
                    payload = await asyncio.wait_for(first_packet, timeout=2.0)
                except Exception:
                    payload = None
                finally:
                    try:
                        await client.stop_notify(uuid)
                    except Exception:
                        pass

                if payload is None:
                    return None
                return _format_value(uuid, payload)

            sensor_type = (device.sensor_type or "").lower()
            name_lower = (device.name or "").lower()
            preferred_uuids: list[str]
            if "heart" in sensor_type or any(token in name_lower for token in ("hr", "heart", "h10", "tickr", "pulse")):
                preferred_uuids = [HEART_RATE_MEASUREMENT_UUID]
            elif "power" in sensor_type or any(token in name_lower for token in ("power", "watt", "assioma", "favero", "rally", "vector")):
                preferred_uuids = [CYCLING_POWER_MEASUREMENT_UUID]
            else:
                preferred_uuids = [HEART_RATE_MEASUREMENT_UUID, CYCLING_POWER_MEASUREMENT_UUID]

            target = device.native_device if device.native_device is not None else device.identifier
            if target == device.identifier:
                try:
                    resolved_device = await BleakScanner.find_device_by_address(device.identifier, timeout=1.0)
                except Exception:
                    resolved_device = None
                if resolved_device is not None:
                    target = resolved_device

            try:
                async with BleakClient(target, timeout=1.6) as client:
                    for uuid in preferred_uuids:
                        notify_value = await _try_notify(client, uuid)
                        if notify_value is not None:
                            return notify_value
                        try:
                            raw = await asyncio.wait_for(client.read_gatt_char(uuid), timeout=0.7)
                            read_value = _format_value(uuid, raw)
                            if read_value is not None:
                                return read_value
                        except Exception:
                            pass
            except Exception:
                return "-"
            return "-"

        try:
            return asyncio.run(_read_once())
        except Exception:
            return "-"

    @staticmethod
    def _parse_ble_heart_rate(data: bytearray | bytes) -> int | None:
        if len(data) < 2:
            return None
        flags = data[0]
        is_uint16 = flags & 0x01
        if is_uint16:
            if len(data) < 3:
                return None
            value = int.from_bytes(data[1:3], "little")
        else:
            value = int(data[1])
        if value <= 0 or value > 250:
            return None
        return value

    @staticmethod
    def _parse_ble_power(data: bytearray | bytes) -> int | None:
        if len(data) < 4:
            return None
        value = int.from_bytes(data[2:4], "little", signed=True)
        if value < 0 or value > 3000:
            return None
        return value

    def _on_sensor_double_click(self, event: tk.Event[tk.Misc]) -> None:
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self.tree.selection_set(row_id)
        device = self.devices[int(row_id)]
        self._open_assign_dialog(device)

    def _assign_device_to_role(self, device: DiscoveredSensor, role: str) -> None:
        self.config.set_sensor(
            SensorBinding(
                role=role,
                name=device.name,
                identifier=device.identifier,
                transport=device.transport,
            )
        )
        # Sync main-window sensor state immediately using the latest discovered devices
        # so status color updates without waiting for another scan/update event.
        known_for_transport = list(self.last_scan_devices_by_transport.get(device.transport, {}).values())
        if known_for_transport:
            self.on_scan_complete(device.transport, known_for_transport)
        self._render_bindings()
        self.on_save()
        self.status_var.set(f"{device.name} is assigned to {SENSOR_ROLES[role].lower()}.")

    def _open_assign_dialog(self, device: DiscoveredSensor) -> None:
        dialog = tk.Toplevel(self.window)
        dialog.title("Assign sensor")
        dialog.transient(self.window)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=device.name, font=("Segoe UI", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(
            frame,
            text=f"{self._transport_short_label(device.transport)}  |  {device.identifier}",
            font=("Segoe UI", 8),
        ).pack(anchor=tk.W, pady=(2, 10))
        detected_role = self._suggested_role_for_device(device)
        detected_label = SENSOR_ROLES[detected_role] if detected_role is not None else "Unknown"
        ttk.Label(
            frame,
            text=f"Detected type: {device.sensor_type} ({detected_label})",
            font=("Segoe UI", 8),
        ).pack(anchor=tk.W, pady=(0, 8))
        ttk.Label(frame, text="Assign this sensor as:").pack(anchor=tk.W)

        buttons = ttk.Frame(frame)
        buttons.pack(anchor=tk.W, pady=(8, 0))

        def _choose(role: str) -> None:
            self._assign_device_to_role(device, role)
            dialog.destroy()

        top_buttons = ttk.Frame(buttons)
        top_buttons.pack(anchor=tk.W)
        ttk.Button(top_buttons, text="Power", command=lambda: _choose("power")).pack(side=tk.LEFT)
        ttk.Button(top_buttons, text="Heart Rate", command=lambda: _choose("heart_rate")).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="Close", command=dialog.destroy).pack(anchor=tk.W, pady=(8, 0))

        dialog.wait_window()

    @staticmethod
    def _suggested_role_for_device(device: DiscoveredSensor) -> str | None:
        sensor_type = (device.sensor_type or "").strip().lower()
        if "heart" in sensor_type:
            return "heart_rate"
        if "power" in sensor_type:
            return "power"
        return None

    def disconnect_all(self) -> None:
        self.config.sensors.clear()
        self._render_bindings()
        self.on_save()
        self.status_var.set("All sensors disconnected.")

    def _render_bindings(self) -> None:
        for role, label in SENSOR_ROLES.items():
            binding = self.config.get_sensor(role)
            if binding is None:
                self.binding_vars[role].set("(none selected)")
                self._set_binding_dot(role, "none")
                continue
            transport = self._transport_short_label(binding.transport)
            self.binding_vars[role].set(f"{binding.name} ({transport})")
            self._set_binding_dot(role, self.get_sensor_state(role))

    @staticmethod
    def _transport_short_label(transport: str) -> str:
        if transport == "ble":
            return "BLE"
        if transport == "ant":
            return "ANT+"
        return TRANSPORTS.get(transport, transport)

    def _set_binding_dot(self, role: str, state: str) -> None:
        colors = {
            "active": "#39b36b",
            "missing": "#d95c5c",
            "unverified": "#d9b43b",
            "unknown": "#d8d8d8",
            "none": "#d8d8d8",
        }
        canvas = self.binding_dot_canvases.get(role)
        dot = self.binding_status_dots.get(role)
        if canvas is None or dot is None:
            return
        canvas.itemconfig(dot, fill=colors.get(state, "#d8d8d8"))

    def _all_selected_sensors_active(self) -> bool:
        if not self.config.sensors:
            return False
        for role in self.config.sensors:
            if self.get_sensor_state(role) != "active":
                return False
        return True

    def _on_close(self) -> None:
        if self.scan_stop_event is not None:
            self.scan_stop_event.set()
        self.live_value_stop_event.set()
        if self.transport_status_after_id is not None:
            try:
                self.window.after_cancel(self.transport_status_after_id)
            except tk.TclError:
                pass
            self.transport_status_after_id = None
        if self.live_value_after_id is not None:
            try:
                self.window.after_cancel(self.live_value_after_id)
            except tk.TclError:
                pass
            self.live_value_after_id = None
        self.window.destroy()

    def _split_available_transports(self, transports: list[str]) -> tuple[list[str], list[str]]:
        available_transports: list[str] = []
        unavailable_messages: list[str] = []
        for transport in transports:
            available, message = self.discovery_service.check_transport_available(transport, force_refresh=True)
            if available:
                available_transports.append(transport)
            else:
                label = "BLE" if transport == "ble" else ("ANT+" if transport == "ant" else transport)
                unavailable_messages.append(f"{label}: {message}")
        return available_transports, unavailable_messages

    def _refresh_transport_status(self) -> None:
        if self.transport_status_in_progress:
            return
        if self.live_value_in_progress:
            self._schedule_transport_status_refresh()
            return
        self.transport_status_in_progress = True
        thread = threading.Thread(target=self._transport_status_worker, daemon=True)
        thread.start()

    def _transport_status_worker(self) -> None:
        status_map: dict[str, tuple[bool, str]] = {}
        for transport in ("ble", "ant"):
            status_map[transport] = self.discovery_service.check_transport_available(transport)
            try:
                self.window.after(
                    0,
                    lambda transport=transport, result=status_map[transport]: self._apply_transport_status_update(
                        transport,
                        result,
                    ),
                )
            except tk.TclError:
                self.transport_status_in_progress = False
                return
        try:
            self.window.after(0, self._finish_transport_status_refresh)
        except tk.TclError:
            self.transport_status_in_progress = False
            return

    def _apply_transport_status_update(self, transport: str, result: tuple[bool, str]) -> None:
        if not self.window.winfo_exists():
            return
        available, _message = result
        self.transport_status[transport] = available
        dot = self.transport_status_dots.get(transport)
        canvas = self.transport_status_canvases.get(transport)
        text_var = self.transport_status_vars.get(transport)
        if canvas is not None and dot is not None:
            canvas.itemconfig(dot, fill="#39b36b" if available else "#d8d8d8")
        if text_var is not None:
            label = "BLE" if transport == "ble" else "ANT+"
            state_label = "available" if available else "not available"
            text_var.set(f"{label} ({state_label})")

    def _finish_transport_status_refresh(self) -> None:
        if not self.window.winfo_exists():
            self.transport_status_in_progress = False
            return
        # Keep selected-sensor status dots synced with main window state.
        self._render_bindings()
        self._ensure_height_fits_content()
        self.transport_status_in_progress = False
        self._schedule_transport_status_refresh()

    def _apply_transport_status(self, status_map: dict[str, tuple[bool, str]]) -> None:
        if not self.window.winfo_exists():
            self.transport_status_in_progress = False
            return
        for transport, (available, _message) in status_map.items():
            self.transport_status[transport] = available
            dot = self.transport_status_dots.get(transport)
            canvas = self.transport_status_canvases.get(transport)
            text_var = self.transport_status_vars.get(transport)
            if canvas is not None and dot is not None:
                canvas.itemconfig(dot, fill="#39b36b" if available else "#d8d8d8")
            if text_var is not None:
                label = "BLE" if transport == "ble" else "ANT+"
                state_label = "available" if available else "not available"
                text_var.set(f"{label} ({state_label})")
        # Keep selected-sensor status dots synced with main window state.
        self._render_bindings()
        self._ensure_height_fits_content()
        self.transport_status_in_progress = False
        self._schedule_transport_status_refresh()

    def _schedule_transport_status_refresh(self) -> None:
        if not self.window.winfo_exists():
            return
        if self.transport_status_after_id is not None:
            try:
                self.window.after_cancel(self.transport_status_after_id)
            except tk.TclError:
                pass
            self.transport_status_after_id = None
        self.transport_status_after_id = self.window.after(
            self.transport_status_refresh_ms,
            self._refresh_transport_status,
        )


class SettingsWindow:
    def __init__(
        self,
        root: tk.Misc,
        config: AppConfig,
        on_save: Callable[[], None],
    ) -> None:
        self.root = root
        self.config = config
        self.on_save = on_save
        self.weight_var = tk.StringVar(value=config.rider_weight_input)
        self.profile_name_var = tk.StringVar(value=config.profile_name)
        self.profile_email_var = tk.StringVar(value=config.profile_email)
        self.topmost_var = tk.BooleanVar(value=config.always_on_top)
        self.power_display_var = tk.StringVar(value=f"{max(1, int(config.power_display_seconds))}s")
        self.wkg_decimals_var = tk.StringVar(value=str(2 if int(config.wkg_decimals) == 2 else 1))
        self.delayed_start_var = tk.StringVar(value=f"{max(10, int(config.delayed_start_seconds))}s")
        self.ui_scale_var = tk.IntVar(value=max(50, min(150, int(config.ui_scale_percent))))
        self.show_custom_avg_var = tk.BooleanVar(value=config.show_custom_avg_power)
        self.custom_avg_seconds_var = tk.StringVar(value=str(max(0, int(config.custom_avg_power_seconds))))
        self.show_session_avg_power_var = tk.BooleanVar(value=config.show_session_avg_power)
        self.show_avg_hr_var = tk.BooleanVar(value=config.show_avg_hr)
        self.show_avg_speed_var = tk.BooleanVar(value=config.show_avg_speed)
        self.show_adjusted_wkg_column_var = tk.BooleanVar(value=config.show_adjusted_wkg_column)
        adjusted_percent = 95 if int(config.adjusted_wkg_percent) == 95 else 90
        self.adjusted_wkg_percent_var = tk.StringVar(value=f"{adjusted_percent}%")
        configured_windows = {int(value) for value in config.avg_power_windows_seconds}
        self.avg_power_preset_vars: dict[int, tk.BooleanVar] = {
            seconds: tk.BooleanVar(value=seconds in configured_windows)
            for seconds in AVG_POWER_PRESET_SECONDS
        }
        self.inactive_background_var = tk.StringVar(value=config.inactive_background)

        self.window = tk.Toplevel(root)
        self.window.title("Settings")
        self.window.geometry("720x560")
        self.window.minsize(660, 520)
        self.window.resizable(True, True)
        self.window.transient(root)
        self.window.grab_set()

        container = ttk.Frame(self.window, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        scroll_canvas = tk.Canvas(container, highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=scroll_canvas.yview)
        scroll_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        frame = ttk.Frame(scroll_canvas)
        canvas_window = scroll_canvas.create_window((0, 0), window=frame, anchor=tk.NW)

        def _on_frame_configure(_event: tk.Event[tk.Misc] | None = None) -> None:
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))

        def _on_canvas_configure(event: tk.Event[tk.Misc]) -> None:
            scroll_canvas.itemconfigure(canvas_window, width=event.width)

        frame.bind("<Configure>", _on_frame_configure)
        scroll_canvas.bind("<Configure>", _on_canvas_configure)
        self._bind_canvas_mousewheel(scroll_canvas)

        top_frame = ttk.LabelFrame(frame, text="General", padding=8)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        top_frame.columnconfigure(1, weight=1)
        top_frame.columnconfigure(3, weight=1)

        profile_frame = ttk.LabelFrame(top_frame, text="Profile", padding=8)
        profile_frame.grid(row=0, column=0, columnspan=2, sticky=tk.NSEW, padx=(0, 16))
        profile_frame.columnconfigure(1, weight=1)
        ttk.Label(profile_frame, text="Weight* (kg)").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(profile_frame, textvariable=self.weight_var, width=10).grid(row=0, column=1, sticky=tk.W, padx=(8, 0))
        ttk.Label(profile_frame, text="Name").grid(row=1, column=0, sticky=tk.W, pady=(8, 0))
        ttk.Entry(profile_frame, textvariable=self.profile_name_var, width=26).grid(row=1, column=1, sticky=tk.W, padx=(8, 0), pady=(8, 0))
        ttk.Label(profile_frame, text="Email").grid(row=2, column=0, sticky=tk.W, pady=(8, 0))
        ttk.Entry(profile_frame, textvariable=self.profile_email_var, width=30).grid(row=2, column=1, sticky=tk.W, padx=(8, 0), pady=(8, 0))

        controls_frame = ttk.Frame(top_frame)
        controls_frame.grid(row=0, column=2, columnspan=2, sticky=tk.NSEW)
        ttk.Label(controls_frame, text="Power display").grid(row=0, column=0, sticky=tk.W)
        ttk.Combobox(
            controls_frame,
            textvariable=self.power_display_var,
            values=["1s", "3s", "5s"],
            width=10,
            state="readonly",
        ).grid(row=0, column=1, sticky=tk.W, padx=(8, 0))
        ttk.Label(controls_frame, text="W/kg decimals").grid(row=1, column=0, sticky=tk.W, pady=(8, 0))
        ttk.Combobox(
            controls_frame,
            textvariable=self.wkg_decimals_var,
            values=["1", "2"],
            width=10,
            state="readonly",
        ).grid(row=1, column=1, sticky=tk.W, padx=(8, 0), pady=(8, 0))

        ttk.Checkbutton(
            controls_frame,
            text="App always on top",
            variable=self.topmost_var,
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(8, 0))
        ttk.Label(controls_frame, text="Delayed start").grid(row=3, column=0, sticky=tk.W, pady=(8, 0))
        ttk.Combobox(
            controls_frame,
            textvariable=self.delayed_start_var,
            values=["10s", "30s", "60s"],
            width=10,
            state="readonly",
        ).grid(row=3, column=1, sticky=tk.W, padx=(8, 0), pady=(8, 0))

        self.ui_scale_label = ttk.Label(controls_frame, text=f"UI scaler: {self.ui_scale_var.get()}%")
        self.ui_scale_label.grid(row=4, column=0, sticky=tk.W, pady=(8, 0))
        ttk.Scale(
            controls_frame,
            from_=50,
            to=150,
            orient=tk.HORIZONTAL,
            variable=self.ui_scale_var,
            command=self._on_ui_scale_change,
            length=220,
        ).grid(row=4, column=1, sticky=tk.W, padx=(8, 0), pady=(8, 0))

        body = ttk.Frame(frame)
        body.pack(fill=tk.X)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)

        avg_power_frame = ttk.LabelFrame(body, text="Avg power rows", padding=8)
        avg_power_frame.grid(row=0, column=0, sticky=tk.NSEW, padx=(0, 6))
        ttk.Label(avg_power_frame, text="Preset windows").pack(anchor=tk.W)
        presets_grid = ttk.Frame(avg_power_frame)
        presets_grid.pack(fill=tk.X, pady=(4, 6))
        for index, seconds in enumerate(AVG_POWER_PRESET_SECONDS):
            label = self._format_duration_for_settings(seconds)
            ttk.Checkbutton(
                presets_grid,
                text=label,
                variable=self.avg_power_preset_vars[seconds],
            ).grid(row=index // 3, column=index % 3, sticky=tk.W, padx=(0, 14), pady=1)

        custom_row = ttk.Frame(avg_power_frame)
        custom_row.pack(fill=tk.X, pady=(4, 0))
        ttk.Checkbutton(
            custom_row,
            text="Custom",
            variable=self.show_custom_avg_var,
        ).pack(side=tk.LEFT)
        ttk.Entry(custom_row, textvariable=self.custom_avg_seconds_var, width=8).pack(side=tk.LEFT, padx=(8, 6))
        ttk.Label(custom_row, text="seconds").pack(side=tk.LEFT)

        right_col = ttk.Frame(body)
        right_col.grid(row=0, column=1, sticky=tk.NSEW, padx=(6, 0))

        visible_fields_frame = ttk.LabelFrame(right_col, text="Visible fields", padding=8)
        visible_fields_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Checkbutton(
            visible_fields_frame,
            text="Power (avg) session",
            variable=self.show_session_avg_power_var,
        ).pack(anchor=tk.W)
        ttk.Checkbutton(
            visible_fields_frame,
            text="HR / AVG / MAX",
            variable=self.show_avg_hr_var,
        ).pack(anchor=tk.W)
        ttk.Checkbutton(
            visible_fields_frame,
            text="Speed / AVG",
            variable=self.show_avg_speed_var,
        ).pack(anchor=tk.W)
        adjusted_row = ttk.Frame(visible_fields_frame)
        adjusted_row.pack(anchor=tk.W, fill=tk.X, pady=(4, 0))
        ttk.Checkbutton(
            adjusted_row,
            text="Extra W/kg column",
            variable=self.show_adjusted_wkg_column_var,
            command=self._toggle_adjusted_wkg_controls,
        ).pack(side=tk.LEFT)
        self.adjusted_wkg_percent_combo = ttk.Combobox(
            adjusted_row,
            textvariable=self.adjusted_wkg_percent_var,
            values=["90%", "95%"],
            width=6,
            state="readonly",
        )
        self.adjusted_wkg_percent_combo.pack(side=tk.LEFT, padx=(8, 0))

        appearance_frame = ttk.LabelFrame(right_col, text="Appearance", padding=8)
        appearance_frame.pack(fill=tk.X)
        ttk.Label(appearance_frame, text="Background color").pack(anchor=tk.W)
        background_row = ttk.Frame(appearance_frame)
        background_row.pack(anchor=tk.W, fill=tk.X, pady=(4, 8))
        ttk.Entry(background_row, textvariable=self.inactive_background_var, width=12).pack(side=tk.LEFT)
        ttk.Button(
            background_row,
            text="Choose",
            command=lambda: self.pick_color(self.inactive_background_var),
        ).pack(side=tk.LEFT, padx=(8, 0))

        buttons = ttk.Frame(self.window, padding=(12, 0, 12, 12))
        buttons.pack(fill=tk.X)
        ttk.Button(buttons, text="Save", command=self.save).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Reset settings", command=self.confirm_reset_defaults).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="Close", command=self.window.destroy).pack(side=tk.LEFT, padx=(8, 0))
        self._toggle_adjusted_wkg_controls()
        self._ensure_above_parent()

    def save(self, close_window: bool = True) -> bool:
        try:
            weight_input = self.weight_var.get().strip()
            if weight_input:
                weight_kg = float(weight_input.replace(",", "."))
                if weight_kg <= 0:
                    raise ValueError
            else:
                weight_kg = self.config.rider_weight_kg if self.config.rider_weight_kg > 0 else 100.0
            power_display_seconds = int(self.power_display_var.get().replace("s", ""))
            if power_display_seconds not in (1, 3, 5):
                raise ValueError
            wkg_decimals = int(self.wkg_decimals_var.get())
            if wkg_decimals not in (1, 2):
                raise ValueError
            delayed_start_seconds = int(self.delayed_start_var.get().replace("s", ""))
            if delayed_start_seconds not in (10, 30, 60):
                raise ValueError
            ui_scale_percent = int(self.ui_scale_var.get())
            if ui_scale_percent < 50 or ui_scale_percent > 150:
                raise ValueError
            adjusted_wkg_percent = int(self.adjusted_wkg_percent_var.get().replace("%", ""))
            if adjusted_wkg_percent not in (90, 95):
                raise ValueError
            custom_seconds = 0
            if self.show_custom_avg_var.get():
                custom_seconds = int(self.custom_avg_seconds_var.get())
                if custom_seconds <= 0:
                    raise ValueError
        except ValueError:
            messagebox.showinfo(
                "Settings",
                "Enter valid values for settings.",
            )
            return False

        selected_windows = [
            seconds
            for seconds, var in self.avg_power_preset_vars.items()
            if var.get()
        ]
        selected_windows = sorted(set(selected_windows))

        self.config.rider_weight_input = weight_input
        self.config.rider_weight_kg = weight_kg
        self.config.profile_name = self.profile_name_var.get().strip()
        self.config.profile_email = self.profile_email_var.get().strip()
        self.config.always_on_top = self.topmost_var.get()
        self.config.power_display_seconds = power_display_seconds
        self.config.wkg_decimals = wkg_decimals
        self.config.delayed_start_seconds = delayed_start_seconds
        self.config.ui_scale_percent = ui_scale_percent
        self.config.avg_power_windows_seconds = selected_windows
        self.config.show_custom_avg_power = self.show_custom_avg_var.get()
        self.config.custom_avg_power_seconds = custom_seconds
        self.config.show_session_avg_power = self.show_session_avg_power_var.get()
        self.config.show_avg_hr = self.show_avg_hr_var.get()
        self.config.show_max_hr = self.show_avg_hr_var.get()
        self.config.show_avg_speed = self.show_avg_speed_var.get()
        self.config.show_adjusted_wkg_column = self.show_adjusted_wkg_column_var.get()
        self.config.adjusted_wkg_percent = adjusted_wkg_percent
        self.config.inactive_background = self.inactive_background_var.get().strip() or "#f3f3f3"
        self.on_save()
        if close_window:
            self.window.destroy()
        return True

    def pick_color(self, variable: tk.StringVar) -> None:
        chosen = colorchooser.askcolor(color=variable.get(), parent=self.window, title="Choose background color")
        if chosen[1]:
            variable.set(chosen[1])

    @staticmethod
    def _format_duration_for_settings(seconds: int) -> str:
        if seconds % 60 == 0:
            return f"{seconds // 60}m"
        return f"{seconds}s"

    def _on_ui_scale_change(self, _value: str) -> None:
        self.ui_scale_label.config(text=f"UI scaler: {int(self.ui_scale_var.get())}%")

    def _bind_canvas_mousewheel(self, canvas: tk.Canvas) -> None:
        def _on_mousewheel(event: tk.Event[tk.Misc]) -> str:
            first, last = canvas.yview()
            if first <= 0.0 and last >= 1.0:
                return "break"
            delta = 0
            if hasattr(event, "delta") and event.delta:
                delta = -int(event.delta / 120)
            elif getattr(event, "num", None) == 4:
                delta = -1
            elif getattr(event, "num", None) == 5:
                delta = 1
            if delta != 0:
                canvas.yview_scroll(delta, "units")
            return "break"

        def _bind(_event: tk.Event[tk.Misc]) -> None:
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _unbind(_event: tk.Event[tk.Misc] | None = None) -> None:
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _bind)
        canvas.bind("<Leave>", _unbind)
        canvas.bind("<Destroy>", _unbind)

    def _toggle_adjusted_wkg_controls(self) -> None:
        state = "readonly" if self.show_adjusted_wkg_column_var.get() else "disabled"
        self.adjusted_wkg_percent_combo.configure(state=state)

    def confirm_reset_defaults(self) -> None:
        confirmed = self._confirm_reset_settings()
        if not confirmed:
            return
        self.reset_defaults()
        self.save(close_window=False)
        self._ensure_above_parent()
        self.window.after(50, self._ensure_above_parent)
        self.window.after(180, self._ensure_above_parent)

    def reset_defaults(self) -> None:
        self.weight_var.set("")
        self.profile_name_var.set("")
        self.profile_email_var.set("")
        self.topmost_var.set(True)
        self.power_display_var.set("3s")
        self.wkg_decimals_var.set("1")
        self.delayed_start_var.set("10s")
        self.ui_scale_var.set(100)
        self.ui_scale_label.config(text="UI scaler: 100%")
        for seconds, var in self.avg_power_preset_vars.items():
            var.set(seconds in (300, 1200))
        self.show_custom_avg_var.set(False)
        self.custom_avg_seconds_var.set("0")
        self.show_session_avg_power_var.set(True)
        self.show_avg_hr_var.set(True)
        self.show_avg_speed_var.set(True)
        self.show_adjusted_wkg_column_var.set(False)
        self.adjusted_wkg_percent_var.set("90%")
        self._toggle_adjusted_wkg_controls()
        self.inactive_background_var.set("#f3f3f3")

    def _confirm_reset_settings(self) -> bool:
        dialog = tk.Toplevel(self.window)
        dialog.title("Reset settings")
        dialog.transient(self.window)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            frame,
            text="Reset all settings to default values?\nChanges are applied and saved immediately.",
            justify=tk.LEFT,
        ).pack(anchor=tk.W)

        result = {"value": False}

        def _set_result(value: bool) -> None:
            result["value"] = value
            dialog.destroy()

        buttons = ttk.Frame(frame)
        buttons.pack(anchor=tk.E, pady=(12, 0))
        ttk.Button(buttons, text="Yes", command=lambda: _set_result(True)).pack(side=tk.LEFT)
        ttk.Button(buttons, text="No", command=lambda: _set_result(False)).pack(side=tk.LEFT, padx=(8, 0))
        dialog.protocol("WM_DELETE_WINDOW", lambda: _set_result(False))
        dialog.wait_window()
        return result["value"]

    def _ensure_above_parent(self) -> None:
        if not self.window.winfo_exists():
            return
        try:
            parent_topmost = bool(self.root.attributes("-topmost"))
        except tk.TclError:
            parent_topmost = False
        try:
            self.window.attributes("-topmost", parent_topmost)
        except tk.TclError:
            pass
        self.window.lift(self.root)
        self.window.focus_force()
        self.window.grab_set()


class ContactWindow:
    def __init__(self, root: tk.Misc, config: AppConfig) -> None:
        self.config = config
        self.window = tk.Toplevel(root)
        self.window.title("Contact")
        self.window.transient(root)
        self.window.grab_set()

        container = ttk.Frame(self.window, padding=12)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        frame = ttk.Frame(container)
        frame.grid(row=0, column=0)

        message = (
            "Need help or want to report an issue?\n"
            "Use Email to send a message or Discord to join the support server."
        )
        ttk.Label(frame, text=message, justify=tk.CENTER, wraplength=360).pack(anchor=tk.CENTER)

        buttons = ttk.Frame(frame)
        buttons.pack(anchor=tk.CENTER, pady=(14, 0))
        ttk.Button(buttons, text="Email", command=self.open_email_form).pack(side=tk.LEFT)
        ttk.Button(buttons, text="Discord", command=self.open_discord).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(buttons, text="Close", command=self.window.destroy).pack(side=tk.LEFT, padx=(8, 0))

        self.window.update_idletasks()
        width = max(360, frame.winfo_reqwidth() + 36)
        height = max(170, frame.winfo_reqheight() + 36)
        self.window.geometry(f"{width}x{height}")
        self.window.minsize(width, height)

    def open_email_form(self) -> None:
        EmailContactWindow(
            self.window,
            self.config,
            default_name=self.config.profile_name,
            default_email=self.config.profile_email,
        )

    def open_discord(self) -> None:
        opened = webbrowser.open(DISCORD_SERVER_URL)
        if not opened:
            messagebox.showwarning(
                "Contact",
                f"Could not open browser automatically.\nOpen this link manually:\n{DISCORD_SERVER_URL}",
                parent=self.window,
            )


class EmailContactWindow:
    def __init__(
        self,
        root: tk.Misc,
        config: AppConfig,
        default_name: str = "",
        default_email: str = "",
    ) -> None:
        self.config = config
        self.window = tk.Toplevel(root)
        self.window.title("Send Email")
        self.window.geometry("520x360")
        self.window.minsize(480, 330)
        self.window.transient(root)
        self.window.grab_set()

        self.name_var = tk.StringVar(value=default_name.strip())
        self.email_var = tk.StringVar(value=default_email.strip())

        frame = ttk.Frame(self.window, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Name").pack(anchor=tk.W)
        self.name_entry = tk.Entry(frame, textvariable=self.name_var, width=34)
        self.name_entry.pack(anchor=tk.W, pady=(4, 8))

        ttk.Label(frame, text="Email").pack(anchor=tk.W)
        self.email_entry = tk.Entry(frame, textvariable=self.email_var, width=40)
        self.email_entry.pack(anchor=tk.W, pady=(4, 8))

        self._set_entry_valid(self.name_entry)
        self._set_entry_valid(self.email_entry)
        self.name_entry.bind("<KeyRelease>", lambda _event: self._set_entry_valid(self.name_entry))
        self.email_entry.bind("<KeyRelease>", lambda _event: self._set_entry_valid(self.email_entry))

        ttk.Label(frame, text="Message").pack(anchor=tk.W)
        self.message_text = tk.Text(frame, height=10, wrap=tk.WORD)
        self.message_text.pack(fill=tk.BOTH, expand=True, pady=(4, 8))

        buttons = ttk.Frame(frame)
        buttons.pack(fill=tk.X)
        self.send_button = ttk.Button(buttons, text="Send", command=self.send_email)
        self.send_button.pack(side=tk.LEFT)
        ttk.Button(buttons, text="Close", command=self.window.destroy).pack(side=tk.LEFT, padx=(8, 0))

    def send_email(self) -> None:
        sender_email = self.email_var.get().strip()
        sender_name = self.name_var.get().strip()
        message = self.message_text.get("1.0", tk.END).strip()

        missing_fields: list[str] = []
        if not sender_name:
            missing_fields.append("name")
            self._set_entry_invalid(self.name_entry)
        else:
            self._set_entry_valid(self.name_entry)

        if not sender_email:
            missing_fields.append("email")
            self._set_entry_invalid(self.email_entry)
        else:
            self._set_entry_valid(self.email_entry)

        if missing_fields:
            messagebox.showinfo("Email", "Please fill in Name and Email.", parent=self.window)
            return

        if not message:
            messagebox.showinfo("Email", "Please enter a message.", parent=self.window)
            return

        subject = "Zwift Overlay Contact"
        body_parts = [
            f"From: {sender_name or '(not provided)'}",
            f"Email: {sender_email}",
            "",
            message,
        ]
        body = "\n".join(body_parts)
        smtp_config = self._resolve_smtp_config()
        if smtp_config is not None:
            self._send_via_smtp_async(subject, body, sender_email, smtp_config)
            return

        self._send_via_web_redirect(sender_email, subject, body)

    def _resolve_smtp_config(self) -> dict[str, str | int | bool] | None:
        # Primary path: app-owner managed secrets via environment variables.
        if all((APP_SMTP_HOST, APP_SMTP_USERNAME, APP_SMTP_PASSWORD, APP_SMTP_FROM_EMAIL)):
            return {
                "host": APP_SMTP_HOST,
                "port": max(1, APP_SMTP_PORT),
                "username": APP_SMTP_USERNAME,
                "password": APP_SMTP_PASSWORD,
                "from_email": APP_SMTP_FROM_EMAIL,
                "use_tls": APP_SMTP_USE_TLS,
            }
        # Backward-compatible fallback: legacy config-based SMTP.
        if self.config.smtp_enabled and all(
            (self.config.smtp_host, self.config.smtp_username, self.config.smtp_password, self.config.smtp_from_email)
        ):
            return {
                "host": self.config.smtp_host,
                "port": max(1, int(self.config.smtp_port)),
                "username": self.config.smtp_username,
                "password": self.config.smtp_password,
                "from_email": self.config.smtp_from_email,
                "use_tls": self.config.smtp_use_tls,
            }
        return None

    def _send_via_smtp_async(
        self,
        subject: str,
        body: str,
        reply_to_email: str,
        smtp_config: dict[str, str | int | bool],
    ) -> None:
        self.send_button.config(state=tk.DISABLED)
        self.send_button.configure(text="Sending...")

        def _worker() -> None:
            error: str | None = None
            try:
                msg = EmailMessage()
                msg["Subject"] = subject
                msg["From"] = str(smtp_config["from_email"])
                msg["To"] = CONTACT_EMAIL
                msg["Reply-To"] = reply_to_email
                msg.set_content(body)
                with smtplib.SMTP(str(smtp_config["host"]), int(smtp_config["port"]), timeout=10) as server:
                    server.ehlo()
                    if bool(smtp_config["use_tls"]):
                        server.starttls(context=ssl.create_default_context())
                        server.ehlo()
                    server.login(str(smtp_config["username"]), str(smtp_config["password"]))
                    server.send_message(msg)
            except Exception as exc:
                error = str(exc)
            try:
                self.window.after(0, lambda: self._finish_smtp_send(error))
            except tk.TclError:
                return

        threading.Thread(target=_worker, daemon=True).start()

    def _finish_smtp_send(self, error: str | None) -> None:
        if not self.window.winfo_exists():
            return
        self.send_button.config(state=tk.NORMAL)
        self.send_button.configure(text="Send")
        if error is None:
            messagebox.showinfo("Email", "Email sent successfully.", parent=self.window)
            self.window.destroy()
            return
        messagebox.showwarning(
            "Email",
            f"Direct send failed:\n{error}\n\nFalling back to webmail redirect.",
            parent=self.window,
        )
        sender_email = self.email_var.get().strip()
        sender_name = self.name_var.get().strip()
        message = self.message_text.get("1.0", tk.END).strip()
        fallback_body = "\n".join(
            [
                f"From: {sender_name or '(not provided)'}",
                f"Email: {sender_email}",
                "",
                message,
            ]
        )
        self._send_via_web_redirect(sender_email, "Zwift Overlay Contact", fallback_body)

    def _send_via_web_redirect(self, sender_email: str, subject: str, body: str) -> None:
        target_label, webmail_url = self._build_webmail_compose_target(sender_email, subject, body)
        confirmed = self._confirm_redirect(target_label)
        if not confirmed:
            return

        if webmail_url is not None:
            opened = webbrowser.open(webmail_url)
            if opened:
                return

        mailto_url = f"mailto:{CONTACT_EMAIL}?subject={quote(subject)}&body={quote(body)}"
        opened = webbrowser.open(mailto_url)
        if opened:
            return

        messagebox.showwarning(
            "Email",
            f"Could not open email client.\nSend manually to: {CONTACT_EMAIL}",
            parent=self.window,
        )

    @staticmethod
    def _build_webmail_compose_target(sender_email: str, subject: str, body: str) -> tuple[str, str | None]:
        domain = sender_email.split("@")[-1].lower()
        encoded_to = quote(CONTACT_EMAIL)
        encoded_subject = quote(subject)
        encoded_body = quote(body)

        if domain in {"gmail.com", "googlemail.com"}:
            return "Gmail", (
                "https://mail.google.com/mail/?view=cm&fs=1"
                f"&to={encoded_to}&su={encoded_subject}&body={encoded_body}"
            )
        if domain in {"outlook.com", "hotmail.com", "live.com", "msn.com"}:
            return "Outlook Web", (
                "https://outlook.live.com/mail/0/deeplink/compose"
                f"?to={encoded_to}&subject={encoded_subject}&body={encoded_body}"
            )
        if domain in {"yahoo.com", "yahoo.se", "ymail.com"}:
            return "Yahoo Mail", (
                "https://compose.mail.yahoo.com/"
                f"?to={encoded_to}&subject={encoded_subject}&body={encoded_body}"
            )
        return "your default mail app", None

    @staticmethod
    def _set_entry_invalid(entry: tk.Entry) -> None:
        entry.configure(highlightthickness=1, highlightbackground="#c0392b", highlightcolor="#c0392b")

    @staticmethod
    def _set_entry_valid(entry: tk.Entry) -> None:
        entry.configure(highlightthickness=1, highlightbackground="#bfbfbf", highlightcolor="#2a7fff")

    def _confirm_redirect(self, target_label: str) -> bool:
        dialog = tk.Toplevel(self.window)
        dialog.title("Send Email")
        dialog.transient(self.window)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding=12)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            frame,
            text=f"You will be redirected to {target_label}.\nAre you sure?",
            justify=tk.LEFT,
        ).pack(anchor=tk.W)

        result = {"value": False}

        def _set_result(value: bool) -> None:
            result["value"] = value
            dialog.destroy()

        buttons = ttk.Frame(frame)
        buttons.pack(anchor=tk.E, pady=(12, 0))
        ttk.Button(buttons, text="Yes", command=lambda: _set_result(True)).pack(side=tk.LEFT)
        ttk.Button(buttons, text="No", command=lambda: _set_result(False)).pack(side=tk.LEFT, padx=(8, 0))
        dialog.protocol("WM_DELETE_WINDOW", lambda: _set_result(False))
        dialog.wait_window()
        return result["value"]

