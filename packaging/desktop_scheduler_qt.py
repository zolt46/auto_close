# -*- coding: utf-8 -*-
"""
새로운 데스크톱 자동 종료·재생 스케줄러 (Qt GUI)

기능 개요
---------
* 요일별 시간표, 음성파일, 자동/수동 지정, 사용 여부를 저장
* 휴일 단일·범위 지정, 재부팅 후에도 설정 유지
* 대상 프로그램 강제 종료 → 음성 재생 → 원격 PC 종료 → 로컬 PC 종료 순으로 실행
* 메인 창을 닫아도 백그라운드에서 동작하고, 트레이 아이콘으로 제어
* 자동 재생용 플레이리스트를 관리하고, 요일별 자동 할당 시 순환

주요 기술 요소
---------------
* PySide6 기반의 QML 느낌의 카드를 입힌 UI
* `ConfigManager`가 AppData(또는 사용자 홈)의 JSON 구성 파일을 관리
* `SchedulerEngine`이 별도 스레드에서 다음 실행을 감지하고 GUI 스레드에 신호 전달
* `AudioService`가 Qt Multimedia로 음성 파일을 재생하고 완료 시 후속 작업을 호출
* 각 편집 탭은 `LiveUpdateMixin`을 통해 변경 즉시 저장 및 프리뷰를 갱신
"""
from __future__ import annotations

import atexit
import hashlib
import html
import json
import os
import platform
import shutil
import socket
import subprocess
import sys
import threading
import time
from dataclasses import dataclass, asdict, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

import psutil

try:
    import paramiko  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    paramiko = None

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QFont, QIcon, QPalette, QColor
from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer

APP_NAME = "AutoClose Studio"
APP_VERSION = "2.1.2"
BUILD_DATE = "2025-11-06"
AUTHOR_NAME = "Zolt46 / PSW / Emanon108"
ORGANIZATION_NAME = "AutoClose Studio"
ORGANIZATION_DOMAIN = "autoclose.local"
DEFAULT_USER_PASSWORD = "0000"
DEFAULT_ADMIN_PASSWORD = "000000"

APP_DIR = Path(__file__).resolve().parent


def resource_path(relative: str) -> Path:
    """exe 패키징(PyInstaller 등) 이후에도 자원 경로를 안전하게 찾는다."""

    base_dir = Path(getattr(sys, "_MEIPASS", APP_DIR))  # type: ignore[attr-defined]
    return base_dir / relative


DEFAULT_ASSET_DIR = "assets"
DEFAULT_APP_ICON = resource_path(f"{DEFAULT_ASSET_DIR}/app_icon.ico")

PREFERRED_UI_FONTS = [
    "Noto Sans KR",
    "Malgun Gothic",
    "Apple SD Gothic Neo",
    "Pretendard",
    "Nanum Gothic",
    "Segoe UI",
    "Arial",
]


def _build_ui_font(point_size: int, weight: int) -> QFont:
    font = QFont()
    set_families = getattr(font, "setFamilies", None)
    if callable(set_families):
        set_families(PREFERRED_UI_FONTS)
    else:  # pragma: no cover - Qt < 6.2 fallback
        font.setFamily(PREFERRED_UI_FONTS[0])
    font.setPointSize(point_size)
    font.setWeight(weight)
    if weight >= QFont.Weight.DemiBold:
        font.setBold(True)
    hinting_pref = getattr(QFont, "HintingPreference", None)
    set_hinting = getattr(font, "setHintingPreference", None)
    if hinting_pref is not None and callable(set_hinting):
        set_hinting(hinting_pref.PreferFullHinting)
    font.setStyleStrategy(QFont.PreferAntialias)
    return font

def hash_password(raw: str) -> str:
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def verify_password(stored_hash: str, attempt: str) -> bool:
    if not stored_hash:
        return True
    try:
        return stored_hash == hash_password(attempt)
    except Exception:
        return False


def coerce_bool(value: object, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in {"1", "true", "yes", "y", "on"}:
            return True
        if lowered in {"0", "false", "no", "n", "off"}:
            return False
    if isinstance(value, (int, float)):
        return bool(value)
    return default

class ConfigLocator:
    """현재 설정 파일 위치를 추적하고 변경을 돕는 도우미."""

    def __init__(self) -> None:
        base_root = Path(
            os.environ.get("APPDATA")
            or os.environ.get("XDG_CONFIG_HOME")
            or Path.home() / ".config"
        )
        self._default_dir = base_root / "auto_close_studio"
        self._default_dir.mkdir(parents=True, exist_ok=True)
        self._pointer_file = self._default_dir / "storage_location.json"
        self._config_dir = self._load_pointer()
        self._config_dir.mkdir(parents=True, exist_ok=True)

    def _load_pointer(self) -> Path:
        if self._pointer_file.exists():
            try:
                data = json.loads(self._pointer_file.read_text(encoding="utf-8"))
                stored = data.get("path")
                if stored:
                    path = Path(stored).expanduser()
                    return path
            except Exception:
                pass
        return self._default_dir

    @property
    def config_dir(self) -> Path:
        return self._config_dir

    @property
    def config_file(self) -> Path:
        return self._config_dir / "settings.json"

    def change_dir(self, new_dir: Path) -> None:
        target = Path(new_dir).expanduser()
        target.mkdir(parents=True, exist_ok=True)
        try:
            resolved = target.resolve()
        except Exception:
            resolved = target
        self._config_dir = resolved
        payload = {"path": str(self._config_dir)}
        self._pointer_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


CONFIG_LOCATOR = ConfigLocator()

DAY_KEYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
DAY_LABEL = {
    "mon": "월요일",
    "tue": "화요일",
    "wed": "수요일",
    "thu": "목요일",
    "fri": "금요일",
    "sat": "토요일",
    "sun": "일요일",
}

DEFAULT_TARGETS = ["chrome.exe", "msedge.exe", "vlc.exe", "YouTube Music.exe"]
DEFAULT_REMOTE = [
    {"host": "192.168.0.31", "username": "admin", "password": "1234", "method": "ssh"},
]


@dataclass
class DaySchedule:
    enabled: bool = True
    time: str = "09:00"  # HH:MM
    auto_assign: bool = True
    audio_path: Optional[str] = None
    allow_remote: bool = True
    allow_local_shutdown: bool = True
    last_ran: Optional[str] = None  # yyyy-mm-dd

    def as_dict(self) -> Dict[str, object]:
        data = asdict(self)
        return data

    @staticmethod
    def from_dict(data: Dict[str, object]) -> "DaySchedule":
        base = DaySchedule()
        for key in ("enabled", "time", "auto_assign", "audio_path", "allow_remote", "allow_local_shutdown", "last_ran"):
            if key in data:
                setattr(base, key, data[key])
        return base


@dataclass
class SchedulerConfig:
    playlist: List[str] = field(default_factory=list)
    playlist_rotation: int = 0
    targets: List[str] = field(default_factory=lambda: DEFAULT_TARGETS.copy())
    remote_hosts: List[Dict[str, str]] = field(default_factory=lambda: DEFAULT_REMOTE.copy())
    enable_remote_shutdown: bool = True
    enable_local_shutdown: bool = True
    shutdown_delay: int = 5
    holidays_enabled: bool = True
    auto_skip_weekends: bool = True
    holidays: List[str] = field(default_factory=list)
    holiday_ranges: List[Dict[str, str]] = field(default_factory=list)
    holiday_labels: Dict[str, str] = field(default_factory=dict)
    start_with_os: bool = False
    theme_accent: str = "#2A5CAA"
    header_logo_path: Optional[str] = None
    audio_volume: float = 0.9
    shutdown_logs: List[Dict[str, str]] = field(default_factory=list)
    user_password_hash: str = field(default_factory=lambda: hash_password(DEFAULT_USER_PASSWORD))
    admin_password_hash: str = field(default_factory=lambda: hash_password(DEFAULT_ADMIN_PASSWORD))
    days: Dict[str, DaySchedule] = field(
        default_factory=lambda: {key: DaySchedule(enabled=(key not in {"sat", "sun"})) for key in DAY_KEYS}
    )

    def as_dict(self) -> Dict[str, object]:
        data = asdict(self)
        data["days"] = {k: v.as_dict() for k, v in self.days.items()}
        return data

    @staticmethod
    def from_dict(data: Dict[str, object]) -> "SchedulerConfig":
        base = SchedulerConfig()
        for key in (
            "playlist",
            "playlist_rotation",
            "targets",
            "remote_hosts",
            "enable_remote_shutdown",
            "enable_local_shutdown",
            "shutdown_delay",
            "holidays_enabled",
            "auto_skip_weekends",
            "holidays",
            "holiday_ranges",
            "holiday_labels",
            "start_with_os",
            "theme_accent",
            "audio_volume",
            "header_logo_path",
            "shutdown_logs",
            "user_password_hash",
            "admin_password_hash",
        ):
            if key in data:
                setattr(base, key, data[key])
        days = {}
        for day_key, day_val in data.get("days", {}).items():
            days[day_key] = DaySchedule.from_dict(day_val)
        for missing in DAY_KEYS:
            days.setdefault(missing, DaySchedule(enabled=(missing not in {"sat", "sun"})))
        base.days = days
        base.auto_skip_weekends = coerce_bool(data.get("auto_skip_weekends", base.auto_skip_weekends), base.auto_skip_weekends)
        if not isinstance(base.holiday_labels, dict):
            base.holiday_labels = {}
        if not isinstance(base.user_password_hash, str) or not base.user_password_hash:
            base.user_password_hash = hash_password(DEFAULT_USER_PASSWORD)
        if not isinstance(base.admin_password_hash, str) or not base.admin_password_hash:
            base.admin_password_hash = hash_password(DEFAULT_ADMIN_PASSWORD)
        if base.header_logo_path and not isinstance(base.header_logo_path, str):
            base.header_logo_path = None
        try:
            base.audio_volume = float(base.audio_volume)
        except (TypeError, ValueError):
            base.audio_volume = 0.9
        base.audio_volume = max(0.0, min(1.0, base.audio_volume))
        return base

def is_holiday(cfg: SchedulerConfig, target: date) -> bool:
    if cfg.auto_skip_weekends and target.weekday() >= 5:
        return True
    if target.isoformat() in cfg.holidays:
        return True
    for rng in cfg.holiday_ranges:
        try:
            start = datetime.strptime(rng["start"], "%Y-%m-%d").date()
            end = datetime.strptime(rng["end"], "%Y-%m-%d").date()
        except Exception:
            continue
        if start <= target <= end:
            return True
    return False


def is_day_eligible(cfg: SchedulerConfig, day_cfg: DaySchedule, current_date: date) -> bool:
    if not day_cfg.enabled:
        return False
    if cfg.holidays_enabled and is_holiday(cfg, current_date):
        return False
    if day_cfg.last_ran == current_date.isoformat():
        return False
    return True


@dataclass
class UpcomingRun:
    when: datetime
    day_key: str
    audio_path: Optional[str]
    auto_assign: bool
    remote_allowed: bool
    local_allowed: bool


def compute_upcoming_runs(cfg: SchedulerConfig, horizon_days: int = 28, limit: Optional[int] = None) -> List[UpcomingRun]:
    now = datetime.now()
    playlist_len = len(cfg.playlist)
    rotation = cfg.playlist_rotation % max(1, playlist_len) if playlist_len else 0
    index = rotation
    runs: List[UpcomingRun] = []
    for offset in range(horizon_days):
        current = now + timedelta(days=offset)
        day_key = DAY_KEYS[current.weekday()]
        day_cfg = cfg.days.get(day_key)
        if not day_cfg or not is_day_eligible(cfg, day_cfg, current.date()):
            continue
        try:
            hh, mm = map(int, day_cfg.time.split(":"))
        except Exception:
            continue
        scheduled = current.replace(hour=hh, minute=mm, second=0, microsecond=0)
        if scheduled < now:
            continue
        if day_cfg.auto_assign and playlist_len:
            audio = cfg.playlist[index % playlist_len]
        else:
            audio = day_cfg.audio_path if day_cfg.audio_path else (cfg.playlist[index % playlist_len] if day_cfg.auto_assign and playlist_len else None)
        runs.append(
            UpcomingRun(
                when=scheduled,
                day_key=day_key,
                audio_path=audio,
                auto_assign=day_cfg.auto_assign,
                remote_allowed=cfg.enable_remote_shutdown and day_cfg.allow_remote,
                local_allowed=cfg.enable_local_shutdown and day_cfg.allow_local_shutdown,
            )
        )
        if day_cfg.auto_assign and playlist_len:
            index = (index + 1) % playlist_len
        if limit is not None and len(runs) >= limit:
            break
    return runs


def predict_playlist_for_day(cfg: SchedulerConfig, target_day: str) -> Optional[str]:
    if target_day not in DAY_KEYS:
        return None
    for run in compute_upcoming_runs(cfg, horizon_days=28):
        if run.day_key == target_day:
            return run.audio_path
    return None


class ConfigManager(QtCore.QObject):
    config_changed = Signal(SchedulerConfig)
    storage_dir_changed = Signal(str)

    def __init__(self) -> None:
        super().__init__()
        self._lock = threading.Lock()
        self.locator = CONFIG_LOCATOR
        self.config = self._load()
        atexit.register(self._flush_on_exit)

    def _load(self) -> SchedulerConfig:
        config_file = self.locator.config_file
        raw_data: Optional[Dict[str, object]] = None
        if config_file.exists():
            try:
                raw_data = json.loads(config_file.read_text(encoding="utf-8"))
                config = SchedulerConfig.from_dict(raw_data)
                if self._apply_migrations(config, raw_data):
                    try:
                        self._write(config)
                    except Exception as exc:  # pragma: no cover - non critical
                        print("[설정 마이그레이션 실패]", exc)
                return config
            except Exception as exc:  # pragma: no cover - fall back to default
                print("[설정 읽기 실패]", exc)
        config = SchedulerConfig()
        if self._apply_migrations(config, raw_data):
            try:
                self._write(config)
            except Exception as exc:  # pragma: no cover - non critical
                print("[기본 설정 저장 실패]", exc)
        return config

    def _apply_migrations(self, config: SchedulerConfig, data: Optional[Dict[str, object]]) -> bool:
        changed = False
        if data is None or "auto_skip_weekends" not in data:
            if config.auto_skip_weekends is not True:
                config.auto_skip_weekends = True
                changed = True
        return changed

    def _write(self, config: SchedulerConfig) -> None:
        config_file = self.locator.config_file
        config_file.parent.mkdir(parents=True, exist_ok=True)
        config_file.write_text(
            json.dumps(config.as_dict(), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def save(self) -> None:
        with self._lock:
            self._write(self.config)
            config = self.config
        self.config_changed.emit(config)

    def update(self, updater) -> None:
        with self._lock:
            updater(self.config)
            config = self.config
            self._write(config)
        self.config_changed.emit(config)

    def change_storage_dir(self, path: Path) -> None:
        with self._lock:
            self.locator.change_dir(path)
            self._write(self.config)
            config = self.config
        self.storage_dir_changed.emit(str(self.locator.config_dir))
        self.config_changed.emit(config)

    def storage_directory(self) -> Path:
        return self.locator.config_dir

    def _flush_on_exit(self) -> None:
        with self._lock:
            try:
                self._write(self.config)
            except Exception:
                pass

class SchedulerEngine(QtCore.QObject):
    schedule_triggered = Signal(str, str, bool, bool)  # day_key, audio_path, allow_remote, allow_local
    next_run_changed = Signal(object)

    def __init__(self, cfg_mgr: ConfigManager) -> None:
        super().__init__()
        self.cfg_mgr = cfg_mgr
        self._thread: Optional[threading.Thread] = None
        self._stop = threading.Event()

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, name="SchedulerEngine", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=1.5)

    def _loop(self) -> None:
        while not self._stop.is_set():
            self._compute_next_run()
            self._check_trigger()
            self._stop.wait(15)

    def _compute_next_run(self) -> None:
        cfg = self.cfg_mgr.config
        runs = compute_upcoming_runs(cfg, horizon_days=28, limit=1)
        self.next_run_changed.emit(runs[0] if runs else None)

    def _is_day_eligible(self, cfg: SchedulerConfig, day_cfg: DaySchedule, current_date: date) -> bool:
        return is_day_eligible(cfg, day_cfg, current_date)

    def _is_holiday(self, cfg: SchedulerConfig, target: date) -> bool:
        return is_holiday(cfg, target)

    def _check_trigger(self) -> None:
        now = datetime.now()
        cfg = self.cfg_mgr.config
        day_key = DAY_KEYS[now.weekday()]
        day_cfg = cfg.days[day_key]
        if not self._is_day_eligible(cfg, day_cfg, now.date()):
            return
        hh, mm = map(int, day_cfg.time.split(":"))
        target_time = now.replace(hour=hh, minute=mm, second=0, microsecond=0)
        if 0 <= (now - target_time).total_seconds() <= 30:
            audio_path = self._resolve_audio(cfg, day_cfg)
            self.schedule_triggered.emit(
                day_key,
                audio_path or "",
                cfg.enable_remote_shutdown and day_cfg.allow_remote,
                cfg.enable_local_shutdown and day_cfg.allow_local_shutdown,
            )
            def _mark():
                self.cfg_mgr.update(lambda c: c.days[day_key].__setattr__("last_ran", now.date().isoformat()))
            QtCore.QTimer.singleShot(0, _mark)

    def _resolve_audio(self, cfg: SchedulerConfig, day_cfg: DaySchedule) -> Optional[str]:
        if not day_cfg.auto_assign and day_cfg.audio_path:
            return day_cfg.audio_path
        if not cfg.playlist:
            return day_cfg.audio_path
        index = cfg.playlist_rotation % len(cfg.playlist)
        path = cfg.playlist[index]
        def _bump(config: SchedulerConfig) -> None:
            config.playlist_rotation = (config.playlist_rotation + 1) % max(1, len(config.playlist))
        QtCore.QTimer.singleShot(0, lambda: self.cfg_mgr.update(_bump))
        return path


class AudioService(QtCore.QObject):
    playback_started = Signal(str)
    playback_finished = Signal(str)

    def __init__(self) -> None:
        super().__init__()
        self.player = QMediaPlayer()
        self.audio_output = QAudioOutput()
        self.player.setAudioOutput(self.audio_output)
        self.player.mediaStatusChanged.connect(self._status_changed)
        self.player.playbackStateChanged.connect(self._playback_changed)
        self._current: Optional[str] = None
        self._volume: float = 0.9
        self.audio_output.setVolume(self._volume)

    def set_volume(self, value: float) -> None:
        try:
            volume = float(value)
        except (TypeError, ValueError):
            volume = 0.9
        self._volume = max(0.0, min(1.0, volume))
        self.audio_output.setVolume(self._volume)

    def play(self, path: str) -> None:
        if not path:
            self.playback_finished.emit("")
            return
        url = QtCore.QUrl.fromLocalFile(path)
        self.player.setSource(url)
        self.audio_output.setVolume(self._volume)
        self._current = path
        self.player.play()
        self.playback_started.emit(path)

    def stop(self) -> None:
        if self.player.playbackState() != QMediaPlayer.StoppedState:
            self.player.stop()

    def _status_changed(self, status: QMediaPlayer.MediaStatus) -> None:  # pragma: no cover - Qt callback
        if status == QMediaPlayer.InvalidMedia:
            path = self._current or ""
            self.playback_finished.emit(path)

    def _playback_changed(self, state: QMediaPlayer.PlaybackState) -> None:  # pragma: no cover - Qt callback
        if state == QMediaPlayer.StoppedState and self._current is not None:
            path = self._current
            self._current = None
            self.playback_finished.emit(path)


def terminate_programs(targets: List[str]) -> None:
    lowered = {t.lower() for t in targets}
    for proc in psutil.process_iter(["name"]):
        try:
            name = (proc.info.get("name") or "").lower()
            if name in lowered and proc.pid != os.getpid():
                proc.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue


def shutdown_remote(hosts: List[Dict[str, str]]) -> None:
    for host in hosts:
        method = host.get("method", "winrm")
        try:
            if method == "ssh" and paramiko:
                ssh = paramiko.SSHClient()
                ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                ssh.connect(
                    hostname=host.get("host"),
                    username=host.get("username"),
                    password=host.get("password"),
                    timeout=10,
                )
                ssh.exec_command("shutdown -h now")
                ssh.close()
            elif method == "winrm":
                os.system(f"shutdown /m \\{host.get('host')} /s /t 0")
        except Exception as exc:
            print(f"[원격 종료 실패] {host.get('host')}: {exc}")


def shutdown_local(delay: int) -> None:
    if sys.platform.startswith("win"):
        os.system(f"shutdown /s /t {max(0, delay)}")
    else:  # linux/mac
        time.sleep(max(0, delay))
        os.system("shutdown -h now")


def set_startup(start_with_os: bool) -> None:
    if not sys.platform.startswith("win"):
        return
    import winreg  # type: ignore

    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS) as key:
        if start_with_os:
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, f'"{sys.executable}" "{Path(__file__).resolve()}"')
        else:
            try:
                winreg.DeleteValue(key, APP_NAME)
            except FileNotFoundError:
                pass

class StyledToggle(QtWidgets.QCheckBox):
    """애니메이션 없이도 고대비 토글 모양을 제공하는 체크박스."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setCursor(Qt.PointingHandCursor)
        self.setTristate(False)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setFixedHeight(30)
        self.setStyleSheet(
            """
            QCheckBox {
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 50px;
                height: 26px;
                border-radius: 13px;
                border: 2px solid #90A4C5;
                background-color: #E3EAF6;
            }
            QCheckBox::indicator:checked {
                background-color: #3461C1;
                border-color: #244B9A;
            }
            QCheckBox::indicator:unchecked {
                background-color: #E3EAF6;
            }
            QCheckBox::indicator:disabled {
                background-color: #C7CFDD;
                border-color: #A5B1C8;
            }
            """
        )


def create_toggle_field(text: str, toggle: QtWidgets.QCheckBox) -> QtWidgets.QWidget:
    wrapper = QtWidgets.QWidget()
    wrapper.setProperty("toggleContainer", True)
    layout = QtWidgets.QHBoxLayout(wrapper)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(10)
    layout.addWidget(toggle, 0, Qt.AlignLeft)
    label = QtWidgets.QLabel(text)
    label.setProperty("toggleLabel", True)
    label.setBuddy(toggle)
    layout.addWidget(label, 0, Qt.AlignLeft)
    layout.addStretch(1)
    return wrapper


class FancyCard(QtWidgets.QFrame):
    def __init__(self, title: str, accent: str, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("FancyCard")
        self._accent = accent
        self.setAttribute(Qt.WA_StyledBackground, True)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)
        self.accent_line = QtWidgets.QFrame()
        self.accent_line.setFixedHeight(4)
        self.accent_line.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        layout.addWidget(self.accent_line)
        self.title = QtWidgets.QLabel(title)
        self.title.setProperty("role", "title")
        layout.addWidget(self.title)
        self.subtitle = QtWidgets.QLabel()
        self.subtitle.setProperty("role", "subtitle")
        layout.addWidget(self.subtitle)
        self.body_layout = QtWidgets.QVBoxLayout()
        layout.addLayout(self.body_layout)
        self._apply_styles()

    def _apply_styles(self) -> None:
        accent = QtGui.QColor(self._accent)
        border = accent.lighter(150).name()
        fill = accent.lighter(220).name()
        title_color = accent.darker(140).name()
        subtitle_color = accent.darker(110).name()
        self.setStyleSheet(
            f"""
            QFrame#FancyCard {{
                border-radius: 18px;
                background: {fill};
                border: 1px solid {border};
            }}
            QFrame#FancyCard QLabel[role="title"] {{
                color: {title_color};
                font-weight: 600;
                font-size: 18px;
            }}
            QFrame#FancyCard QLabel[role="subtitle"] {{
                color: {subtitle_color};
                font-size: 13px;
            }}
            """
        )
        self.accent_line.setStyleSheet(f"background-color: {accent.name()}; border-radius: 2px;")

    def set_subtitle(self, text: str) -> None:
        self.subtitle.setText(text)

    def set_accent(self, accent: str) -> None:
        self._accent = accent
        self._apply_styles()


class DayCard(FancyCard):
    changed = Signal(str)

    def __init__(self, day_key: str, cfg_mgr: ConfigManager, accent: str, parent=None) -> None:
        super().__init__(DAY_LABEL[day_key], accent, parent)
        self.day_key = day_key
        self.cfg_mgr = cfg_mgr
        self.set_subtitle("해당 요일의 실행 시간과 동작을 설정합니다")
        self._build_ui()
        self.sync_from_config()

    def _build_ui(self) -> None:
        layout = QtWidgets.QGridLayout()
        layout.setHorizontalSpacing(12)
        layout.setVerticalSpacing(10)
        enable_chk = StyledToggle()
        enable_chk.setToolTip("이 요일의 일정을 활성화하거나 비활성화합니다.")
        auto_chk = StyledToggle()
        time_edit = QtWidgets.QTimeEdit()
        time_edit.setDisplayFormat("HH:mm")
        manual_combo = QtWidgets.QComboBox()
        manual_combo.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToContents)
        manual_combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
        manual_combo.setEditable(False)
        manual_combo.addItem("선택 안 함", None)
        manual_combo.setCurrentIndex(0)
        remote_chk = StyledToggle()
        local_chk = StyledToggle()
        auto_hint = QtWidgets.QLabel()
        auto_hint.setProperty("role", "subtitle")
        auto_hint.setWordWrap(True)
        auto_hint.hide()
        layout.addWidget(create_toggle_field("사용", enable_chk), 0, 0)
        auto_chk.setToolTip("플레이리스트에서 자동으로 음성을 지정할지 여부입니다.")
        layout.addWidget(create_toggle_field("자동 음성", auto_chk), 0, 1)
        layout.addWidget(time_edit, 0, 2)
        remote_chk.setToolTip("일정 종료 시 원격 PC 종료 명령을 실행합니다.")
        layout.addWidget(create_toggle_field("원격 종료 허용", remote_chk), 1, 0)
        local_chk.setToolTip("일정 종료 시 이 PC를 종료합니다.")
        layout.addWidget(create_toggle_field("본체 종료", local_chk), 1, 1)
        layout.addWidget(manual_combo, 2, 0, 1, 3)
        layout.addWidget(auto_hint, 3, 0, 1, 3)
        container = QtWidgets.QWidget()
        container.setLayout(layout)
        self.body_layout.addWidget(container)
        self.enable_chk = enable_chk
        self.auto_chk = auto_chk
        self.time_edit = time_edit
        self.manual_combo = manual_combo
        self.remote_chk = remote_chk
        self.local_chk = local_chk
        self.auto_hint = auto_hint
        enable_chk.stateChanged.connect(lambda _: self._persist())
        auto_chk.stateChanged.connect(lambda _: self._persist())
        remote_chk.stateChanged.connect(lambda _: self._persist())
        local_chk.stateChanged.connect(lambda _: self._persist())
        time_edit.timeChanged.connect(lambda _: self._persist())
        manual_combo.currentIndexChanged.connect(lambda _: self._persist())
        auto_chk.stateChanged.connect(lambda _: self._update_mode())
        self._update_mode()

    def sync_from_config(self) -> None:
        cfg = self.cfg_mgr.config.days[self.day_key]
        self.enable_chk.blockSignals(True)
        self.enable_chk.setChecked(cfg.enabled)
        self.enable_chk.blockSignals(False)
        self.auto_chk.blockSignals(True)
        self.auto_chk.setChecked(cfg.auto_assign)
        self.auto_chk.blockSignals(False)
        hh, mm = map(int, cfg.time.split(":"))
        self.time_edit.blockSignals(True)
        self.time_edit.setTime(QtCore.QTime(hh, mm))
        self.time_edit.blockSignals(False)
        self._populate_manual_options(cfg.audio_path)
        self.remote_chk.blockSignals(True)
        self.remote_chk.setChecked(cfg.allow_remote)
        self.remote_chk.blockSignals(False)
        self.local_chk.blockSignals(True)
        self.local_chk.setChecked(cfg.allow_local_shutdown)
        self.local_chk.blockSignals(False)
        self._update_mode()

    def _populate_manual_options(self, selected_path: Optional[str]) -> None:
        playlist = self.cfg_mgr.config.playlist
        self.manual_combo.blockSignals(True)
        current_data = self.manual_combo.currentData()
        self.manual_combo.clear()
        self.manual_combo.addItem("선택 안 함", None)
        for path in playlist:
            self.manual_combo.addItem(Path(path).name, path)
        target = selected_path or current_data
        index = self.manual_combo.findData(target, Qt.UserRole)
        if index < 0:
            index = 0
        self.manual_combo.setCurrentIndex(index)
        self.manual_combo.blockSignals(False)
        self.manual_combo.setToolTip(self.manual_combo.currentData() or "")

    def _update_mode(self) -> None:
        is_auto = self.auto_chk.isChecked()
        has_playlist = self.manual_combo.count() > 1
        self.manual_combo.setEnabled(not is_auto and has_playlist)
        if not has_playlist:
            self.manual_combo.setToolTip("플레이리스트를 먼저 구성하세요")
        self.auto_hint.setVisible(is_auto)
        self._update_auto_hint()

    def _persist(self) -> None:
        def updater(cfg: SchedulerConfig) -> None:
            day_cfg = cfg.days[self.day_key]
            day_cfg.enabled = self.enable_chk.isChecked()
            day_cfg.auto_assign = self.auto_chk.isChecked()
            selected = self.manual_combo.currentData()
            day_cfg.audio_path = selected if selected else None
            day_cfg.time = self.time_edit.time().toString("HH:mm")
            day_cfg.allow_remote = self.remote_chk.isChecked()
            day_cfg.allow_local_shutdown = self.local_chk.isChecked()
        self.cfg_mgr.update(updater)
        self.changed.emit(self.day_key)
        self._update_mode()

    def _update_auto_hint(self) -> None:
        day_cfg = self.cfg_mgr.config.days[self.day_key]
        if not day_cfg.enabled:
            self.auto_hint.setText("일정이 비활성화되어 있습니다")
            return
        if not day_cfg.auto_assign:
            self.auto_hint.hide()
            return
        next_audio = predict_playlist_for_day(self.cfg_mgr.config, self.day_key)
        if next_audio:
            name = Path(next_audio).name
            self.auto_hint.setText(f"자동 음성: {name}")
            self.auto_hint.setToolTip(next_audio)
        else:
            self.auto_hint.setText("자동으로 사용할 음성이 없습니다. 플레이리스트를 확인하세요.")
            self.auto_hint.setToolTip("")

    preview_requested = Signal(str)
    stop_preview_requested = Signal()

class PlaylistPanel(FancyCard):
    def __init__(self, cfg_mgr: ConfigManager, accent: str, parent=None) -> None:
        super().__init__("플레이리스트", accent, parent)
        self.cfg_mgr = cfg_mgr
        self.set_subtitle("자동 음성 지정 시 순차 사용 · 미리 듣기 지원")
        self._preview_listeners: List[Callable[[str], None]] = []
        self._stop_preview_listeners: List[Callable[[], None]] = []
        layout = QtWidgets.QVBoxLayout()
        layout.setSpacing(12)
        self.list_widget = QtWidgets.QListWidget()
        self.list_widget.setAlternatingRowColors(True)
        self.list_widget.itemDoubleClicked.connect(lambda _: self._preview_selected())
        layout.addWidget(self.list_widget)
        btn_row = QtWidgets.QHBoxLayout()
        add_btn = QtWidgets.QPushButton("추가")
        remove_btn = QtWidgets.QPushButton("삭제")
        up_btn = QtWidgets.QPushButton("▲")
        down_btn = QtWidgets.QPushButton("▼")
        preview_btn = QtWidgets.QPushButton("미리 듣기")
        stop_btn = QtWidgets.QPushButton("정지")
        for btn in (add_btn, remove_btn, up_btn, down_btn, preview_btn, stop_btn):
            btn.setCursor(Qt.PointingHandCursor)
        btn_row.addWidget(add_btn)
        btn_row.addWidget(remove_btn)
        btn_row.addStretch(1)
        btn_row.addWidget(preview_btn)
        btn_row.addWidget(stop_btn)
        btn_row.addWidget(up_btn)
        btn_row.addWidget(down_btn)
        layout.addLayout(btn_row)
        volume_row = QtWidgets.QHBoxLayout()
        volume_label = QtWidgets.QLabel("재생 볼륨")
        volume_label.setProperty("role", "subtitle")
        self.volume_slider = QtWidgets.QSlider(Qt.Horizontal)
        self.volume_slider.setRange(0, 100)
        self.volume_slider.setPageStep(5)
        self.volume_slider.setSingleStep(1)
        self.volume_slider.setCursor(Qt.PointingHandCursor)
        self.volume_value = QtWidgets.QLabel()
        self.volume_value.setFixedWidth(60)
        volume_row.addWidget(volume_label)
        volume_row.addWidget(self.volume_slider, 1)
        volume_row.addWidget(self.volume_value, 0)
        layout.addLayout(volume_row)
        container = QtWidgets.QWidget()
        container.setLayout(layout)
        self.body_layout.addWidget(container)
        add_btn.clicked.connect(self._add_files)
        remove_btn.clicked.connect(self._remove_selected)
        up_btn.clicked.connect(lambda: self._move_selected(-1))
        down_btn.clicked.connect(lambda: self._move_selected(1))
        preview_btn.clicked.connect(self._preview_selected)
        stop_btn.clicked.connect(self._emit_stop_preview)
        self._volume_timer = QtCore.QTimer(self)
        self._volume_timer.setSingleShot(True)
        self._volume_timer.setInterval(300)
        self._volume_timer.timeout.connect(self._persist_volume)
        self.volume_slider.valueChanged.connect(self._on_volume_changed)
        self.refresh()

    def refresh(self) -> None:
        self.list_widget.clear()
        for path in self.cfg_mgr.config.playlist:
            item = QtWidgets.QListWidgetItem(Path(path).name)
            item.setData(Qt.UserRole, path)
            item.setToolTip(path)
            self.list_widget.addItem(item)
        self._sync_volume()

    def _add_files(self) -> None:
        paths, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "음성 파일 추가", str(Path.home()), "Audio Files (*.mp3 *.wav *.ogg)")
        if not paths:
            return
        def updater(cfg: SchedulerConfig) -> None:
            for path in paths:
                if path not in cfg.playlist:
                    cfg.playlist.append(path)
        self.cfg_mgr.update(updater)
        self.refresh()

    def _remove_selected(self) -> None:
        selected = self.list_widget.selectedItems()
        if not selected:
            return
        remove_paths = {item.data(Qt.UserRole) for item in selected}
        def updater(cfg: SchedulerConfig) -> None:
            cfg.playlist = [p for p in cfg.playlist if p not in remove_paths]
            cfg.playlist_rotation = 0
        self.cfg_mgr.update(updater)
        self.refresh()

    def _move_selected(self, direction: int) -> None:
        row = self.list_widget.currentRow()
        if row < 0:
            return
        target = row + direction
        if not (0 <= target < self.list_widget.count()):
            return
        def updater(cfg: SchedulerConfig) -> None:
            cfg.playlist[row], cfg.playlist[target] = cfg.playlist[target], cfg.playlist[row]
        self.cfg_mgr.update(updater)
        self.refresh()
        self.list_widget.setCurrentRow(target)

    def _preview_selected(self) -> None:
        item = self.list_widget.currentItem()
        if not item:
            show_info_message(self, "안내", "미리 듣기할 파일을 선택하세요.")
            return
        path = item.data(Qt.UserRole)
        if not path or not Path(path).exists():
            show_warning_message(self, "재생 불가", "파일을 찾을 수 없습니다. 경로를 확인해주세요.")
            return
        self._emit_preview(path)

    def add_preview_listener(self, callback: Callable[[str], None]) -> None:
        if callback not in self._preview_listeners:
            self._preview_listeners.append(callback)

    def add_stop_preview_listener(self, callback: Callable[[], None]) -> None:
        if callback not in self._stop_preview_listeners:
            self._stop_preview_listeners.append(callback)

    def _emit_preview(self, path: str) -> None:
        signal = getattr(self, "preview_requested", None)
        if signal is not None:
            signal.emit(path)
        for callback in list(self._preview_listeners):
            callback(path)

    def _emit_stop_preview(self) -> None:
        # 일부 PySide6 배포본에서는 사용자 정의 Signal 속성이 지연 초기화되면서
        # 객체 생성 직후에는 hasattr 체크가 필요할 수 있다. getattr을 사용해
        # 존재할 때만 emit을 호출해 예외를 방지한다.
        signal = getattr(self, "stop_preview_requested", None)
        if signal is not None:
            signal.emit()
        for callback in list(self._stop_preview_listeners):
            callback()

    def _on_volume_changed(self, value: int) -> None:
        self.volume_value.setText(f"{value}%")
        self._volume_timer.start()

    def _persist_volume(self) -> None:
        value = max(0, min(100, self.volume_slider.value())) / 100

        def updater(cfg: SchedulerConfig) -> None:
            cfg.audio_volume = value

        self.cfg_mgr.update(updater)

    def _sync_volume(self) -> None:
        target = int(round(self.cfg_mgr.config.audio_volume * 100))
        target = max(0, min(100, target))
        block = self.volume_slider.blockSignals(True)
        self.volume_slider.setValue(target)
        self.volume_slider.blockSignals(block)
        self.volume_value.setText(f"{target}%")

class AutoAssignmentPreviewCard(FancyCard):
    def __init__(self, cfg_mgr: ConfigManager, accent: str, parent=None) -> None:
        super().__init__("자동 음성 배정 미리보기", accent, parent)
        self.cfg_mgr = cfg_mgr
        self.set_subtitle("향후 실행 순서와 자동 지정 음성을 확인합니다")
        layout = QtWidgets.QVBoxLayout()
        layout.setSpacing(8)
        hint = QtWidgets.QLabel("플레이리스트와 요일 설정을 변경하면 자동으로 갱신됩니다")
        hint.setProperty("role", "subtitle")
        layout.addWidget(hint)
        self.table = QtWidgets.QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels([
            "날짜",
            "요일",
            "시간",
            "지정 방식",
            "음성",
        ])
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QtWidgets.QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QtWidgets.QHeaderView.Stretch)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.table.setAlternatingRowColors(True)
        layout.addWidget(self.table)
        wrapper = QtWidgets.QWidget()
        wrapper.setLayout(layout)
        self.body_layout.addWidget(wrapper)
        self.refresh()

    def refresh(self) -> None:
        cfg = self.cfg_mgr.config
        runs = compute_upcoming_runs(cfg, horizon_days=35, limit=14)
        self.table.clearContents()
        self.table.clearSpans()
        if runs:
            self.table.setRowCount(len(runs))
            for row, run in enumerate(runs):
                date_item = QtWidgets.QTableWidgetItem(run.when.strftime("%Y-%m-%d"))
                day_item = QtWidgets.QTableWidgetItem(DAY_LABEL.get(run.day_key, run.day_key))
                time_item = QtWidgets.QTableWidgetItem(run.when.strftime("%H:%M"))
                if run.auto_assign:
                    mode_text = "자동"
                    if not run.audio_path:
                        mode_text += " (대기)"
                else:
                    mode_text = "수동"
                mode_item = QtWidgets.QTableWidgetItem(mode_text)
                if run.audio_path:
                    audio_name = Path(run.audio_path).name
                elif run.auto_assign:
                    audio_name = "플레이리스트 없음"
                else:
                    audio_name = "지정된 파일 없음"
                audio_item = QtWidgets.QTableWidgetItem(audio_name)
                if run.audio_path:
                    audio_item.setToolTip(run.audio_path)
                for col, item in enumerate([date_item, day_item, time_item, mode_item, audio_item]):
                    item.setFlags(Qt.ItemIsEnabled)
                    self.table.setItem(row, col, item)
        else:
            self.table.setRowCount(1)
            empty_item = QtWidgets.QTableWidgetItem("예정된 실행이 없습니다")
            empty_item.setFlags(Qt.ItemIsEnabled)
            self.table.setItem(0, 0, empty_item)
            self.table.setSpan(0, 0, 1, 5)

class HolidayPanel(FancyCard):
    def __init__(self, cfg_mgr: ConfigManager, accent: str, parent=None) -> None:
        super().__init__("휴일 설정", accent, parent)
        self.cfg_mgr = cfg_mgr
        self.set_subtitle("지정된 날짜에는 스케줄이 실행되지 않습니다")
        layout = QtWidgets.QVBoxLayout()
        layout.setSpacing(10)
        toggle = StyledToggle()
        toggle.setToolTip("휴일에 등록된 날짜에는 일정이 실행되지 않습니다.")
        weekend_toggle = StyledToggle()
        weekend_toggle.setToolTip("토요일과 일요일을 자동으로 휴일로 처리합니다.")
        layout.addWidget(create_toggle_field("휴일 기능 사용", toggle))
        layout.addWidget(create_toggle_field("주말(토·일) 자동 제외", weekend_toggle))
        summary = QtWidgets.QLabel()
        summary.setProperty("role", "subtitle")
        layout.addWidget(summary)
        button_row = QtWidgets.QHBoxLayout()
        add_single_btn = QtWidgets.QPushButton("날짜 추가")
        add_range_btn = QtWidgets.QPushButton("기간 추가")
        add_weekend_btn = QtWidgets.QPushButton("주말 일괄 추가")
        import_btn = QtWidgets.QPushButton("ICS 가져오기")
        for btn in (add_single_btn, add_range_btn, add_weekend_btn, import_btn):
            btn.setCursor(Qt.PointingHandCursor)
            button_row.addWidget(btn)
        button_row.addStretch(1)
        layout.addLayout(button_row)
        self.single_list = QtWidgets.QListWidget()
        self.single_list.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        layout.addWidget(self.single_list)
        single_btns = QtWidgets.QHBoxLayout()
        self.single_delete_btn = QtWidgets.QPushButton("선택 삭제")
        self.single_delete_btn.setCursor(Qt.PointingHandCursor)
        single_btns.addWidget(self.single_delete_btn)
        single_btns.addStretch(1)
        layout.addLayout(single_btns)
        self.range_list = QtWidgets.QListWidget()
        self.range_list.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        layout.addWidget(self.range_list)
        range_btns = QtWidgets.QHBoxLayout()
        self.range_delete_btn = QtWidgets.QPushButton("선택 삭제")
        self.range_delete_btn.setCursor(Qt.PointingHandCursor)
        range_btns.addWidget(self.range_delete_btn)
        range_btns.addStretch(1)
        layout.addLayout(range_btns)
        container = QtWidgets.QWidget()
        container.setLayout(layout)
        self.body_layout.addWidget(container)
        toggle.stateChanged.connect(lambda _: self._persist())
        weekend_toggle.stateChanged.connect(lambda _: self._persist())
        add_single_btn.clicked.connect(self._add_single)
        add_range_btn.clicked.connect(self._add_range)
        add_weekend_btn.clicked.connect(self._add_weekend_range)
        import_btn.clicked.connect(self._import_ics)
        self.single_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.single_list.customContextMenuRequested.connect(lambda pos: self._context_remove(self.single_list, pos))
        self.range_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.range_list.customContextMenuRequested.connect(lambda pos: self._context_remove(self.range_list, pos))
        self.single_delete_btn.clicked.connect(lambda: self._remove_selected(self.single_list))
        self.range_delete_btn.clicked.connect(lambda: self._remove_selected(self.range_list))
        self.toggle = toggle
        self.weekend_toggle = weekend_toggle
        self.summary_label = summary
        self.refresh()

    def refresh(self) -> None:
        cfg = self.cfg_mgr.config
        self.toggle.blockSignals(True)
        self.toggle.setChecked(cfg.holidays_enabled)
        self.toggle.blockSignals(False)
        self.weekend_toggle.blockSignals(True)
        self.weekend_toggle.setChecked(cfg.auto_skip_weekends)
        self.weekend_toggle.blockSignals(False)
        singles = sorted(cfg.holidays)
        self.single_list.clear()
        for iso in singles:
            label = cfg.holiday_labels.get(iso, "")
            text = f"{iso} · {label}" if label else iso
            item = QtWidgets.QListWidgetItem(text)
            item.setData(Qt.UserRole, iso)
            if label:
                item.setToolTip(label)
            self.single_list.addItem(item)
        self.range_list.clear()
        for rng in cfg.holiday_ranges:
            text = f"{rng['start']} ~ {rng['end']}"
            item = QtWidgets.QListWidgetItem(text)
            item.setData(Qt.UserRole, (rng.get("start"), rng.get("end")))
            self.range_list.addItem(item)
        total_days = len(singles)
        total_ranges = len(cfg.holiday_ranges)
        weekend_text = "주말 자동 제외" if cfg.auto_skip_weekends else "주말 포함"
        self.summary_label.setText(f"단일 휴일 {total_days}개 · 기간 {total_ranges}건 · {weekend_text}")

    def _persist(self) -> None:
        enabled = self.toggle.isChecked()
        skip_weekends = self.weekend_toggle.isChecked()

        current = self.cfg_mgr.config
        if current.holidays_enabled == enabled and current.auto_skip_weekends == skip_weekends:
            return

        def updater(cfg: SchedulerConfig) -> None:
            cfg.holidays_enabled = enabled
            cfg.auto_skip_weekends = skip_weekends

        self.cfg_mgr.update(updater)

    def _add_single(self) -> None:
        dlg = QtWidgets.QCalendarWidget()
        dlg.setGridVisible(True)
        dialog = QtWidgets.QDialog(self)
        dialog.setWindowTitle("휴일 추가")
        layout = QtWidgets.QVBoxLayout(dialog)
        layout.addWidget(dlg)
        btn = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        layout.addWidget(btn)
        btn.accepted.connect(dialog.accept)
        btn.rejected.connect(dialog.reject)
        if dialog.exec() == QtWidgets.QDialog.Accepted:
            selected = dlg.selectedDate().toPython().isoformat()

            def updater(cfg: SchedulerConfig) -> None:
                if selected not in cfg.holidays:
                    cfg.holidays.append(selected)

            self.cfg_mgr.update(updater)
            self.refresh()

    def _add_range(self) -> None:
        dialog = DateRangeDialog(self)
        if dialog.exec() == QtWidgets.QDialog.Accepted:
            start, end = dialog.result_range

            def updater(cfg: SchedulerConfig) -> None:
                cfg.holiday_ranges.append({"start": start, "end": end})

            self.cfg_mgr.update(updater)
            self.refresh()

    def _add_weekend_range(self) -> None:
        dialog = DateRangeDialog(self)
        dialog.setWindowTitle("주말 일괄 추가")
        if dialog.exec() == QtWidgets.QDialog.Accepted:
            start_str, end_str = dialog.result_range
            try:
                start = datetime.strptime(start_str, "%Y-%m-%d").date()
                end = datetime.strptime(end_str, "%Y-%m-%d").date()
            except ValueError:
                return
            if end < start:
                return

            def updater(cfg: SchedulerConfig) -> None:
                current = start
                while current <= end:
                    if current.weekday() >= 5:
                        iso = current.isoformat()
                        if iso not in cfg.holidays:
                            cfg.holidays.append(iso)
                    current += timedelta(days=1)

            self.cfg_mgr.update(updater)
            self.refresh()

    def _import_ics(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "iCalendar 가져오기", str(Path.home()), "iCalendar (*.ics)")
        if not path:
            return
        try:
            try:
                text = Path(path).read_text(encoding="utf-8")
            except UnicodeDecodeError:
                text = Path(path).read_text(encoding="cp949")
        except Exception as exc:
            show_warning_message(self, "읽기 실패", f"파일을 읽을 수 없습니다.\n{exc}")
            return
        unfolded: List[str] = []
        for line in text.splitlines():
            if line.startswith(" ") or line.startswith("\t"):
                if unfolded:
                    unfolded[-1] += line.strip()
            else:
                unfolded.append(line.strip())
        events: List[Tuple[str, str]] = []
        current_summary = ""
        current_date = ""
        inside = False
        for line in unfolded:
            if line == "BEGIN:VEVENT":
                inside = True
                current_summary = ""
                current_date = ""
                continue
            if line == "END:VEVENT":
                if inside and current_date:
                    events.append((current_date, current_summary))
                inside = False
                continue
            if not inside:
                continue
            if line.startswith("SUMMARY"):
                parts = line.split(":", 1)
                if len(parts) == 2:
                    current_summary = parts[1].strip()
            elif line.startswith("DTSTART"):
                parts = line.split(":", 1)
                if len(parts) == 2:
                    value = parts[1].strip()
                    digits = value[:8]
                    try:
                        iso = datetime.strptime(digits, "%Y%m%d").date().isoformat()
                        current_date = iso
                    except ValueError:
                        continue
        if not events:
            show_info_message(self, "안내", "추출된 휴일 정보가 없습니다.")
            return
        added = 0
        labeled = 0

        def updater(cfg: SchedulerConfig) -> None:
            nonlocal added, labeled
            for iso, summary in events:
                if iso not in cfg.holidays:
                    cfg.holidays.append(iso)
                    added += 1
                if summary:
                    if cfg.holiday_labels.get(iso) != summary:
                        cfg.holiday_labels[iso] = summary
                        labeled += 1

        self.cfg_mgr.update(updater)
        self.refresh()
        show_success_message(
            self,
            "가져오기 완료",
            f"총 {len(events)}건 중 새로 추가 {added}건, 라벨 갱신 {labeled}건",
        )

    def _context_remove(self, widget: QtWidgets.QListWidget, pos: QtCore.QPoint) -> None:
        item = widget.itemAt(pos)
        if not item:
            return
        menu = QtWidgets.QMenu(widget)
        act = menu.addAction("삭제")
        if menu.exec(widget.mapToGlobal(pos)) == act:
            key = item.data(Qt.UserRole)

            def updater(cfg: SchedulerConfig) -> None:
                if widget is self.single_list:
                    iso = key or item.text()
                    cfg.holidays = [d for d in cfg.holidays if d != iso]
                    cfg.holiday_labels.pop(iso, None)
                else:
                    if isinstance(key, tuple):
                        start, end = key
                    else:
                        parts = item.text().split(" ~ ")
                        start, end = parts if len(parts) == 2 else (None, None)
                    cfg.holiday_ranges = [
                        rng
                        for rng in cfg.holiday_ranges
                        if not (start and end and rng.get("start") == start and rng.get("end") == end)
                    ]

            self.cfg_mgr.update(updater)
            self.refresh()

    def _remove_selected(self, widget: QtWidgets.QListWidget) -> None:
        items = widget.selectedItems()
        if not items:
            show_info_message(self, "안내", "삭제할 항목을 선택하세요.")
            return
        for item in items:
            key = item.data(Qt.UserRole)

            def updater(cfg: SchedulerConfig, item=item, key=key) -> None:
                if widget is self.single_list:
                    iso = key or item.text()
                    cfg.holidays = [d for d in cfg.holidays if d != iso]
                    cfg.holiday_labels.pop(iso, None)
                else:
                    if isinstance(key, tuple):
                        start, end = key
                    else:
                        parts = item.text().split(" ~ ")
                        start, end = parts if len(parts) == 2 else (None, None)
                    cfg.holiday_ranges = [
                        rng
                        for rng in cfg.holiday_ranges
                        if not (start and end and rng.get("start") == start and rng.get("end") == end)
                    ]

            self.cfg_mgr.update(updater)
        self.refresh()


class SettingsPanel(FancyCard):
    test_completed = Signal(bool, str)
    log_generated = Signal(str)

    def __init__(self, cfg_mgr: ConfigManager, accent: str, parent=None) -> None:
        super().__init__("고급 설정", accent, parent)
        self.cfg_mgr = cfg_mgr
        self.set_subtitle("종료 정책과 네트워크, 테마 설정")
        outer = QtWidgets.QVBoxLayout()
        outer.setSpacing(16)
        form = QtWidgets.QFormLayout()
        form.setSpacing(12)
        self.target_edit = QtWidgets.QLineEdit(", ".join(cfg_mgr.config.targets))
        self.remote_toggle = StyledToggle()
        self.remote_toggle.setChecked(cfg_mgr.config.enable_remote_shutdown)
        self.remote_toggle.setToolTip("스케줄 종료 후 원격 PC 종료 명령을 전송합니다.")
        self.local_toggle = StyledToggle()
        self.local_toggle.setChecked(cfg_mgr.config.enable_local_shutdown)
        self.local_toggle.setToolTip("스케줄 종료 후 이 PC를 종료합니다.")
        self.startup_toggle = StyledToggle()
        self.startup_toggle.setChecked(cfg_mgr.config.start_with_os)
        self.startup_toggle.setToolTip("Windows 로그인 시 프로그램을 자동 실행합니다.")
        self.delay_spin = QtWidgets.QSpinBox()
        self.delay_spin.setRange(0, 300)
        self.delay_spin.setValue(cfg_mgr.config.shutdown_delay)
        self.accent_btn = QtWidgets.QPushButton("테마 색상 변경")
        self.accent_btn.setCursor(Qt.PointingHandCursor)
        path_row = QtWidgets.QHBoxLayout()
        self.config_path_label = QtWidgets.QLabel(str(self.cfg_mgr.storage_directory()))
        self.config_path_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.config_path_label.setWordWrap(True)
        path_change_btn = QtWidgets.QPushButton("변경")
        path_change_btn.setCursor(Qt.PointingHandCursor)
        path_row.addWidget(self.config_path_label, 1)
        path_row.addWidget(path_change_btn)
        logo_row = QtWidgets.QHBoxLayout()
        self.logo_path_label = QtWidgets.QLabel()
        self.logo_path_label.setWordWrap(True)
        self.logo_path_label.setProperty("role", "subtitle")
        self.logo_choose_btn = QtWidgets.QPushButton("로고 선택")
        self.logo_clear_btn = QtWidgets.QPushButton("기본값")
        for btn in (self.logo_choose_btn, self.logo_clear_btn):
            btn.setCursor(Qt.PointingHandCursor)
        logo_row.addWidget(self.logo_path_label, 1)
        logo_row.addWidget(self.logo_choose_btn)
        logo_row.addWidget(self.logo_clear_btn)
        password_row = QtWidgets.QHBoxLayout()
        self.user_password_btn = QtWidgets.QPushButton("일반 비밀번호 변경")
        self.admin_password_btn = QtWidgets.QPushButton("관리자 비밀번호 변경")
        for btn in (self.user_password_btn, self.admin_password_btn):
            btn.setCursor(Qt.PointingHandCursor)
        password_row.addWidget(self.user_password_btn)
        password_row.addWidget(self.admin_password_btn)
        password_row.addStretch(1)
        form.addRow("종료 대상 프로그램", self.target_edit)
        form.addRow("원격 종료", create_toggle_field("사용", self.remote_toggle))
        form.addRow("본체 종료", create_toggle_field("사용", self.local_toggle))
        form.addRow("종료 지연(초)", self.delay_spin)
        form.addRow("시작 프로그램 등록", create_toggle_field("사용", self.startup_toggle))
        form.addRow("테마", self.accent_btn)
        form.addRow("설정 저장 위치", path_row)
        form.addRow("상단 로고", logo_row)
        form.addRow("비밀번호 관리", password_row)
        outer.addLayout(form)
        hosts_group = QtWidgets.QGroupBox("원격 PC 목록")
        hosts_layout = QtWidgets.QVBoxLayout(hosts_group)
        hosts_layout.setContentsMargins(12, 12, 12, 12)
        hosts_layout.setSpacing(8)
        hosts_hint = QtWidgets.QLabel("IP 또는 호스트 이름과 접속 정보를 입력하면 원격 종료에 활용됩니다.")
        hosts_hint.setWordWrap(True)
        hosts_layout.addWidget(hosts_hint)
        self.host_table = QtWidgets.QTableWidget(0, 4)
        self.host_table.setHorizontalHeaderLabels(["IP/호스트", "계정", "비밀번호", "방식"])
        self.host_table.horizontalHeader().setStretchLastSection(True)
        self.host_table.verticalHeader().setVisible(False)
        self.host_table.setAlternatingRowColors(True)
        self.host_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.host_table.setEditTriggers(
            QtWidgets.QAbstractItemView.DoubleClicked
            | QtWidgets.QAbstractItemView.SelectedClicked
            | QtWidgets.QAbstractItemView.EditKeyPressed
        )
        hosts_layout.addWidget(self.host_table)
        host_btn_row = QtWidgets.QHBoxLayout()
        self.add_host_btn = QtWidgets.QPushButton("추가")
        self.remove_host_btn = QtWidgets.QPushButton("삭제")
        self.test_host_btn = QtWidgets.QPushButton("연결 시험")
        for btn in (self.add_host_btn, self.remove_host_btn):
            btn.setCursor(Qt.PointingHandCursor)
        self.test_host_btn.setCursor(Qt.PointingHandCursor)
        host_btn_row.addWidget(self.add_host_btn)
        host_btn_row.addWidget(self.remove_host_btn)
        host_btn_row.addWidget(self.test_host_btn)
        host_btn_row.addStretch(1)
        hosts_layout.addLayout(host_btn_row)
        outer.addWidget(hosts_group)
        container = QtWidgets.QWidget()
        container.setLayout(outer)
        self.body_layout.addWidget(container)
        self.target_edit.editingFinished.connect(self._persist_targets)
        self.remote_toggle.stateChanged.connect(lambda _: self._persist())
        self.local_toggle.stateChanged.connect(lambda _: self._persist())
        self.startup_toggle.stateChanged.connect(lambda _: self._persist())
        self.delay_spin.valueChanged.connect(lambda _: self._persist())
        self.accent_btn.clicked.connect(self._pick_color)
        self.add_host_btn.clicked.connect(self._add_host)
        self.remove_host_btn.clicked.connect(self._remove_host)
        self.test_host_btn.clicked.connect(self._test_host)
        self.host_table.itemChanged.connect(lambda _: self._persist_hosts())
        path_change_btn.clicked.connect(self._choose_config_dir)
        self.cfg_mgr.storage_dir_changed.connect(self._on_storage_dir_changed)
        self.user_password_btn.clicked.connect(self._change_user_password)
        self.admin_password_btn.clicked.connect(self._change_admin_password)
        self.logo_choose_btn.clicked.connect(self._choose_logo_image)
        self.logo_clear_btn.clicked.connect(self._clear_logo_image)
        self.test_completed.connect(self._on_test_result)
        self.log_generated.connect(self._append_test_log)
        self._targets_timer = QtCore.QTimer(self)
        self._targets_timer.setSingleShot(True)
        self._targets_timer.setInterval(400)
        self._targets_timer.timeout.connect(self._persist_targets)
        self.target_edit.textChanged.connect(lambda _: self._targets_timer.start())
        self._loading_hosts = False
        self._update_logo_summary(self.cfg_mgr.config.header_logo_path)
        self._load_hosts()
        self._log_dialog: Optional[TerminalLogDialog] = None

    def _persist_targets(self) -> None:
        raw = [p.strip() for p in self.target_edit.text().split(",") if p.strip()]
        if not raw:
            raw = DEFAULT_TARGETS.copy()
        def updater(cfg: SchedulerConfig) -> None:
            cfg.targets = raw
        self.cfg_mgr.update(updater)

    def _persist(self) -> None:
        start = self.startup_toggle.isChecked()
        def updater(cfg: SchedulerConfig) -> None:
            cfg.enable_remote_shutdown = self.remote_toggle.isChecked()
            cfg.enable_local_shutdown = self.local_toggle.isChecked()
            cfg.start_with_os = start
            cfg.shutdown_delay = self.delay_spin.value()
        self.cfg_mgr.update(updater)
        set_startup(start)

    def _pick_color(self) -> None:
        color = QtWidgets.QColorDialog.getColor(QtGui.QColor(self.cfg_mgr.config.theme_accent), self)
        if color.isValid():
            hex_color = color.name()
            self.cfg_mgr.update(lambda cfg: setattr(cfg, "theme_accent", hex_color))

    def sync_from_config(self) -> None:
        cfg = self.cfg_mgr.config
        self.target_edit.blockSignals(True)
        self.target_edit.setText(", ".join(cfg.targets))
        self.target_edit.blockSignals(False)
        for toggle, value in (
            (self.remote_toggle, cfg.enable_remote_shutdown),
            (self.local_toggle, cfg.enable_local_shutdown),
            (self.startup_toggle, cfg.start_with_os),
        ):
            toggle.blockSignals(True)
            toggle.setChecked(value)
            toggle.blockSignals(False)
        self.delay_spin.blockSignals(True)
        self.delay_spin.setValue(cfg.shutdown_delay)
        self.delay_spin.blockSignals(False)
        self._load_hosts()
        self._update_config_path_label()
        self._update_logo_summary(cfg.header_logo_path)

    def _load_hosts(self) -> None:
        self._loading_hosts = True
        self.host_table.setRowCount(0)
        for host in self.cfg_mgr.config.remote_hosts:
            row = self.host_table.rowCount()
            self.host_table.insertRow(row)
            for col, key in enumerate(["host", "username", "password", "method"]):
                value = host.get(key, "")
                if key == "method" and not value:
                    value = "ssh"
                item = QtWidgets.QTableWidgetItem(value)
                self.host_table.setItem(row, col, item)
        self._loading_hosts = False

    def _table_text(self, row: int, column: int) -> str:
        item = self.host_table.item(row, column)
        return item.text().strip() if item else ""

    def _persist_hosts(self) -> None:
        if getattr(self, "_loading_hosts", False):
            return
        hosts: List[Dict[str, str]] = []
        for row in range(self.host_table.rowCount()):
            host_entry = {
                "host": self._table_text(row, 0),
                "username": self._table_text(row, 1) or "",
                "password": self._table_text(row, 2) or "",
                "method": self._table_text(row, 3) or "ssh",
            }
            if not host_entry["host"]:
                continue
            hosts.append(host_entry)
        self.cfg_mgr.update(lambda cfg: setattr(cfg, "remote_hosts", hosts))

    def _test_host(self) -> None:
        row = self.host_table.currentRow()
        if row < 0:
            show_info_message(self, "안내", "시험할 원격 PC 행을 선택하세요.")
            return
        entry = {
            "host": self._table_text(row, 0),
            "username": self._table_text(row, 1) or "",
            "password": self._table_text(row, 2) or "",
            "method": (self._table_text(row, 3) or "ssh").lower(),
        }
        host = entry["host"].strip()
        if not host:
            show_warning_message(self, "오류", "대상 호스트를 먼저 입력하세요.")
            return
        self.test_host_btn.setEnabled(False)
        self.test_host_btn.setText("시험 중…")
        dialog = self._ensure_log_dialog()
        dialog.start_session(host)
        self.log_generated.emit("============================================================")
        self.log_generated.emit(f"TARGET   : {host} ({entry['method'].upper()})")
        user = entry.get("username") or "-"
        self.log_generated.emit(f"USERNAME : {user}")

        def worker(payload: Dict[str, str]) -> None:
            def emit_log(line: str) -> None:
                timestamp = datetime.now().strftime("%H:%M:%S")
                self.log_generated.emit(f"[{timestamp}] {line}")

            try:
                success, message = self._perform_connection_test(payload, emit_log)
            except Exception as exc:  # pragma: no cover - 네트워크 예외 보호
                emit_log(f"예상치 못한 오류 발생: {exc}")
                success, message = False, f"예상치 못한 오류: {exc}"
            self.test_completed.emit(success, message)

        threading.Thread(target=worker, args=(entry,), daemon=True).start()

    def _ensure_log_dialog(self) -> TerminalLogDialog:
        if self._log_dialog is None:
            self._log_dialog = TerminalLogDialog(self)
        return self._log_dialog

    def _append_test_log(self, line: str) -> None:
        dialog = self._ensure_log_dialog()
        dialog.append_line(line)

    @staticmethod
    def _split_host_port(raw: str, default_port: int) -> Tuple[str, int]:
        target = raw.strip()
        if not target:
            return "", default_port
        if "://" in target:
            target = target.split("://", 1)[1]
        if target.startswith("[") and "]" in target:
            host_part, _, remainder = target.partition("]")
            host = host_part.strip("[]")
            if remainder.startswith(":"):
                try:
                    return host, int(remainder[1:])
                except ValueError:
                    return host, default_port
            return host, default_port
        if target.count(":") == 1:
            name, port_str = target.split(":", 1)
            if port_str.isdigit():
                return name, int(port_str)
        return target, default_port

    @staticmethod
    def _ping_host(host: str, timeout: int = 2) -> Tuple[Optional[bool], str]:
        if not shutil.which("ping"):
            return None, "ping 명령을 찾을 수 없어 포트 연결만 확인했습니다."
        system = platform.system().lower()
        if "win" in system:
            command = ["ping", "-n", "1", "-w", str(max(1, timeout) * 1000), host]
        else:
            command = ["ping", "-c", "1", "-W", str(max(1, timeout)), host]
        try:
            result = subprocess.run(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                encoding="utf-8",
                errors="ignore",
                timeout=max(3, timeout + 2),
            )
        except Exception as exc:
            return None, f"ping 명령을 실행할 수 없습니다: {exc}"
        output = (result.stdout or "").strip()
        return (result.returncode == 0), output

    def _perform_connection_test(
        self, entry: Dict[str, str], log: Optional[Callable[[str], None]] = None
    ) -> Tuple[bool, str]:
        def emit(message: str) -> None:
            if log is not None:
                log(message)

        method = entry.get("method", "ssh").lower()
        host_text = entry.get("host", "").strip()
        if method == "ssh":
            default_port = 22
        elif method in {"winrm", "winrm-http"}:
            default_port = 5985
        elif method in {"winrm-https"}:
            default_port = 5986
        else:
            default_port = 22
        host_name, port = self._split_host_port(host_text, default_port)
        if not host_name:
            return False, "호스트 정보를 해석할 수 없습니다."
        emit(f"호스트 분석 완료 → {host_name}:{port}")
        emit("ping 검사 시작")
        ping_result, ping_detail = self._ping_host(host_name)
        if ping_result is False:
            detail = ping_detail or "네트워크 응답이 없습니다."
            emit("ping 실패")
            if ping_detail:
                for line in ping_detail.splitlines():
                    emit(f"PING> {line}")
            return False, f"{host_name}에 ping 응답이 없습니다.\n{detail}"
        notes: List[str] = []
        if ping_result is True:
            notes.append(f"{host_name} ping 응답 확인")
            emit("ping 성공")
        elif ping_detail:
            notes.append(ping_detail)
            emit("ping 상세 로그 수신")
            for line in ping_detail.splitlines():
                emit(f"PING> {line}")
        else:
            emit("ping 명령을 사용할 수 없어 포트 검사로 계속 진행")
        if method == "ssh":
            emit("SSH 연결 시도 준비")
            if not paramiko:
                details = "Paramiko 모듈이 설치되어 있지 않아 SSH 연결을 시험할 수 없습니다."
                emit(details)
                if notes:
                    details = "\n".join(notes + [details])
                return False, details
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            try:
                emit("SSH 세션 생성 및 인증 요청")
                ssh.connect(
                    hostname=host_name,
                    port=port,
                    username=entry.get("username") or None,
                    password=entry.get("password") or None,
                    timeout=8,
                )
                ssh.exec_command("echo connection_test")
                ssh.close()
                notes.append(f"{host_name}:{port} SSH 연결에 성공했습니다.")
                emit("SSH 인증 성공 → 명령 에코 확인 완료")
                return True, "\n".join(notes)
            except Exception as exc:
                notes.append(f"{host_name}:{port} SSH 연결 실패: {exc}")
                emit(f"SSH 예외: {exc}")
                return False, "\n".join(notes)
        else:
            try:
                emit(f"소켓 연결 시도 ({host_name}:{port})")
                with socket.create_connection((host_name, port), timeout=6):
                    notes.append(f"{host_name}:{port} 포트에 접속할 수 있습니다.")
                    emit("포트 연결 성공")
                    return True, "\n".join(notes)
            except Exception as exc:
                notes.append(f"{host_name}:{port} 포트 연결 실패: {exc}")
                emit(f"포트 연결 실패: {exc}")
                return False, "\n".join(notes)

    def _on_test_result(self, success: bool, message: str) -> None:
        self.test_host_btn.setEnabled(True)
        self.test_host_btn.setText("연결 시험")
        if success:
            show_success_message(self, "연결 성공", message)
            self.log_generated.emit("✔ 최종 결과: 연결 성공")
        else:
            show_warning_message(self, "연결 실패", message)
            self.log_generated.emit("✖ 최종 결과: 연결 실패")
        for line in message.splitlines():
            self.log_generated.emit(f"SUMMARY> {line}")

    def _add_host(self) -> None:
        row = self.host_table.rowCount()
        self.host_table.insertRow(row)
        defaults = ["", "admin", "", "ssh"]
        self._loading_hosts = True
        for col, value in enumerate(defaults):
            item = QtWidgets.QTableWidgetItem(value)
            self.host_table.setItem(row, col, item)
        self._loading_hosts = False
        self.host_table.setCurrentCell(row, 0)
        self.host_table.editItem(self.host_table.item(row, 0))

    def _remove_host(self) -> None:
        rows = sorted({index.row() for index in self.host_table.selectedIndexes()}, reverse=True)
        if not rows:
            show_info_message(self, "안내", "삭제할 항목을 선택하세요.")
            return
        self._loading_hosts = True
        for row in rows:
            self.host_table.removeRow(row)
        self._loading_hosts = False
        self._persist_hosts()

    def _choose_config_dir(self) -> None:
        current = str(self.cfg_mgr.storage_directory())
        new_path = QtWidgets.QFileDialog.getExistingDirectory(self, "설정 저장 폴더 선택", current)
        if not new_path:
            return
        self.cfg_mgr.change_storage_dir(Path(new_path))

    def _on_storage_dir_changed(self, path: str) -> None:
        self._update_config_path_label(path)

    def _update_config_path_label(self, path: Optional[str] = None) -> None:
        target = path or str(self.cfg_mgr.storage_directory())
        self.config_path_label.setText(target)
        self.config_path_label.setToolTip(target)

    def _update_logo_summary(self, path: Optional[str]) -> None:
        if path:
            name = Path(path).name
            exists = Path(path).exists()
            status = "사용 중" if exists else "확인 필요"
            text = f"{status}: {name}"
            tooltip = path
        else:
            text = "기본 로고 사용 중 (자동 생성)"
            tooltip = "고급 설정에서 로고 이미지를 선택해 상단 바를 꾸밀 수 있습니다."
        self.logo_path_label.setText(text)
        self.logo_path_label.setToolTip(tooltip)

    def _change_user_password(self) -> None:
        dialog = PasswordChangeDialog("일반 사용자 비밀번호 변경", require_current=False, parent=self)
        if dialog.exec() != QtWidgets.QDialog.Accepted:
            return

        def updater(cfg: SchedulerConfig) -> None:
            cfg.user_password_hash = hash_password(dialog.new_password)

        self.cfg_mgr.update(updater)
        show_success_message(self, "완료", "일반 사용자 비밀번호가 변경되었습니다.")

    def _change_admin_password(self) -> None:
        dialog = PasswordChangeDialog("관리자 비밀번호 변경", require_current=True, parent=self)
        if dialog.exec() != QtWidgets.QDialog.Accepted:
            return
        if not verify_password(self.cfg_mgr.config.admin_password_hash, dialog.current_password):
            show_warning_message(self, "오류", "현재 관리자 비밀번호가 일치하지 않습니다.")
            return

        def updater(cfg: SchedulerConfig) -> None:
            cfg.admin_password_hash = hash_password(dialog.new_password)

        self.cfg_mgr.update(updater)
        show_success_message(self, "완료", "관리자 비밀번호가 변경되었습니다.")

    def _choose_logo_image(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            "로고 이미지 선택",
            str(Path.home()),
            "이미지 파일 (*.png *.jpg *.jpeg *.bmp *.gif)",
        )
        if not path:
            return

        def updater(cfg: SchedulerConfig) -> None:
            cfg.header_logo_path = path

        self.cfg_mgr.update(updater)
        self._update_logo_summary(path)
        show_success_message(self, "상단 로고", "선택한 이미지가 상단 바에 적용되었습니다.")

    def _clear_logo_image(self) -> None:
        if not self.cfg_mgr.config.header_logo_path:
            self._update_logo_summary(None)
            return

        def updater(cfg: SchedulerConfig) -> None:
            cfg.header_logo_path = None

        self.cfg_mgr.update(updater)
        self._update_logo_summary(None)
        show_info_message(self, "상단 로고", "기본 로고로 돌아갔습니다.")

class DateRangeDialog(QtWidgets.QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("기간 선택")
        self.setModal(True)
        self.setMinimumWidth(600)
        layout = QtWidgets.QVBoxLayout(self)
        self.start_calendar = QtWidgets.QCalendarWidget()
        self.end_calendar = QtWidgets.QCalendarWidget()
        start_label = QtWidgets.QLabel("시작일")
        start_label.setProperty("popup-role", "body")
        layout.addWidget(start_label)
        layout.addWidget(self.start_calendar)
        end_label = QtWidgets.QLabel("종료일")
        end_label.setProperty("popup-role", "body")
        layout.addWidget(end_label)
        layout.addWidget(self.end_calendar)
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        layout.addWidget(buttons)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        self.result_range = ("", "")
        _apply_popup_typography(self)

    def accept(self) -> None:
        start = self.start_calendar.selectedDate().toPython()
        end = self.end_calendar.selectedDate().toPython()
        if end < start:
            show_warning_message(self, "오류", "종료일은 시작일 이후여야 합니다.")
            return
        self.result_range = (start.isoformat(), end.isoformat())
        super().accept()

class PasswordChangeDialog(QtWidgets.QDialog):
    def __init__(self, title: str, require_current: bool, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(520)
        form = QtWidgets.QFormLayout(self)
        form.setSpacing(12)
        self._require_current = require_current
        self.current_edit: Optional[QtWidgets.QLineEdit] = None
        if require_current:
            current = QtWidgets.QLineEdit()
            current.setEchoMode(QtWidgets.QLineEdit.Password)
            current.setPlaceholderText("현재 비밀번호")
            form.addRow("현재 비밀번호", current)
            self.current_edit = current
        self.new_edit = QtWidgets.QLineEdit()
        self.new_edit.setEchoMode(QtWidgets.QLineEdit.Password)
        self.new_edit.setPlaceholderText("새 비밀번호")
        form.addRow("새 비밀번호", self.new_edit)
        self.confirm_edit = QtWidgets.QLineEdit()
        self.confirm_edit.setEchoMode(QtWidgets.QLineEdit.Password)
        self.confirm_edit.setPlaceholderText("새 비밀번호 확인")
        form.addRow("비밀번호 확인", self.confirm_edit)
        self.error_label = QtWidgets.QLabel()
        self.error_label.setProperty("popup-role", "error")
        form.addRow("", self.error_label)
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._try_accept)
        buttons.rejected.connect(self.reject)
        form.addRow(buttons)
        self.new_password: str = ""
        self.current_password: str = ""
        for row in range(form.rowCount()):
            item = form.itemAt(row, QtWidgets.QFormLayout.LabelRole)
            if item is None:
                continue
            widget = item.widget()
            if isinstance(widget, QtWidgets.QLabel):
                widget.setProperty("popup-role", "body")
        _apply_popup_typography(self)

    def _try_accept(self) -> None:
        new_value = self.new_edit.text().strip()
        confirm = self.confirm_edit.text().strip()
        if not new_value:
            self.error_label.setText("새 비밀번호를 입력하세요.")
            self.new_edit.setFocus()
            return
        if len(new_value) < 4:
            self.error_label.setText("비밀번호는 최소 4자 이상이어야 합니다.")
            self.new_edit.setFocus()
            return
        if new_value != confirm:
            self.error_label.setText("비밀번호 확인이 일치하지 않습니다.")
            self.confirm_edit.setFocus()
            return
        if self._require_current and self.current_edit is not None:
            current = self.current_edit.text().strip()
            if not current:
                self.error_label.setText("현재 비밀번호를 입력하세요.")
                self.current_edit.setFocus()
                return
            self.current_password = current
        self.new_password = new_value
        self.accept()

def _apply_popup_typography(dialog: QtWidgets.QDialog) -> None:
    text_color = QColor("#0F172A")
    hint_color = QColor("#26364A")
    error_color = QColor("#D32F2F")
    accent_color = "#2A5CAA"
    disabled_accent = "#B0C6F0"
    font_stack = ", ".join(f"'{family}'" for family in PREFERRED_UI_FONTS)
    body_font = _build_ui_font(18, QFont.Weight.DemiBold)
    hint_font = _build_ui_font(17, QFont.Weight.Medium)
    error_font = _build_ui_font(16, QFont.Weight.Bold)
    input_font = _build_ui_font(17, QFont.Weight.Medium)
    button_font = _build_ui_font(17, QFont.Weight.Bold)
    dialog.setStyleSheet(
        f"""
        QDialog {{
            background-color: #F8FAFF;
            font-family: {font_stack};
        }}
        QDialog QLineEdit {{
            font-size: 17px;
            padding: 10px 14px;
            border-radius: 10px;
            border: 2px solid #9FB4D9;
            background: #FFFFFF;
            color: {text_color.name()};
        }}
        QDialog QLineEdit:focus {{
            border-color: {accent_color};
        }}
        QDialog QPushButton {{
            background-color: {accent_color};
            border: none;
            color: #FFFFFF;
            font-size: 17px;
            font-weight: 700;
            padding: 8px 22px;
            border-radius: 12px;
        }}
        QDialog QPushButton:hover {{
            background-color: #386AD6;
        }}
        QDialog QPushButton:disabled {{
            background-color: {disabled_accent};
            color: rgba(255, 255, 255, 0.85);
        }}
        """
    )
    layout = dialog.layout()
    if isinstance(layout, (QtWidgets.QVBoxLayout, QtWidgets.QFormLayout)):
        layout.setContentsMargins(28, 28, 28, 28)
        layout.setSpacing(18)
        if isinstance(layout, QtWidgets.QFormLayout):
            layout.setHorizontalSpacing(24)
            layout.setLabelAlignment(Qt.AlignLeft | Qt.AlignVCenter)
    for label in dialog.findChildren(QtWidgets.QLabel):
        role = label.property("popup-role")
        palette = label.palette()
        if role == "hint":
            label.setFont(hint_font)
            palette.setColor(QPalette.WindowText, hint_color)
        elif role == "error":
            label.setFont(error_font)
            palette.setColor(QPalette.WindowText, error_color)
        else:
            label.setFont(body_font)
            palette.setColor(QPalette.WindowText, text_color)
        label.setPalette(palette)
    for edit in dialog.findChildren(QtWidgets.QLineEdit):
        edit.setFont(input_font)
        palette = edit.palette()
        palette.setColor(QPalette.Text, text_color)
        palette.setColor(QPalette.PlaceholderText, hint_color)
        edit.setPalette(palette)
    for button in dialog.findChildren(QtWidgets.QPushButton):
        button.setFont(button_font)
        button.setMinimumHeight(42)
        button.setCursor(Qt.PointingHandCursor)

def _format_message_html(text: str) -> str:
    blocks = [block.strip("\n") for block in text.split("\n\n")]
    paragraphs = []
    for block in blocks:
        if not block:
            continue
        lines = [html.escape(part) for part in block.splitlines()]
        paragraphs.append("<p style='margin:0 0 10px 0;'>" + "<br>".join(lines) + "</p>")
    if not paragraphs:
        return "<p style='margin:0;'>" + html.escape(text.strip()) + "</p>"
    return "".join(paragraphs)


def _apply_messagebox_typography(box: QtWidgets.QMessageBox, level: str) -> None:
    palette = box.palette()
    palette.setColor(QPalette.Window, QColor("#F4F7FF"))
    palette.setColor(QPalette.WindowText, QColor("#0F172A"))
    box.setPalette(palette)
    accent = "#2A5CAA"
    tone_map = {
        "info": "#0F172A",
        "success": "#0F5B3A",
        "warning": "#8C3C10",
        "error": "#9C1F1F",
    }
    text_color = tone_map.get(level, tone_map["info"])
    font_stack = ", ".join(f"'{family}'" for family in PREFERRED_UI_FONTS)
    box.setStyleSheet(
        f"""
        QMessageBox {{
            background-color: #F4F7FF;
            border-radius: 18px;
            font-family: {font_stack};
        }}
        QMessageBox QLabel {{
            color: {text_color};
            font-weight: 600;
        }}
        QMessageBox QPushButton {{
            background-color: {accent};
            border-radius: 12px;
            padding: 8px 22px;
            color: #FFFFFF;
            font-family: {font_stack};
            font-weight: 700;
            font-size: 15px;
        }}
        QMessageBox QPushButton:hover {{
            background-color: #3B6FD6;
        }}
        QMessageBox QPushButton:pressed {{
            background-color: #305DC1;
        }}
        """
    )
    label = box.findChild(QtWidgets.QLabel, "qt_msgbox_label")
    if label is not None:
        font = _build_ui_font(16, QFont.Weight.DemiBold)
        label.setFont(font)
        label.setWordWrap(True)
        label.setTextFormat(Qt.RichText)
        label.setTextInteractionFlags(Qt.TextSelectableByMouse)
    info = box.findChild(QtWidgets.QLabel, "qt_msgboxex_icon_label")
    if info is not None:
        info.setMaximumWidth(64)
    button_font = _build_ui_font(15, QFont.Weight.Bold)
    for button in box.findChildren(QtWidgets.QPushButton):
        button.setCursor(Qt.PointingHandCursor)
        button.setFont(button_font)


def _show_message(parent, title: str, text: str, level: str = "info") -> QtWidgets.QMessageBox.StandardButton:
    icon_map = {
        "info": QtWidgets.QMessageBox.Information,
        "success": QtWidgets.QMessageBox.Information,
        "warning": QtWidgets.QMessageBox.Warning,
        "error": QtWidgets.QMessageBox.Critical,
    }
    box = QtWidgets.QMessageBox(parent)
    box.setWindowTitle(title)
    box.setIcon(icon_map.get(level, QtWidgets.QMessageBox.Information))
    box.setTextFormat(Qt.RichText)
    box.setText(_format_message_html(text))
    box.setStandardButtons(QtWidgets.QMessageBox.Ok)
    _apply_messagebox_typography(box, level)
    return box.exec()


def show_info_message(parent, title: str, text: str) -> QtWidgets.QMessageBox.StandardButton:
    return _show_message(parent, title, text, "info")


def show_success_message(parent, title: str, text: str) -> QtWidgets.QMessageBox.StandardButton:
    return _show_message(parent, title, text, "success")


def show_warning_message(parent, title: str, text: str) -> QtWidgets.QMessageBox.StandardButton:
    return _show_message(parent, title, text, "warning")


def show_error_message(parent, title: str, text: str) -> QtWidgets.QMessageBox.StandardButton:
    return _show_message(parent, title, text, "error")


class PasswordPrompt(QtWidgets.QDialog):
    def __init__(self, title: str, prompt: str, validator: Callable[[str], bool], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(520)
        self.validator = validator
        layout = QtWidgets.QVBoxLayout(self)
        message = QtWidgets.QLabel(prompt)
        message.setProperty("popup-role", "body")
        message.setWordWrap(True)
        layout.addWidget(message)
        self.password_edit = QtWidgets.QLineEdit()
        self.password_edit.setEchoMode(QtWidgets.QLineEdit.Password)
        self.password_edit.returnPressed.connect(self._attempt_login)
        layout.addWidget(self.password_edit)
        hint = QtWidgets.QLabel("비밀번호를 정확히 입력한 후 Enter 키를 누를 수 있습니다.")
        hint.setProperty("popup-role", "hint")
        hint.setWordWrap(True)
        layout.addWidget(hint)
        self.error_label = QtWidgets.QLabel()
        self.error_label.setProperty("popup-role", "error")
        layout.addWidget(self.error_label)
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._attempt_login)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        _apply_popup_typography(self)
        self.password_edit.setFocus()

    def _attempt_login(self) -> None:
        value = self.password_edit.text().strip()
        if self.validator(value):
            self.accept()
        else:
            self.error_label.setText("비밀번호가 일치하지 않습니다.")
            self.password_edit.selectAll()
            self.password_edit.setFocus()


class HelpDialog(QtWidgets.QDialog):
    def __init__(self, title: str, lines: List[str], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(560)
        layout = QtWidgets.QVBoxLayout(self)
        label = QtWidgets.QLabel("\n".join(lines))
        label.setProperty("popup-role", "body")
        label.setWordWrap(True)
        label.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        layout.addWidget(label)
        close_btn = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        close_btn.rejected.connect(self.reject)
        close_btn.accepted.connect(self.accept)
        layout.addWidget(close_btn)
        _apply_popup_typography(self)

class TerminalLogDialog(QtWidgets.QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("연결 시험 로그")
        self.setModal(False)
        self.setMinimumSize(640, 420)
        layout = QtWidgets.QVBoxLayout(self)
        caption = QtWidgets.QLabel(
            "실시간 진단 로그입니다. 테스트가 완료될 때까지 창을 열어두세요."
        )
        caption.setWordWrap(True)
        caption.setProperty("popup-role", "body")
        layout.addWidget(caption)
        self.output = QtWidgets.QPlainTextEdit()
        self.output.setReadOnly(True)
        font = QtGui.QFont("Cascadia Code", 12)
        self.output.setFont(font)
        self.output.setStyleSheet(
            "background-color: #071425; color: #6CFFB8; border-radius: 12px; padding: 12px;"
        )
        layout.addWidget(self.output, 1)
        self.close_btn = QtWidgets.QPushButton("닫기")
        self.close_btn.clicked.connect(self.hide)
        layout.addWidget(self.close_btn, 0, Qt.AlignRight)
        _apply_popup_typography(self)

    def start_session(self, target: str) -> None:
        self.output.clear()
        self.append_line(f"▶ {datetime.now():%H:%M:%S} - {target} 연결 시험을 시작합니다")
        self.show()
        self.raise_()
        self.activateWindow()

    def append_line(self, text: str) -> None:
        self.output.appendPlainText(text)
        self.output.moveCursor(QtGui.QTextCursor.End)


class CreditsDialog(QtWidgets.QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("제작 정보")
        self.setModal(True)
        self.setMinimumWidth(520)
        layout = QtWidgets.QVBoxLayout(self)
        title = QtWidgets.QLabel(f"<b>{APP_NAME}</b> 제작 크레딧")
        title.setAlignment(Qt.AlignCenter)
        title.setProperty("popup-role", "body")
        layout.addWidget(title)
        grid = QtWidgets.QFormLayout()
        grid.setLabelAlignment(Qt.AlignRight)
        grid.setFormAlignment(Qt.AlignLeft | Qt.AlignTop)
        grid.setVerticalSpacing(8)
        grid.addRow("버전", QtWidgets.QLabel(APP_VERSION))
        grid.addRow("제작자", QtWidgets.QLabel(AUTHOR_NAME))
        grid.addRow("제작 날짜", QtWidgets.QLabel(BUILD_DATE))
        grid.addRow("저작권", QtWidgets.QLabel("© 2025 Zolt46 / PSW / Emanon108. All rights reserved."))
        grid.addRow("문의", QtWidgets.QLabel("다산정보관 참고자료실 데스크"))
        grid_widget = QtWidgets.QWidget()
        grid_widget.setLayout(grid)
        layout.addWidget(grid_widget)
        note = QtWidgets.QLabel(
            "• 고급 설정 → 상단 로고에서 원하는 이미지를 선택하면 상단 중앙 로고가 교체됩니다.\n"
            "• 로고를 더블 클릭하면 ??"
        )
        note.setWordWrap(True)
        note.setProperty("popup-role", "body")
        layout.addWidget(note)
        close_btn = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        close_btn.rejected.connect(self.reject)
        close_btn.accepted.connect(self.accept)
        layout.addWidget(close_btn)
        _apply_popup_typography(self)


class EasterEggDialog(QtWidgets.QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("AutoClose Secret Studio")
        self.setModal(True)
        self.setMinimumWidth(520)
        layout = QtWidgets.QVBoxLayout(self)
        label = QtWidgets.QLabel(
            "<pre style='color:#3BFFB6; font-family: "
            "Cascadia Code, Consolas, monospace; font-size: 13px;'>\n"
            "   █████╗ ██╗   ██╗████████╗ ██████╗  ██████╗██╗      ██████╗ ███████╗\n"
            "  ██╔══██╗██║   ██║╚══██╔══╝██╔═══██╗██╔════╝██║     ██╔═══██╗██╔════╝\n"
            "  ███████║██║   ██║   ██║   ██║   ██║██║     ██║     ██║   ██║█████╗  \n"
            "  ██╔══██║██║   ██║   ██║   ██║   ██║██║     ██║     ██║   ██║██╔══╝  \n"
            "  ██║  ██║╚██████╔╝   ██║   ╚██████╔╝╚██████╗███████╗╚██████╔╝██║     \n"
            "  ╚═╝  ╚═╝ ╚═════╝    ╚═╝    ╚═════╝  ╚═════╝╚══════╝ ╚═════╝ ╚═╝     \n"
            "</pre>"
        )
        label.setAlignment(Qt.AlignCenter)
        label.setTextFormat(Qt.RichText)
        layout.addWidget(label)
        message = QtWidgets.QLabel(
            "숨은 공방을 찾아내셨군요!\n"
            "마감이 귀찮은 당신을 위한,\n 자동화 작업은 언제나 \n AutoClose가 든든하게 함께합니다."
        )
        message.setAlignment(Qt.AlignCenter)
        message.setWordWrap(True)
        message.setProperty("popup-role", "body")
        layout.addWidget(message)
        close_btn = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Close)
        close_btn.rejected.connect(self.reject)
        close_btn.accepted.connect(self.accept)
        layout.addWidget(close_btn)
        _apply_popup_typography(self)


class DashboardCard(FancyCard):
    request_force_run = Signal()

    def __init__(self, accent: str, parent=None) -> None:
        super().__init__("다음 실행 내역", accent, parent)
        self.set_subtitle("예약된 스케줄과 재생 정보를 함께 확인합니다")
        self.timer_label = QtWidgets.QLabel("다음 일정 계산 중…")
        self.timer_label.setProperty("role", "title")
        self.timer_label.setAlignment(Qt.AlignCenter)
        self.body_layout.addWidget(self.timer_label)
        self.detail_label = QtWidgets.QLabel()
        self.detail_label.setAlignment(Qt.AlignCenter)
        self.detail_label.setProperty("role", "subtitle")
        self.body_layout.addWidget(self.detail_label)
        self.audio_label = QtWidgets.QLabel("음성: -")
        self.audio_label.setAlignment(Qt.AlignCenter)
        self.audio_label.setProperty("role", "subtitle")
        self.body_layout.addWidget(self.audio_label)
        self.config_label = QtWidgets.QLabel("원격 종료: - · 본체 종료: -")
        self.config_label.setAlignment(Qt.AlignCenter)
        self.config_label.setProperty("role", "subtitle")
        self.body_layout.addWidget(self.config_label)
        action = QtWidgets.QPushButton("지금 즉시 실행")
        action.setCursor(Qt.PointingHandCursor)
        action.clicked.connect(self.request_force_run.emit)
        self.body_layout.addWidget(action)

    def update_next_run(self, run: Optional[UpcomingRun]) -> None:
        if run is None:
            self.timer_label.setText("예정된 실행이 없습니다")
            self.detail_label.setText("활성화된 요일과 휴일 설정을 확인하세요")
            self.audio_label.setText("음성: -")
            self.audio_label.setToolTip("")
            self.config_label.setText("원격 종료: - · 본체 종료: -")
            return
        when = run.when
        now = datetime.now()
        diff = when - now
        hours, remainder = divmod(int(diff.total_seconds()), 3600)
        minutes, _ = divmod(remainder, 60)
        day_label = DAY_LABEL.get(run.day_key, "")
        self.timer_label.setText(f"{when:%Y-%m-%d %H:%M}")
        self.detail_label.setText(f"{day_label} · {hours}시간 {minutes}분 후 실행")
        if run.audio_path:
            name = Path(run.audio_path).name
            mode = "자동" if run.auto_assign else "수동"
            self.audio_label.setText(f"음성: {name} ({mode})")
            self.audio_label.setToolTip(run.audio_path)
        else:
            if run.auto_assign:
                self.audio_label.setText("음성: 자동 지정 대기 (플레이리스트 비어 있음)")
            else:
                self.audio_label.setText("음성: 지정된 파일 없음")
            self.audio_label.setToolTip("")
        remote_state = "허용" if run.remote_allowed else "미사용"
        local_state = "허용" if run.local_allowed else "미사용"
        self.config_label.setText(f"원격 종료: {remote_state} · 본체 종료: {local_state}")

class TodaySummaryCard(FancyCard):
    def __init__(self, accent: str, parent=None) -> None:
        super().__init__("오늘 일정", accent, parent)
        self.set_subtitle("금일 예약된 시간과 재생 정보를 확인합니다")
        layout = QtWidgets.QFormLayout()
        layout.setLabelAlignment(Qt.AlignLeft)
        layout.setFormAlignment(Qt.AlignLeft | Qt.AlignTop)
        layout.setHorizontalSpacing(12)
        layout.setVerticalSpacing(8)
        self.status_value = QtWidgets.QLabel("계산 중…")
        self.time_value = QtWidgets.QLabel("-")
        self.audio_value = QtWidgets.QLabel("-")
        self.remote_value = QtWidgets.QLabel("-")
        self.local_value = QtWidgets.QLabel("-")
        for label in (
            self.status_value,
            self.time_value,
            self.audio_value,
            self.remote_value,
            self.local_value,
        ):
            label.setProperty("role", "subtitle")
        layout.addRow("상태", self.status_value)
        layout.addRow("예약 시간", self.time_value)
        layout.addRow("재생 음성", self.audio_value)
        layout.addRow("원격 종료", self.remote_value)
        layout.addRow("본체 종료", self.local_value)
        wrapper = QtWidgets.QWidget()
        wrapper.setLayout(layout)
        self.body_layout.addWidget(wrapper)

    def update_from_config(self, cfg: SchedulerConfig, audio_preview: Optional[str]) -> None:
        today = datetime.now()
        day_key = DAY_KEYS[today.weekday()]
        day_cfg = cfg.days.get(day_key)
        if not day_cfg or not day_cfg.enabled:
            self.status_value.setText("오늘은 일정이 비활성화되어 있습니다")
            self.time_value.setText("-")
            self.audio_value.setText("-")
            self.remote_value.setText("-")
            self.local_value.setText("-")
            return
        self.status_value.setText("예정된 일정이 활성화되어 있습니다")
        self.time_value.setText(day_cfg.time)
        if audio_preview:
            name = Path(audio_preview).name
            self.audio_value.setText(name)
            self.audio_value.setToolTip(audio_preview)
        else:
            self.audio_value.setText("지정된 음성 없음")
            self.audio_value.setToolTip("")
        remote_state = "허용" if (cfg.enable_remote_shutdown and day_cfg.allow_remote) else "미사용"
        local_state = "허용" if (cfg.enable_local_shutdown and day_cfg.allow_local_shutdown) else "미사용"
        self.remote_value.setText(remote_state)
        self.local_value.setText(local_state)

    def update_next_run(self, run: Optional[UpcomingRun]) -> None:
        if run is None:
            return
        today = datetime.now().date()
        when = run.when
        if when.date() != today:
            return
        self.status_value.setText(f"오늘 {when:%H:%M}에 실행 예정")
        if run.audio_path:
            name = Path(run.audio_path).name
            self.audio_value.setText(name)
            self.audio_value.setToolTip(run.audio_path)
        elif run.auto_assign:
            self.audio_value.setText("자동 지정 대기")
            self.audio_value.setToolTip("")


class ShutdownLogCard(FancyCard):
    clear_requested = Signal()

    def __init__(self, accent: str, parent=None) -> None:
        super().__init__("원격/종료 로그", accent, parent)
        self.set_subtitle("최근 종료 요청 내역을 확인합니다")
        self.list_widget = QtWidgets.QListWidget()
        self.list_widget.setAlternatingRowColors(True)
        self.list_widget.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.body_layout.addWidget(self.list_widget)
        btn_row = QtWidgets.QHBoxLayout()
        btn_row.addStretch(1)
        self.clear_btn = QtWidgets.QPushButton("로그 비우기")
        self.clear_btn.setCursor(Qt.PointingHandCursor)
        self.clear_btn.clicked.connect(self.clear_requested.emit)
        btn_row.addWidget(self.clear_btn)
        self.body_layout.addLayout(btn_row)

    def update_logs(self, logs: List[Dict[str, str]]) -> None:
        self.list_widget.clear()
        for entry in logs:
            timestamp = entry.get("at", "")
            kind = entry.get("type", "")
            detail = entry.get("detail", "")
            text = f"[{timestamp}] {kind} - {detail}".strip()
            self.list_widget.addItem(text)
        if not logs:
            self.list_widget.addItem("기록된 종료 내역이 없습니다")

class StatusOverlay(QtWidgets.QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        self.label = QtWidgets.QLabel("작업 중")
        self.label.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("background-color: rgba(15, 41, 64, 220); border-radius: 18px;")
        self.label.setStyleSheet("color: white; font-size: 20px; font-weight: 600;")
        layout.addWidget(self.label)
        self.resize(280, 120)

    def show_message(self, text: str) -> None:
        self.label.setText(text)
        screen = QtGui.QGuiApplication.primaryScreen()
        if screen:
            center = screen.geometry().center()
            frame = self.frameGeometry()
            frame.moveCenter(center)
            self.move(frame.topLeft())
        self.show()

def _load_brand_icon() -> Optional[QIcon]:
    if DEFAULT_APP_ICON.exists():
        icon = QIcon(str(DEFAULT_APP_ICON))
        if not icon.isNull():
            return icon
    return None


def create_tray_icon(accent_color: str) -> QIcon:
    pixmap = QtGui.QPixmap(64, 64)
    pixmap.fill(Qt.transparent)
    painter = QtGui.QPainter(pixmap)
    painter.setRenderHint(QtGui.QPainter.Antialiasing)
    painter.setBrush(QtGui.QColor(accent_color))
    painter.setPen(QtGui.QColor("white"))
    painter.drawRoundedRect(4, 4, 56, 56, 18, 18)
    font = QtGui.QFont("Segoe UI", 18, QtGui.QFont.Bold)
    painter.setFont(font)
    painter.setPen(Qt.white)
    painter.drawText(pixmap.rect(), Qt.AlignCenter, "AC")
    painter.end()
    return QIcon(pixmap)


class MainWindow(QtWidgets.QMainWindow):
    show_login_requested = Signal()
    logout_requested = Signal()
    admin_login_requested = Signal()
    help_requested = Signal(str)

    def __init__(self, cfg_mgr: ConfigManager, brand_icon: Optional[QIcon] = None) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setFixedSize(1040, 720)
        self.cfg_mgr = cfg_mgr
        self._brand_icon = brand_icon
        self.scheduler = SchedulerEngine(cfg_mgr)
        self.audio_service = AudioService()
        self.audio_service.set_volume(self.cfg_mgr.config.audio_volume)
        self.overlay = StatusOverlay()
        self._pending_follow_up: Optional[Tuple[bool, bool]] = None
        self._playback_mode: str = "idle"
        self._active_day_key: Optional[str] = None
        self._cards: List[FancyCard] = []
        self._ignore_playback_finished = False
        self._mode: str = "user"
        self._locked: bool = True
        self._secret_clicks: int = 0
        self._last_secret_time: float = 0.0
        self.tray: Optional[QtWidgets.QSystemTrayIcon] = None
        self._build_palette()
        initial_icon = self._brand_icon or create_tray_icon(self.cfg_mgr.config.theme_accent)
        self.setWindowIcon(initial_icon)
        self._build_ui()
        self._connect_signals()
        self.scheduler.start()
        self._update_header_logo(self.cfg_mgr.config.header_logo_path)

    def _build_palette(self) -> None:
        accent = QtGui.QColor(self.cfg_mgr.config.theme_accent)
        background = QtGui.QColor("#FFFFFF")
        base = QtGui.QColor("#FFFFFF")
        text = QtGui.QColor("#1C2B3E")
        outline = QtGui.QColor("#D6E2F5")
        palette = self.palette()
        palette.setColor(QPalette.Window, background)
        palette.setColor(QPalette.WindowText, text)
        palette.setColor(QPalette.Base, base)
        palette.setColor(QPalette.Text, text)
        palette.setColor(QPalette.Button, accent)
        palette.setColor(QPalette.ButtonText, Qt.white)
        palette.setColor(QPalette.Highlight, accent)
        palette.setColor(QPalette.HighlightedText, Qt.white)
        self.setPalette(palette)
        accent_hex = accent.name()
        accent_hover = QtGui.QColor(accent).lighter(120).name()
        accent_border = QtGui.QColor(accent).darker(120).name()
        accent_disabled_bg = QtGui.QColor(accent).lighter(200).name()
        accent_disabled_border = QtGui.QColor(accent).lighter(170).name()
        list_selected = QtGui.QColor(accent).lighter(140).name()
        drawer_bg = QtGui.QColor(accent).lighter(200).name()
        drawer_checked = QtGui.QColor(accent).lighter(170).name()
        header_bg = QtGui.QColor(accent).lighter(210).name()
        text_hex = text.name()
        outline_hex = outline.name()
        self.setStyleSheet(
            f"""
            QMainWindow {{ background: {background.name()}; }}
            QPushButton {{
                background-color: {accent_hex};
                border: 1px solid {accent_border};
                color: white;
                padding: 8px 16px;
                border-radius: 10px;
                font-weight: 600;
            }}
            QPushButton:hover {{ background-color: {accent_hover}; }}
            QPushButton:disabled {{
                background-color: {accent_disabled_bg};
                border: 1px solid {accent_disabled_border};
                color: rgba(255, 255, 255, 0.75);
            }}
            QPushButton#HelpButton {{
                background-color: #FF7043;
                border: 1px solid #E64A19;
            }}
            QPushButton#HelpButton:hover {{
                background-color: #FF8A65;
                border: 1px solid #F4511E;
            }}
            QCheckBox, QLabel {{ color: {text_hex}; }}
            QWidget[toggleContainer="true"] {{
                background: transparent;
            }}
            QWidget[toggleContainer="true"] QLabel[toggleLabel="true"] {{
                color: {text_hex};
                font-weight: 600;
                font-size: 14px;
            }}
            QLabel#PageTitle {{
                font-weight: 700;
                font-size: 20px;
                color: {text_hex};
            }}
            QLabel#PageTitle:hover {{
                color: {accent_hex};
                text-decoration: underline;
            }}
            QLabel#HeaderLogo {{
                background: rgba(255, 255, 255, 0.78);
                border-radius: 18px;
                padding: 6px 18px;
                border: 1px solid {outline_hex};
            }}
            QListWidget {{
                background: #FFFFFF;
                color: {text_hex};
                border-radius: 12px;
                border: 1px solid {outline_hex};
                padding: 8px;
            }}
            QLineEdit, QTimeEdit, QSpinBox {{
                background: #FFFFFF;
                color: {text_hex};
                border-radius: 10px;
                padding: 6px 10px;
                border: 1px solid {outline_hex};
            }}
            QLineEdit:disabled, QTimeEdit:disabled, QSpinBox:disabled {{
                background: #F3F6FC;
                color: #7588A6;
            }}
            QListWidget::item:selected {{
                background: {list_selected};
                color: white;
            }}
            QMenu {{
                background: #FFFFFF;
                color: {text_hex};
                border: 1px solid {outline_hex};
            }}
            QScrollArea {{ background: transparent; }}
            QFrame#NavDrawer {{
                background: {drawer_bg};
                border-right: 1px solid {outline_hex};
            }}
            QFrame#NavDrawer QPushButton {{
                background: transparent;
                border: none;
                color: {text_hex};
                padding: 10px 14px;
                border-radius: 14px;
                text-align: left;
                font-weight: 600;
            }}
            QFrame#NavDrawer QPushButton:hover {{ background: {drawer_checked}; }}
            QFrame#NavDrawer QPushButton:checked {{
                background: {drawer_checked};
                color: {accent_hex};
            }}
            QToolButton#MenuButton {{
                background: transparent;
                border: none;
                font-size: 20px;
                color: {accent_hex};
                padding: 4px 10px;
                font-weight: 700;
            }}
            QToolButton#MenuButton:hover {{ color: {accent_hover}; }}
            QFrame#TopBar {{
                background: {header_bg};
                border-bottom: 1px solid {outline_hex};
            }}
        """
        )

    def _build_ui(self) -> None:
        self._nav_buttons: Dict[str, QtWidgets.QPushButton] = {}
        self._page_indices: Dict[str, int] = {}
        root = QtWidgets.QWidget()
        root_layout = QtWidgets.QHBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        self.drawer = QtWidgets.QFrame()
        self.drawer.setObjectName("NavDrawer")
        self.drawer_width = 240
        self.drawer.setMinimumWidth(0)
        self.drawer.setMaximumWidth(self.drawer_width)
        drawer_layout = QtWidgets.QVBoxLayout(self.drawer)
        drawer_layout.setContentsMargins(20, 28, 20, 28)
        drawer_layout.setSpacing(12)
        drawer_title = QtWidgets.QLabel("메뉴")
        drawer_title.setProperty("role", "title")
        drawer_layout.addWidget(drawer_title)
        drawer_layout.addSpacing(4)
        self.nav_group = QtWidgets.QButtonGroup(self)
        self.nav_group.setExclusive(True)

        content_frame = QtWidgets.QFrame()
        content_layout = QtWidgets.QVBoxLayout(content_frame)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)
        top_bar = QtWidgets.QFrame()
        top_bar.setObjectName("TopBar")
        top_layout = QtWidgets.QHBoxLayout(top_bar)
        top_layout.setContentsMargins(24, 18, 24, 18)
        top_layout.setSpacing(12)
        self.menu_button = QtWidgets.QToolButton()
        self.menu_button.setObjectName("MenuButton")
        self.menu_button.setText("☰")
        self.menu_button.setCursor(Qt.PointingHandCursor)
        top_layout.addWidget(self.menu_button, 0)
        self.page_title = QtWidgets.QLabel("홈")
        self.page_title.setObjectName("PageTitle")
        self.page_title.setProperty("role", "title")
        self.page_title.setCursor(Qt.PointingHandCursor)
        self.page_title.setToolTip("상단 제목(현재 페이지 이름)을 5회 연속 클릭하면 관리자 모드로 전환할 수 있습니다.")
        top_layout.addWidget(self.page_title, 0)
        top_layout.addStretch(1)
        self.logo_container = QtWidgets.QWidget()
        self.logo_container.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        logo_layout = QtWidgets.QHBoxLayout(self.logo_container)
        logo_layout.setContentsMargins(0, 0, 0, 0)
        logo_layout.setSpacing(0)
        logo_layout.addStretch(1)
        self.logo_label = QtWidgets.QLabel()
        self.logo_label.setObjectName("HeaderLogo")
        self.logo_label.setAlignment(Qt.AlignCenter)
        self.logo_label.setFixedHeight(46)
        self.logo_label.setCursor(Qt.PointingHandCursor)
        logo_layout.addWidget(self.logo_label, 0, Qt.AlignCenter)
        logo_layout.addStretch(1)
        top_layout.addWidget(self.logo_container, 1)
        top_layout.addStretch(1)
        self.info_button = QtWidgets.QPushButton("제작 정보")
        self.info_button.setCursor(Qt.PointingHandCursor)
        self.info_button.setToolTip("프로그램 제작자와 버전 정보를 확인합니다.")
        top_layout.addWidget(self.info_button, 0)
        self.help_button = QtWidgets.QPushButton("도움말")
        self.help_button.setObjectName("HelpButton")
        self.help_button.setCursor(Qt.PointingHandCursor)
        self.help_button.setToolTip("일반 기능 사용법")
        top_layout.addWidget(self.help_button, 0)
        self.admin_exit_button = QtWidgets.QPushButton("관리자 모드 종료")
        self.admin_exit_button.setCursor(Qt.PointingHandCursor)
        self.admin_exit_button.setVisible(False)
        top_layout.addWidget(self.admin_exit_button, 0)
        self.lock_button = QtWidgets.QPushButton("잠금")
        self.lock_button.setCursor(Qt.PointingHandCursor)
        self.lock_button.setToolTip("창을 숨기고 다시 열 때 비밀번호를 요구합니다.")
        top_layout.addWidget(self.lock_button, 0)
        content_layout.addWidget(top_bar, 0)
        self.content_stack = QtWidgets.QStackedWidget()
        content_layout.addWidget(self.content_stack, 1)

        root_layout.addWidget(self.drawer)
        root_layout.addWidget(content_frame, 1)
        self.setCentralWidget(root)

        # 페이지 구성
        self.dashboard = DashboardCard(self.cfg_mgr.config.theme_accent)
        self.today_card = TodaySummaryCard(self.cfg_mgr.config.theme_accent)
        self.log_card = ShutdownLogCard(self.cfg_mgr.config.theme_accent)
        home_container = QtWidgets.QWidget()
        home_layout = QtWidgets.QVBoxLayout(home_container)
        home_layout.setContentsMargins(24, 20, 24, 24)
        home_layout.setSpacing(16)
        home_layout.addWidget(self.today_card)
        home_layout.addWidget(self.dashboard)
        home_layout.addWidget(self.log_card)
        home_layout.addStretch(1)
        home_page = self._wrap_scroll(home_container)

        day_container = QtWidgets.QWidget()
        day_layout = QtWidgets.QGridLayout(day_container)
        day_layout.setSpacing(16)
        day_layout.setContentsMargins(0, 0, 0, 0)
        day_layout.setColumnStretch(0, 1)
        day_layout.setColumnStretch(1, 1)
        self.day_cards = {}
        for idx, key in enumerate(DAY_KEYS):
            card = DayCard(key, self.cfg_mgr, self.cfg_mgr.config.theme_accent)
            row, col = divmod(idx, 2)
            day_layout.addWidget(card, row, col)
            self.day_cards[key] = card
        day_wrapper = QtWidgets.QWidget()
        day_wrapper_layout = QtWidgets.QVBoxLayout(day_wrapper)
        day_wrapper_layout.setContentsMargins(24, 20, 24, 24)
        day_wrapper_layout.setSpacing(16)
        day_wrapper_layout.addWidget(day_container)
        day_wrapper_layout.addStretch(1)
        day_page = self._wrap_scroll(day_wrapper)

        self.playlist_panel = PlaylistPanel(self.cfg_mgr, self.cfg_mgr.config.theme_accent)
        self.assignment_preview = AutoAssignmentPreviewCard(self.cfg_mgr, self.cfg_mgr.config.theme_accent)
        playlist_wrapper = QtWidgets.QWidget()
        playlist_layout = QtWidgets.QVBoxLayout(playlist_wrapper)
        playlist_layout.setContentsMargins(24, 20, 24, 24)
        playlist_layout.setSpacing(16)
        playlist_layout.addWidget(self.playlist_panel)
        playlist_layout.addWidget(self.assignment_preview)
        playlist_layout.addStretch(1)
        playlist_page = self._wrap_scroll(playlist_wrapper)

        self.holiday_panel = HolidayPanel(self.cfg_mgr, self.cfg_mgr.config.theme_accent)
        holiday_wrapper = QtWidgets.QWidget()
        holiday_layout = QtWidgets.QVBoxLayout(holiday_wrapper)
        holiday_layout.setContentsMargins(24, 20, 24, 24)
        holiday_layout.setSpacing(16)
        holiday_layout.addWidget(self.holiday_panel)
        holiday_layout.addStretch(1)
        holiday_page = self._wrap_scroll(holiday_wrapper)

        self.settings_panel = SettingsPanel(self.cfg_mgr, self.cfg_mgr.config.theme_accent)
        settings_wrapper = QtWidgets.QWidget()
        settings_layout = QtWidgets.QVBoxLayout(settings_wrapper)
        settings_layout.setContentsMargins(24, 20, 24, 24)
        settings_layout.setSpacing(16)
        settings_layout.addWidget(self.settings_panel)
        settings_layout.addStretch(1)
        settings_page = self._wrap_scroll(settings_wrapper)

        page_definitions = [
            ("홈", home_page),
            ("요일 일정", day_page),
            ("플레이리스트", playlist_page),
            ("휴일", holiday_page),
            ("고급 설정", settings_page),
        ]

        for name, widget in page_definitions:
            index = self.content_stack.addWidget(widget)
            self._page_indices[name] = index
            btn = QtWidgets.QPushButton(name)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.clicked.connect(lambda _=False, n=name: self._set_active_page(n))
            self.nav_group.addButton(btn)
            self._nav_buttons[name] = btn
            drawer_layout.addWidget(btn)

        drawer_layout.addStretch(1)
        self.drawer.setMaximumWidth(self.drawer_width)
        self.drawer.hide()
        self.menu_button.clicked.connect(self._toggle_drawer)
        self.log_card.clear_requested.connect(self._clear_logs)
        self.help_button.clicked.connect(self._on_help_clicked)
        self.lock_button.clicked.connect(self._request_lock)
        self.admin_exit_button.clicked.connect(lambda: self.set_mode("user"))
        self._cards.extend(
            [
                self.dashboard,
                self.today_card,
                self.log_card,
                *self.day_cards.values(),
                self.playlist_panel,
                self.assignment_preview,
                self.holiday_panel,
                self.settings_panel,
            ]
        )
        self._set_active_page("홈")
        self.today_card.update_from_config(self.cfg_mgr.config, self._preview_audio_for_today())
        self.log_card.update_logs(self.cfg_mgr.config.shutdown_logs)
        self._create_tray()
        self.page_title.installEventFilter(self)
        self.logo_label.installEventFilter(self)
        self.set_mode("user")

    def _wrap_scroll(self, content: QtWidgets.QWidget) -> QtWidgets.QScrollArea:
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        scroll.setWidget(content)
        return scroll

    def _toggle_drawer(self) -> None:
        if self.drawer.isVisible():
            self.drawer.hide()
        else:
            self.drawer.show()
            self.drawer.raise_()

    def set_locked(self, locked: bool) -> None:
        self._locked = locked
        if locked:
            self.set_mode("user")
            self.overlay.hide()
            self.hide()

    def is_locked(self) -> bool:
        return self._locked

    def set_mode(self, mode: str) -> None:
        if mode not in {"user", "admin"}:
            return
        self._mode = mode
        is_admin = mode == "admin"
        self.log_card.setVisible(is_admin)
        for key in ("휴일", "고급 설정"):
            btn = self._nav_buttons.get(key)
            if btn:
                btn.setVisible(is_admin)
        if not is_admin and self.page_title.text() in {"휴일", "고급 설정"}:
            self._set_active_page("홈")
        self.admin_exit_button.setVisible(is_admin)
        self.help_button.setToolTip("일반 기능 사용법" if not is_admin else "고급 기능 사용법")

    def _request_lock(self) -> None:
        self.logout_requested.emit()

    def _on_help_clicked(self) -> None:
        self.help_requested.emit(self._mode)

    def _set_active_page(self, name: str) -> None:
        index = self._page_indices.get(name)
        if index is None:
            return
        current = self.content_stack.currentIndex()
        if current != index:
            self.content_stack.setCurrentIndex(index)
        self.page_title.setText(name)
        button = self._nav_buttons.get(name)
        if button and not button.isChecked():
            button.setChecked(True)
        if self.drawer.isVisible():
            self.drawer.hide()

    def _preview_audio_for_today(self) -> Optional[str]:
        cfg = self.cfg_mgr.config
        today_key = DAY_KEYS[datetime.now().weekday()]
        day_cfg = cfg.days.get(today_key)
        if not day_cfg or not day_cfg.enabled:
            return None
        return predict_playlist_for_day(cfg, today_key)

    def _on_next_run_changed(self, run: Optional[UpcomingRun]) -> None:
        self.dashboard.update_next_run(run)
        self.today_card.update_next_run(run)

    def _update_today_summary(self) -> None:
        self.today_card.update_from_config(self.cfg_mgr.config, self._preview_audio_for_today())

    def _append_shutdown_log(self, kind: str, detail: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

        def updater(cfg: SchedulerConfig) -> None:
            entry = {"at": timestamp, "type": kind, "detail": detail}
            cfg.shutdown_logs.insert(0, entry)
            del cfg.shutdown_logs[30:]

        self.cfg_mgr.update(updater)

    def _clear_logs(self) -> None:
        self.cfg_mgr.update(lambda cfg: setattr(cfg, "shutdown_logs", []))

    def _on_day_card_changed(self, _: str) -> None:
        self.scheduler._compute_next_run()
        self._update_today_summary()

    def _create_tray(self) -> None:
        self.tray = QtWidgets.QSystemTrayIcon(create_tray_icon(self.cfg_mgr.config.theme_accent), self)
        menu = QtWidgets.QMenu()
        show_action = menu.addAction("열기")
        run_action = menu.addAction("지금 실행")
        menu.addSeparator()
        exit_action = menu.addAction("완전히 종료")
        show_action.triggered.connect(self._handle_tray_show)
        run_action.triggered.connect(self._force_execute)
        exit_action.triggered.connect(self._exit_all)
        self.tray.setContextMenu(menu)
        self.tray.setToolTip(APP_NAME)
        self.tray.activated.connect(self._on_tray_activated)
        self.tray.show()

    def _connect_signals(self) -> None:
        self.dashboard.request_force_run.connect(self._force_execute)
        self.scheduler.schedule_triggered.connect(self._on_schedule_triggered)
        self.scheduler.next_run_changed.connect(self._on_next_run_changed)
        self.audio_service.playback_started.connect(self._on_playback_started)
        self.audio_service.playback_finished.connect(self._on_playback_finished)
        preview_signal = getattr(self.playlist_panel, "preview_requested", None)
        if preview_signal is not None:
            preview_signal.connect(self._on_preview_requested)
        else:
            self.playlist_panel.add_preview_listener(self._on_preview_requested)
        stop_signal = getattr(self.playlist_panel, "stop_preview_requested", None)
        if stop_signal is not None:
            stop_signal.connect(self._on_stop_preview)
        else:
            self.playlist_panel.add_stop_preview_listener(self._on_stop_preview)
        self.cfg_mgr.config_changed.connect(self._on_config_changed)
        for card in self.day_cards.values():
            card.changed.connect(self._on_day_card_changed)
        self.info_button.clicked.connect(self._show_credits_dialog)

    def _apply_theme(self, accent: str) -> None:
        self._build_palette()
        window_icon = self._brand_icon or create_tray_icon(accent)
        self.setWindowIcon(window_icon)
        app = QtWidgets.QApplication.instance()
        if app is not None:
            app.setWindowIcon(window_icon)
        tray_icon = create_tray_icon(accent)
        if self.tray:
            self.tray.setIcon(tray_icon)
        for card in self._cards:
            card.set_accent(accent)

    def _show_credits_dialog(self) -> None:
        dialog = CreditsDialog(self)
        dialog.exec()


    def _on_config_changed(self, cfg: SchedulerConfig) -> None:
        self.audio_service.set_volume(cfg.audio_volume)
        self._apply_theme(cfg.theme_accent)
        self._update_header_logo(cfg.header_logo_path)
        self.playlist_panel.refresh()
        self.assignment_preview.refresh()
        self.holiday_panel.refresh()
        self.settings_panel.sync_from_config()
        for card in self.day_cards.values():
            card.sync_from_config()
        self._update_today_summary()
        self.log_card.update_logs(cfg.shutdown_logs)
        self.scheduler._compute_next_run()

    def _generate_header_logo(self) -> QtGui.QPixmap:
        width, height = 260, 44
        pixmap = QtGui.QPixmap(width, height)
        pixmap.fill(Qt.transparent)
        painter = QtGui.QPainter(pixmap)
        painter.setRenderHints(
            QtGui.QPainter.Antialiasing | QtGui.QPainter.TextAntialiasing
        )

        accent = QtGui.QColor(self.cfg_mgr.config.theme_accent)
        accent_outline = QtGui.QColor(accent).darker(130)
        background = QtGui.QColor(255, 255, 255, 235)
        tagline_color = QtGui.QColor(accent).lighter(160)

        # 로고 캡슐 배경
        outline_pen = QtGui.QPen(accent_outline)
        outline_pen.setWidthF(1.6)
        outline_pen.setJoinStyle(Qt.RoundJoin)
        painter.setPen(outline_pen)
        painter.setBrush(background)
        painter.drawRoundedRect(QtCore.QRectF(1.5, 1.5, width - 3, height - 3), 16, 16)

        # 좌측 전원 아이콘
        icon_rect = QtCore.QRectF(14, 6, 32, 32)
        painter.setPen(Qt.NoPen)
        painter.setBrush(accent)
        painter.drawEllipse(icon_rect)

        power_pen = QtGui.QPen(Qt.white)
        power_pen.setWidthF(2.8)
        power_pen.setCapStyle(Qt.RoundCap)
        painter.setPen(power_pen)
        arc_rect = icon_rect.adjusted(6, 4, -6, -4)
        painter.drawArc(arc_rect, 40 * 16, 280 * 16)
        painter.drawLine(
            QtCore.QPointF(icon_rect.center().x(), icon_rect.top() + 7),
            QtCore.QPointF(icon_rect.center().x(), icon_rect.center().y() + 6),
        )

        # 텍스트 구성
        text_left = icon_rect.right() + 12
        text_rect = QtCore.QRectF(text_left, 6, width - text_left - 18, 22)
        title_font = _build_ui_font(16, QFont.Weight.Bold)
        painter.setFont(title_font)
        painter.setPen(accent_outline)
        metrics = painter.fontMetrics()
        baseline = text_rect.top() + (text_rect.height() + metrics.ascent() - metrics.descent()) / 2
        painter.drawText(QtCore.QPointF(text_rect.left(), baseline), "AutoClose")
        painter.setPen(QtGui.QPen(accent))
        painter.drawText(
            QtCore.QPointF(
                text_rect.left() + metrics.horizontalAdvance("AutoClose "), baseline
            ),
            "Studio",
        )

        tagline_font = _build_ui_font(10, QFont.Weight.Medium)
        tagline_font.setCapitalization(QFont.AllUppercase)
        painter.setFont(tagline_font)
        painter.setPen(QtGui.QPen(tagline_color))
        tagline_rect = QtCore.QRectF(text_left, height - 15, width - text_left - 18, 12)
        painter.drawText(tagline_rect, Qt.AlignLeft | Qt.AlignBottom, "Schedule Automation")

        painter.end()
        return pixmap

    def _update_header_logo(self, path: Optional[str]) -> None:
        pixmap: Optional[QtGui.QPixmap] = None
        tooltip = "상단 바에 표시할 로고 이미지를 고급 설정에서 선택하세요."
        if path:
            candidate_path = Path(path)
            if candidate_path.exists():
                candidate = QtGui.QPixmap(str(candidate_path))
                if not candidate.isNull():
                    pixmap = candidate
                    tooltip = str(candidate_path)
        if pixmap is None:
            pixmap = self._generate_header_logo()
            tooltip = "기본 전원 로고가 적용되었습니다. 고급 설정에서 이미지를 교체할 수 있습니다."
        scaled = pixmap.scaledToHeight(44, Qt.SmoothTransformation)
        self.logo_label.setPixmap(scaled)
        self.logo_label.setToolTip(tooltip)

    def eventFilter(self, obj: QtCore.QObject, event: QtCore.QEvent) -> bool:
        if obj is self.page_title and event.type() == QtCore.QEvent.MouseButtonPress:
            if self._locked or self._mode == "admin":
                return super().eventFilter(obj, event)
            if isinstance(event, QtGui.QMouseEvent) and event.button() == Qt.LeftButton:
                now = time.time()
                if now - self._last_secret_time > 2:
                    self._secret_clicks = 0
                self._secret_clicks += 1
                self._last_secret_time = now
                if self._secret_clicks >= 5:
                    self._secret_clicks = 0
                    self.admin_login_requested.emit()
        if obj is self.logo_label and event.type() == QtCore.QEvent.MouseButtonDblClick:
            if isinstance(event, QtGui.QMouseEvent) and event.button() == Qt.LeftButton:
                EasterEggDialog(self).exec()
                return True
        return super().eventFilter(obj, event)

    def _handle_tray_show(self) -> None:
        if self._locked:
            self.show_login_requested.emit()
        else:
            self.showNormal()
            self.raise_()
            self.activateWindow()

    def _on_tray_activated(self, reason: QtWidgets.QSystemTrayIcon.ActivationReason) -> None:
        if reason in (QtWidgets.QSystemTrayIcon.Trigger, QtWidgets.QSystemTrayIcon.DoubleClick):
            self._handle_tray_show()

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:  # pragma: no cover - Qt callback
        event.ignore()
        if self._locked:
            self.hide()
            if self.tray:
                self.tray.showMessage(APP_NAME, "프로그램은 계속 백그라운드에서 실행됩니다.")
        else:
            self.logout_requested.emit()

    def _on_schedule_triggered(self, day_key: str, audio_path: str, allow_remote: bool, allow_local: bool) -> None:
        if self.audio_service.player.playbackState() != QMediaPlayer.StoppedState:
            if self._playback_mode == "schedule":
                self._pending_follow_up = None
            if self._playback_mode != "idle":
                self._ignore_playback_finished = True
            self.overlay.hide()
            self.audio_service.stop()
        self._active_day_key = day_key
        self._pending_follow_up = (allow_remote, allow_local)
        self._playback_mode = "schedule"
        day_label = DAY_LABEL[day_key]
        terminate_programs(self.cfg_mgr.config.targets)
        if audio_path:
            self.overlay.show_message(f"{day_label} 일정 실행 준비 - {Path(audio_path).name}")
        else:
            self.overlay.show_message(f"{day_label} 일정 실행 준비")
        if self.tray:
            self.tray.showMessage(APP_NAME, f"{day_label} 일정이 시작되었습니다.", QtWidgets.QSystemTrayIcon.Information, 3000)
        self.audio_service.play(audio_path)

    def _force_execute(self) -> None:
        now_key = DAY_KEYS[datetime.now().weekday()]
        cfg = self.cfg_mgr.config
        day_cfg = cfg.days[now_key]
        audio = self.scheduler._resolve_audio(cfg, day_cfg) or ""
        allow_remote = cfg.enable_remote_shutdown and day_cfg.allow_remote
        allow_local = cfg.enable_local_shutdown and day_cfg.allow_local_shutdown
        self._on_schedule_triggered(now_key, audio, allow_remote, allow_local)

    def _on_preview_requested(self, path: str) -> None:
        if self._playback_mode == "schedule":
            show_warning_message(self, "진행 중", "일정 실행 중에는 미리 듣기를 사용할 수 없습니다.")
            return
        if self.audio_service.player.playbackState() != QMediaPlayer.StoppedState:
            self._ignore_playback_finished = True
            self.audio_service.stop()
        self._pending_follow_up = None
        self._active_day_key = None
        self.overlay.hide()
        self._playback_mode = "preview"
        self.audio_service.play(path)
        if self.tray:
            self.tray.showMessage(APP_NAME, f"미리 듣기: {Path(path).name}", QtWidgets.QSystemTrayIcon.Information, 3000)

    def _on_stop_preview(self) -> None:
        if self._playback_mode == "preview":
            self.audio_service.stop()
        else:
            show_info_message(self, "안내", "현재 미리 듣기 중이 아닙니다.")

    def _on_playback_started(self, path: str) -> None:
        if self._playback_mode != "schedule":
            return
        day_label = DAY_LABEL.get(self._active_day_key or "", "일정")
        name = Path(path).name if path else "음성 없음"
        self.overlay.show_message(f"{day_label} - {name} 재생 중")

    def _on_playback_finished(self, _: str) -> None:
        if self._ignore_playback_finished:
            self._ignore_playback_finished = False
            return
        if self._playback_mode == "schedule":
            allow_remote, allow_local = self._pending_follow_up or (False, False)
            cfg_snapshot = self.cfg_mgr.config
            if self._active_day_key:
                day_cfg = cfg_snapshot.days.get(self._active_day_key)
            else:
                day_cfg = None
            if day_cfg is not None:
                allow_remote = allow_remote and cfg_snapshot.enable_remote_shutdown and day_cfg.allow_remote
                allow_local = allow_local and cfg_snapshot.enable_local_shutdown and day_cfg.allow_local_shutdown
            else:
                allow_remote = allow_remote and cfg_snapshot.enable_remote_shutdown
                allow_local = allow_local and cfg_snapshot.enable_local_shutdown
            self._pending_follow_up = None
            self.overlay.hide()
            if allow_remote:
                hosts = [host.get("host", "") for host in self.cfg_mgr.config.remote_hosts if host.get("host")]
                detail = ", ".join(hosts) if hosts else "등록된 대상 없음"
                self._append_shutdown_log("원격 종료", detail)
                threading.Thread(target=shutdown_remote, args=(self.cfg_mgr.config.remote_hosts,), daemon=True).start()
            if allow_local:
                self._append_shutdown_log(
                    "본체 종료",
                    f"{self.cfg_mgr.config.shutdown_delay}초 후 종료",
                )
                threading.Thread(target=shutdown_local, args=(self.cfg_mgr.config.shutdown_delay,), daemon=True).start()
            if self.tray:
                self.tray.showMessage(APP_NAME, "일정 실행이 완료되었습니다.", QtWidgets.QSystemTrayIcon.Information, 3000)
            self.scheduler._compute_next_run()
        elif self._playback_mode == "preview":
            if self.tray:
                self.tray.showMessage(APP_NAME, "미리 듣기가 종료되었습니다.", QtWidgets.QSystemTrayIcon.Information, 2000)
        self._playback_mode = "idle"
        self._active_day_key = None
        self._update_today_summary()

    def _exit_all(self) -> None:
        self.scheduler.stop()
        self.audio_service.stop()
        self.tray.hide()
        QtWidgets.QApplication.quit()


class App(QtWidgets.QApplication):
    def __init__(self, argv: List[str]) -> None:
        super().__init__(argv)
        QtCore.QCoreApplication.setApplicationName(APP_NAME)
        QtCore.QCoreApplication.setApplicationVersion(APP_VERSION)
        QtCore.QCoreApplication.setOrganizationName(ORGANIZATION_NAME)
        QtCore.QCoreApplication.setOrganizationDomain(ORGANIZATION_DOMAIN)
        self.setQuitOnLastWindowClosed(False)
        self.cfg_mgr = ConfigManager()
        brand_icon = _load_brand_icon()
        accent = self.cfg_mgr.config.theme_accent
        tray_icon = create_tray_icon(accent)
        self.setWindowIcon(brand_icon or tray_icon)
        self.window = MainWindow(self.cfg_mgr, brand_icon)
        self.window.set_locked(True)
        self.window.show_login_requested.connect(lambda: self._show_user_login(initial=False))
        self.window.logout_requested.connect(self._lock_from_user)
        self.window.admin_login_requested.connect(self._show_admin_login)
        self.window.help_requested.connect(self._show_help)
        self._show_user_login(initial=True)

    def _show_user_login(self, *, initial: bool) -> None:
        if not self.window.is_locked():
            return
        prompt = PasswordPrompt(
            "일반 사용자 로그인",
            "일반 사용자 비밀번호를 입력하세요.",
            lambda pw: verify_password(self.cfg_mgr.config.user_password_hash, pw),
            parent=self.window,
        )
        result = prompt.exec()
        if result == QtWidgets.QDialog.Accepted:
            self.window.set_locked(False)
            self.window.set_mode("user")
            self.window.showNormal()
            self.window.raise_()
            self.window.activateWindow()
        elif initial:
            QtWidgets.QApplication.quit()

    def _lock_from_user(self) -> None:
        if not self.window.is_locked():
            self.window.set_locked(True)
            if self.window.tray:
                self.window.tray.showMessage(
                    APP_NAME,
                    "화면이 잠겼습니다. 트레이에서 다시 열면 비밀번호를 입력할 수 있습니다.",
                    QtWidgets.QSystemTrayIcon.Information,
                    3000,
                )

    def _show_admin_login(self) -> None:
        if self.window.is_locked() or self.window._mode == "admin":
            return
        prompt = PasswordPrompt(
            "관리자 로그인",
            "관리자 비밀번호를 입력하세요.",
            lambda pw: verify_password(self.cfg_mgr.config.admin_password_hash, pw),
            parent=self.window,
        )
        if prompt.exec() == QtWidgets.QDialog.Accepted:
            self.window.set_mode("admin")
            show_success_message(
                self.window,
                "관리자 모드",
                "고급 기능이 활성화되었습니다. '관리자 모드 종료' 버튼으로 일반 모드로 돌아갈 수 있습니다.",
            )

    def _show_help(self, mode: str) -> None:
        if mode == "admin":
            lines = [
                "관리자 모드는 상단 제목(상단 바의 현재 페이지 이름)을 5회 연속 클릭하면 로그인 창이 나타납니다.",
                "로그인 후 '휴일'과 '고급 설정' 페이지가 열리며, 원격 종료·설정 파일 위치 등을 관리할 수 있습니다.",
                "휴일 목록에서는 날짜나 기간을 추가하고 선택한 항목을 삭제할 수 있으며, 우클릭으로도 삭제할 수 있습니다.",
                "고급 설정에서 일반/관리자 비밀번호를 변경하고 원격 PC 목록, 테마, 시작 프로그램 등록을 조정하세요.",
                "관리자 모드 종료 버튼을 누르면 일반 모드로 돌아갑니다. 잠금 버튼은 창을 숨기고 다시 로그인하도록 합니다.",
            ]
            title = "고급 기능 도움말"
        else:
            lines = [
                "프로그램을 실행하면 일반 사용자 비밀번호를 입력해 홈 화면에 접근합니다.",
                "홈에서는 오늘 일정과 다음 실행 예정 정보를 확인할 수 있고, 필요하면 '지금 즉시 실행'으로 강제 실행할 수 있습니다.",
                "'요일 일정' 페이지에서 각 요일의 사용 여부, 시간, 자동/수동 음성을 조정합니다.",
                "'플레이리스트'에서는 자동 지정에 사용할 음성 목록을 관리하고, 자동 배정 미리보기 표에서 적용될 음성을 확인합니다.",
                "상단의 잠금 버튼을 누르면 창이 숨겨지고, 다시 열려면 트레이에서 '열기'를 눌러 비밀번호를 입력해야 합니다.",
            ]
            title = "일반 기능 도움말"
        dialog = HelpDialog(title, lines, parent=self.window)
        dialog.exec()


if __name__ == "__main__":
    app = App(sys.argv)
    sys.exit(app.exec())