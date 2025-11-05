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
import json
import os
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
from PySide6.QtGui import QIcon, QPalette, QColor
from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer

APP_NAME = "AutoClose Studio"
CONFIG_DIR = (
    Path(os.environ.get("APPDATA") or os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    / "auto_close_studio"
)
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = CONFIG_DIR / "settings.json"

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
    holidays: List[str] = field(default_factory=list)
    holiday_ranges: List[Dict[str, str]] = field(default_factory=list)
    start_with_os: bool = False
    theme_accent: str = "#2A5CAA"
    shutdown_logs: List[Dict[str, str]] = field(default_factory=list)
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
            "holidays",
            "holiday_ranges",
            "start_with_os",
            "theme_accent",
            "shutdown_logs",
        ):
            if key in data:
                setattr(base, key, data[key])
        days = {}
        for day_key, day_val in data.get("days", {}).items():
            days[day_key] = DaySchedule.from_dict(day_val)
        for missing in DAY_KEYS:
            days.setdefault(missing, DaySchedule(enabled=(missing not in {"sat", "sun"})))
        base.days = days
        return base

def is_holiday(cfg: SchedulerConfig, target: date) -> bool:
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


def predict_playlist_for_day(cfg: SchedulerConfig, target_day: str) -> Optional[str]:
    playlist_len = len(cfg.playlist)
    if target_day not in DAY_KEYS:
        return None
    rotation = cfg.playlist_rotation % max(1, playlist_len)
    now = datetime.now()
    index = rotation
    for offset in range(0, 28):
        current = now + timedelta(days=offset)
        day_key = DAY_KEYS[current.weekday()]
        day_cfg = cfg.days.get(day_key)
        if not day_cfg or not is_day_eligible(cfg, day_cfg, current.date()):
            continue
        if day_cfg.auto_assign:
            candidate = cfg.playlist[index % playlist_len] if playlist_len else None
            if day_key == target_day:
                return candidate
            if playlist_len:
                index = (index + 1) % playlist_len
        else:
            if day_key == target_day:
                return day_cfg.audio_path
    return None


class ConfigManager(QtCore.QObject):
    config_changed = Signal(SchedulerConfig)

    def __init__(self) -> None:
        super().__init__()
        self._lock = threading.Lock()
        self.config = self._load()
        atexit.register(self._flush_on_exit)

    def _load(self) -> SchedulerConfig:
        if CONFIG_FILE.exists():
            try:
                data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
                return SchedulerConfig.from_dict(data)
            except Exception as exc:  # pragma: no cover - fall back to default
                print("[설정 읽기 실패]", exc)
        return SchedulerConfig()

    def _write(self, config: SchedulerConfig) -> None:
        CONFIG_FILE.write_text(
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

    def _flush_on_exit(self) -> None:
        with self._lock:
            try:
                self._write(self.config)
            except Exception:
                pass

class SchedulerEngine(QtCore.QObject):
    schedule_triggered = Signal(str, str, bool, bool)  # day_key, audio_path, allow_remote, allow_local
    next_run_changed = Signal(Optional[datetime])

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
        now = datetime.now()
        best: Optional[datetime] = None
        cfg = self.cfg_mgr.config
        for offset in range(0, 14):
            d = now + timedelta(days=offset)
            day_key = DAY_KEYS[d.weekday()]
            day_cfg = cfg.days[day_key]
            if not self._is_day_eligible(cfg, day_cfg, d.date()):
                continue
            hh, mm = map(int, day_cfg.time.split(":"))
            candidate = d.replace(hour=hh, minute=mm, second=0, microsecond=0)
            if candidate < now:
                continue
            if best is None or candidate < best:
                best = candidate
        self.next_run_changed.emit(best)

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

    def play(self, path: str) -> None:
        if not path:
            self.playback_finished.emit("")
            return
        url = QtCore.QUrl.fromLocalFile(path)
        self.player.setSource(url)
        self.audio_output.setVolume(0.9)
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


class FancyCard(QtWidgets.QFrame):
    def __init__(self, title: str, accent: str, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("FancyCard")
        self._accent = accent
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
        border = accent.lighter(130).name()
        fill = accent.lighter(180).name()
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
        enable_chk = QtWidgets.QCheckBox("사용")
        auto_chk = QtWidgets.QCheckBox("자동 음성")
        time_edit = QtWidgets.QTimeEdit()
        time_edit.setDisplayFormat("HH:mm")
        manual_combo = QtWidgets.QComboBox()
        manual_combo.setSizeAdjustPolicy(QtWidgets.QComboBox.AdjustToContents)
        manual_combo.setInsertPolicy(QtWidgets.QComboBox.NoInsert)
        manual_combo.setEditable(False)
        manual_combo.addItem("선택 안 함", None)
        manual_combo.setCurrentIndex(0)
        remote_chk = QtWidgets.QCheckBox("원격 종료 허용")
        local_chk = QtWidgets.QCheckBox("본체 종료")
        auto_hint = QtWidgets.QLabel()
        auto_hint.setProperty("role", "subtitle")
        auto_hint.setWordWrap(True)
        auto_hint.hide()
        for widget in (enable_chk, auto_chk, remote_chk, local_chk):
            widget.setCursor(Qt.PointingHandCursor)
        layout.addWidget(enable_chk, 0, 0)
        layout.addWidget(auto_chk, 0, 1)
        layout.addWidget(time_edit, 0, 2)
        layout.addWidget(remote_chk, 1, 0)
        layout.addWidget(local_chk, 1, 1)
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
        container = QtWidgets.QWidget()
        container.setLayout(layout)
        self.body_layout.addWidget(container)
        add_btn.clicked.connect(self._add_files)
        remove_btn.clicked.connect(self._remove_selected)
        up_btn.clicked.connect(lambda: self._move_selected(-1))
        down_btn.clicked.connect(lambda: self._move_selected(1))
        preview_btn.clicked.connect(self._preview_selected)
        stop_btn.clicked.connect(self._emit_stop_preview)
        self.refresh()

    def refresh(self) -> None:
        self.list_widget.clear()
        for path in self.cfg_mgr.config.playlist:
            item = QtWidgets.QListWidgetItem(Path(path).name)
            item.setData(Qt.UserRole, path)
            item.setToolTip(path)
            self.list_widget.addItem(item)

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
            QtWidgets.QMessageBox.information(self, "안내", "미리 듣기할 파일을 선택하세요.")
            return
        path = item.data(Qt.UserRole)
        if not path or not Path(path).exists():
            QtWidgets.QMessageBox.warning(self, "재생 불가", "파일을 찾을 수 없습니다. 경로를 확인해주세요.")
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



class HolidayPanel(FancyCard):
    def __init__(self, cfg_mgr: ConfigManager, accent: str, parent=None) -> None:
        super().__init__("휴일 설정", accent, parent)
        self.cfg_mgr = cfg_mgr
        self.set_subtitle("지정된 날짜에는 스케줄이 실행되지 않습니다")
        layout = QtWidgets.QVBoxLayout()
        layout.setSpacing(10)
        toggle = QtWidgets.QCheckBox("휴일 기능 사용")
        toggle.setCursor(Qt.PointingHandCursor)
        layout.addWidget(toggle)
        single_row = QtWidgets.QHBoxLayout()
        add_single_btn = QtWidgets.QPushButton("날짜 추가")
        add_single_btn.setCursor(Qt.PointingHandCursor)
        single_row.addWidget(add_single_btn)
        single_row.addStretch(1)
        layout.addLayout(single_row)
        self.single_list = QtWidgets.QListWidget()
        layout.addWidget(self.single_list)
        range_row = QtWidgets.QHBoxLayout()
        add_range_btn = QtWidgets.QPushButton("기간 추가")
        add_range_btn.setCursor(Qt.PointingHandCursor)
        range_row.addWidget(add_range_btn)
        range_row.addStretch(1)
        layout.addLayout(range_row)
        self.range_list = QtWidgets.QListWidget()
        layout.addWidget(self.range_list)
        container = QtWidgets.QWidget()
        container.setLayout(layout)
        self.body_layout.addWidget(container)
        toggle.stateChanged.connect(lambda _: self._persist())
        add_single_btn.clicked.connect(self._add_single)
        add_range_btn.clicked.connect(self._add_range)
        self.single_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.single_list.customContextMenuRequested.connect(lambda pos: self._context_remove(self.single_list, pos))
        self.range_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.range_list.customContextMenuRequested.connect(lambda pos: self._context_remove(self.range_list, pos))
        self.toggle = toggle
        self.refresh()

    def refresh(self) -> None:
        cfg = self.cfg_mgr.config
        self.toggle.setChecked(cfg.holidays_enabled)
        self.single_list.clear()
        for item in cfg.holidays:
            self.single_list.addItem(item)
        self.range_list.clear()
        for rng in cfg.holiday_ranges:
            self.range_list.addItem(f"{rng['start']} ~ {rng['end']}")

    def _persist(self) -> None:
        enabled = self.toggle.isChecked()
        self.cfg_mgr.update(lambda cfg: setattr(cfg, "holidays_enabled", enabled))

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
            self.cfg_mgr.update(lambda cfg: cfg.holidays.append(selected) if selected not in cfg.holidays else None)
            self.refresh()

    def _add_range(self) -> None:
        dialog = DateRangeDialog(self)
        if dialog.exec() == QtWidgets.QDialog.Accepted:
            start, end = dialog.result_range
            def updater(cfg: SchedulerConfig) -> None:
                cfg.holiday_ranges.append({"start": start, "end": end})
            self.cfg_mgr.update(updater)
            self.refresh()

    def _context_remove(self, widget: QtWidgets.QListWidget, pos: QtCore.QPoint) -> None:
        item = widget.itemAt(pos)
        if not item:
            return
        menu = QtWidgets.QMenu(widget)
        act = menu.addAction("삭제")
        if menu.exec(widget.mapToGlobal(pos)) == act:
            text = item.text()
            def updater(cfg: SchedulerConfig) -> None:
                if widget is self.single_list:
                    cfg.holidays = [d for d in cfg.holidays if d != text]
                else:
                    cfg.holiday_ranges = [rng for rng in cfg.holiday_ranges if f"{rng['start']} ~ {rng['end']}" != text]
            self.cfg_mgr.update(updater)
            self.refresh()


class SettingsPanel(FancyCard):
    def __init__(self, cfg_mgr: ConfigManager, accent: str, parent=None) -> None:
        super().__init__("고급 설정", accent, parent)
        self.cfg_mgr = cfg_mgr
        self.set_subtitle("종료 정책과 네트워크, 테마 설정")
        outer = QtWidgets.QVBoxLayout()
        outer.setSpacing(16)
        form = QtWidgets.QFormLayout()
        form.setSpacing(12)
        self.target_edit = QtWidgets.QLineEdit(", ".join(cfg_mgr.config.targets))
        self.remote_toggle = QtWidgets.QCheckBox("원격 종료 활성화")
        self.remote_toggle.setChecked(cfg_mgr.config.enable_remote_shutdown)
        self.local_toggle = QtWidgets.QCheckBox("본체 종료 활성화")
        self.local_toggle.setChecked(cfg_mgr.config.enable_local_shutdown)
        self.startup_toggle = QtWidgets.QCheckBox("Windows 시작 시 자동 실행")
        self.startup_toggle.setChecked(cfg_mgr.config.start_with_os)
        self.delay_spin = QtWidgets.QSpinBox()
        self.delay_spin.setRange(0, 300)
        self.delay_spin.setValue(cfg_mgr.config.shutdown_delay)
        self.accent_btn = QtWidgets.QPushButton("테마 색상 변경")
        self.accent_btn.setCursor(Qt.PointingHandCursor)
        form.addRow("종료 대상 프로그램", self.target_edit)
        form.addRow("원격 종료", self.remote_toggle)
        form.addRow("본체 종료", self.local_toggle)
        form.addRow("종료 지연(초)", self.delay_spin)
        form.addRow("시작 프로그램 등록", self.startup_toggle)
        form.addRow("테마", self.accent_btn)
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
        for btn in (self.add_host_btn, self.remove_host_btn):
            btn.setCursor(Qt.PointingHandCursor)
        host_btn_row.addWidget(self.add_host_btn)
        host_btn_row.addWidget(self.remove_host_btn)
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
        self.host_table.itemChanged.connect(lambda _: self._persist_hosts())
        self._targets_timer = QtCore.QTimer(self)
        self._targets_timer.setSingleShot(True)
        self._targets_timer.setInterval(400)
        self._targets_timer.timeout.connect(self._persist_targets)
        self.target_edit.textChanged.connect(lambda _: self._targets_timer.start())
        self._loading_hosts = False
        self._load_hosts()

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
            QtWidgets.QMessageBox.information(self, "안내", "삭제할 항목을 선택하세요.")
            return
        self._loading_hosts = True
        for row in rows:
            self.host_table.removeRow(row)
        self._loading_hosts = False
        self._persist_hosts()


class DateRangeDialog(QtWidgets.QDialog):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("기간 선택")
        layout = QtWidgets.QVBoxLayout(self)
        self.start_calendar = QtWidgets.QCalendarWidget()
        self.end_calendar = QtWidgets.QCalendarWidget()
        layout.addWidget(QtWidgets.QLabel("시작일"))
        layout.addWidget(self.start_calendar)
        layout.addWidget(QtWidgets.QLabel("종료일"))
        layout.addWidget(self.end_calendar)
        buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        layout.addWidget(buttons)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        self.result_range = ("", "")

    def accept(self) -> None:
        start = self.start_calendar.selectedDate().toPython()
        end = self.end_calendar.selectedDate().toPython()
        if end < start:
            QtWidgets.QMessageBox.warning(self, "오류", "종료일은 시작일 이후여야 합니다")
            return
        self.result_range = (start.isoformat(), end.isoformat())
        super().accept()


class DashboardCard(FancyCard):
    request_force_run = Signal()

    def __init__(self, accent: str, parent=None) -> None:
        super().__init__("다음 실행", accent, parent)
        self.set_subtitle("예약된 스케줄 정보를 확인합니다")
        self.timer_label = QtWidgets.QLabel("다음 일정 계산 중…")
        self.timer_label.setProperty("role", "title")
        self.timer_label.setAlignment(Qt.AlignCenter)
        self.body_layout.addWidget(self.timer_label)
        self.detail_label = QtWidgets.QLabel()
        self.detail_label.setAlignment(Qt.AlignCenter)
        self.detail_label.setProperty("role", "subtitle")
        self.body_layout.addWidget(self.detail_label)
        action = QtWidgets.QPushButton("지금 즉시 실행")
        action.setCursor(Qt.PointingHandCursor)
        action.clicked.connect(self.request_force_run.emit)
        self.body_layout.addWidget(action)

    def update_next_run(self, when: Optional[datetime]) -> None:
        if when is None:
            self.timer_label.setText("예정된 실행이 없습니다")
            self.detail_label.setText("활성화된 요일을 확인하세요")
            return
        now = datetime.now()
        diff = when - now
        hours, remainder = divmod(int(diff.total_seconds()), 3600)
        minutes, _ = divmod(remainder, 60)
        self.timer_label.setText(f"{when:%Y-%m-%d %H:%M}")
        self.detail_label.setText(f"{hours}시간 {minutes}분 후 실행")

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

    def update_next_run(self, when: Optional[datetime]) -> None:
        if when is None:
            return
        today = datetime.now().date()
        if when.date() != today:
            return
        self.status_value.setText(f"오늘 {when:%H:%M}에 실행 예정")


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
    def __init__(self, cfg_mgr: ConfigManager) -> None:
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setFixedSize(1040, 720)
        self.cfg_mgr = cfg_mgr
        self.scheduler = SchedulerEngine(cfg_mgr)
        self.audio_service = AudioService()
        self.overlay = StatusOverlay()
        self._pending_follow_up: Optional[Tuple[bool, bool]] = None
        self._playback_mode: str = "idle"
        self._active_day_key: Optional[str] = None
        self._cards: List[FancyCard] = []
        self._ignore_playback_finished = False
        self._build_palette()
        self._build_ui()
        self._connect_signals()
        self.scheduler.start()

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
            QCheckBox, QLabel {{ color: {text_hex}; }}
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
        self.drawer.setFixedWidth(240)
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
        self.page_title.setProperty("role", "title")
        top_layout.addWidget(self.page_title, 0)
        top_layout.addStretch(1)
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
        home_layout.addWidget(self.dashboard)
        home_layout.addWidget(self.today_card)
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
        playlist_wrapper = QtWidgets.QWidget()
        playlist_layout = QtWidgets.QVBoxLayout(playlist_wrapper)
        playlist_layout.setContentsMargins(24, 20, 24, 24)
        playlist_layout.setSpacing(16)
        playlist_layout.addWidget(self.playlist_panel)
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
        self.drawer.hide()
        self.menu_button.clicked.connect(self._toggle_drawer)
        self.log_card.clear_requested.connect(self._clear_logs)
        self._cards.extend(
            [
                self.dashboard,
                self.today_card,
                self.log_card,
                *self.day_cards.values(),
                self.playlist_panel,
                self.holiday_panel,
                self.settings_panel,
            ]
        )
        self._set_active_page("홈")
        self.today_card.update_from_config(self.cfg_mgr.config, self._preview_audio_for_today())
        self.log_card.update_logs(self.cfg_mgr.config.shutdown_logs)
        self._create_tray()

    def _wrap_scroll(self, content: QtWidgets.QWidget) -> QtWidgets.QScrollArea:
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setFrameShape(QtWidgets.QFrame.NoFrame)
        scroll.setWidget(content)
        return scroll

    def _toggle_drawer(self) -> None:
        self.drawer.setVisible(not self.drawer.isVisible())

    def _set_active_page(self, name: str) -> None:
        index = self._page_indices.get(name)
        if index is None:
            return
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

    def _on_next_run_changed(self, when: Optional[datetime]) -> None:
        self.dashboard.update_next_run(when)
        self.today_card.update_next_run(when)

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
        show_action.triggered.connect(self.showNormal)
        run_action.triggered.connect(self._force_execute)
        exit_action.triggered.connect(self._exit_all)
        self.tray.setContextMenu(menu)
        self.tray.setToolTip(APP_NAME)
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

    def _apply_theme(self, accent: str) -> None:
        self._build_palette()
        self.tray.setIcon(create_tray_icon(accent))
        for card in self._cards:
            card.set_accent(accent)

    def _on_config_changed(self, cfg: SchedulerConfig) -> None:
        self._apply_theme(cfg.theme_accent)
        self.playlist_panel.refresh()
        self.holiday_panel.refresh()
        self.settings_panel.sync_from_config()
        for card in self.day_cards.values():
            card.sync_from_config()
        self._update_today_summary()
        self.log_card.update_logs(cfg.shutdown_logs)

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:  # pragma: no cover - Qt callback
        event.ignore()
        self.hide()
        self.tray.showMessage(APP_NAME, "프로그램은 계속 백그라운드에서 실행됩니다.")

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
        self._on_schedule_triggered(now_key, self.scheduler._resolve_audio(self.cfg_mgr.config, self.cfg_mgr.config.days[now_key]) or "", True, True)

    def _on_preview_requested(self, path: str) -> None:
        if self._playback_mode == "schedule":
            QtWidgets.QMessageBox.warning(self, "진행 중", "일정 실행 중에는 미리 듣기를 사용할 수 없습니다.")
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
            QtWidgets.QMessageBox.information(self, "안내", "현재 미리 듣기 중이 아닙니다.")

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
        self.setQuitOnLastWindowClosed(False)
        self.cfg_mgr = ConfigManager()
        self.window = MainWindow(self.cfg_mgr)
        self.window.show()


if __name__ == "__main__":
    app = App(sys.argv)
    sys.exit(app.exec())