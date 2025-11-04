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

import json
import os
import sys
import threading
import time
from dataclasses import dataclass, asdict, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional

import psutil

try:
    import paramiko  # type: ignore
except ImportError:  # pragma: no cover - optional dependency
    paramiko = None

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QIcon, QPalette, QColor
from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer, QMediaContent

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
    theme_accent: str = "#6F7BF7"
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


class ConfigManager(QtCore.QObject):
    config_changed = Signal(SchedulerConfig)

    def __init__(self) -> None:
        super().__init__()
        self._lock = threading.Lock()
        self.config = self._load()

    def _load(self) -> SchedulerConfig:
        if CONFIG_FILE.exists():
            try:
                data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
                return SchedulerConfig.from_dict(data)
            except Exception as exc:  # pragma: no cover - fall back to default
                print("[설정 읽기 실패]", exc)
        return SchedulerConfig()

    def save(self) -> None:
        with self._lock:
            CONFIG_FILE.write_text(json.dumps(self.config.as_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
        self.config_changed.emit(self.config)

    def update(self, updater) -> None:
        with self._lock:
            updater(self.config)
            CONFIG_FILE.write_text(json.dumps(self.config.as_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
            config = self.config
        self.config_changed.emit(config)


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
        if not day_cfg.enabled:
            return False
        if cfg.holidays_enabled and self._is_holiday(cfg, current_date):
            return False
        if day_cfg.last_ran == current_date.isoformat():
            return False
        return True

    def _is_holiday(self, cfg: SchedulerConfig, target: date) -> bool:
        if target.isoformat() in cfg.holidays:
            return True
        for rng in cfg.holiday_ranges:
            try:
                start = datetime.strptime(rng["start"], "%Y-%m-%d").date()
                end = datetime.strptime(rng["end"], "%Y-%m-%d").date()
                if start <= target <= end:
                    return True
            except Exception:
                continue
        return False

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
        self.setStyleSheet(
            """
            #FancyCard {
                border-radius: 18px;
                background: rgba(255, 255, 255, 0.1);
                border: 1px solid rgba(255, 255, 255, 0.3);
            }
            #FancyCard QLabel[role="title"] {
                color: white;
                font-weight: 600;
                font-size: 18px;
            }
            #FancyCard QLabel[role="subtitle"] {
                color: rgba(255, 255, 255, 0.8);
                font-size: 13px;
            }
            """
        )
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        self.title = QtWidgets.QLabel(title)
        self.title.setProperty("role", "title")
        layout.addWidget(self.title)
        self.subtitle = QtWidgets.QLabel()
        self.subtitle.setProperty("role", "subtitle")
        layout.addWidget(self.subtitle)
        self.body_layout = QtWidgets.QVBoxLayout()
        layout.addLayout(self.body_layout)

    def set_subtitle(self, text: str) -> None:
        self.subtitle.setText(text)


class DayCard(FancyCard):
    changed = Signal(str)

    def __init__(self, day_key: str, cfg_mgr: ConfigManager, accent: str, parent=None) -> None:
        super().__init__(DAY_LABEL[day_key], accent, parent)
        self.day_key = day_key
        self.cfg_mgr = cfg_mgr
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
        file_edit = QtWidgets.QLineEdit()
        file_edit.setPlaceholderText("직접 파일 지정 (자동 음성 해제 시)")
        browse_btn = QtWidgets.QPushButton("찾기")
        remote_chk = QtWidgets.QCheckBox("원격 종료 허용")
        local_chk = QtWidgets.QCheckBox("본체 종료")
        for widget in (enable_chk, auto_chk, remote_chk, local_chk):
            widget.setCursor(Qt.PointingHandCursor)
        browse_btn.setCursor(Qt.PointingHandCursor)
        layout.addWidget(enable_chk, 0, 0)
        layout.addWidget(auto_chk, 0, 1)
        layout.addWidget(time_edit, 0, 2)
        layout.addWidget(remote_chk, 1, 0)
        layout.addWidget(local_chk, 1, 1)
        layout.addWidget(file_edit, 2, 0, 1, 2)
        layout.addWidget(browse_btn, 2, 2)
        container = QtWidgets.QWidget()
        container.setLayout(layout)
        self.body_layout.addWidget(container)
        self.enable_chk = enable_chk
        self.auto_chk = auto_chk
        self.time_edit = time_edit
        self.file_edit = file_edit
        self.remote_chk = remote_chk
        self.local_chk = local_chk
        browse_btn.clicked.connect(self._pick_file)
        enable_chk.stateChanged.connect(lambda _: self._persist())
        auto_chk.stateChanged.connect(lambda _: self._persist())
        remote_chk.stateChanged.connect(lambda _: self._persist())
        local_chk.stateChanged.connect(lambda _: self._persist())
        time_edit.timeChanged.connect(lambda _: self._persist())
        file_edit.editingFinished.connect(self._persist)

    def sync_from_config(self) -> None:
        cfg = self.cfg_mgr.config.days[self.day_key]
        self.enable_chk.setChecked(cfg.enabled)
        self.auto_chk.setChecked(cfg.auto_assign)
        hh, mm = map(int, cfg.time.split(":"))
        self.time_edit.setTime(QtCore.QTime(hh, mm))
        self.file_edit.setText(cfg.audio_path or "")
        self.remote_chk.setChecked(cfg.allow_remote)
        self.local_chk.setChecked(cfg.allow_local_shutdown)

    def _pick_file(self) -> None:
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "음성 파일 선택", str(Path.home()), "Audio Files (*.mp3 *.wav *.ogg)")
        if file_path:
            self.file_edit.setText(file_path)
            self.auto_chk.setChecked(False)
            self._persist()

    def _persist(self) -> None:
        def updater(cfg: SchedulerConfig) -> None:
            day_cfg = cfg.days[self.day_key]
            day_cfg.enabled = self.enable_chk.isChecked()
            day_cfg.auto_assign = self.auto_chk.isChecked()
            day_cfg.audio_path = self.file_edit.text().strip() or None
            day_cfg.time = self.time_edit.time().toString("HH:mm")
            day_cfg.allow_remote = self.remote_chk.isChecked()
            day_cfg.allow_local_shutdown = self.local_chk.isChecked()
        self.cfg_mgr.update(updater)
        self.changed.emit(self.day_key)


class PlaylistPanel(FancyCard):
    def __init__(self, cfg_mgr: ConfigManager, accent: str, parent=None) -> None:
        super().__init__("플레이리스트", accent, parent)
        self.cfg_mgr = cfg_mgr
        self.set_subtitle("자동 음성 지정 시 순차적으로 사용")
        layout = QtWidgets.QVBoxLayout()
        layout.setSpacing(12)
        self.list_widget = QtWidgets.QListWidget()
        self.list_widget.setAlternatingRowColors(True)
        layout.addWidget(self.list_widget)
        btn_row = QtWidgets.QHBoxLayout()
        add_btn = QtWidgets.QPushButton("추가")
        remove_btn = QtWidgets.QPushButton("삭제")
        up_btn = QtWidgets.QPushButton("▲")
        down_btn = QtWidgets.QPushButton("▼")
        for btn in (add_btn, remove_btn, up_btn, down_btn):
            btn.setCursor(Qt.PointingHandCursor)
        btn_row.addWidget(add_btn)
        btn_row.addWidget(remove_btn)
        btn_row.addStretch(1)
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
        self.refresh()

    def refresh(self) -> None:
        self.list_widget.clear()
        for path in self.cfg_mgr.config.playlist:
            item = QtWidgets.QListWidgetItem(Path(path).name)
            item.setData(Qt.UserRole, path)
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
        layout = QtWidgets.QFormLayout()
        layout.setSpacing(12)
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
        layout.addRow("종료 대상 프로그램", self.target_edit)
        layout.addRow("원격 종료", self.remote_toggle)
        layout.addRow("본체 종료", self.local_toggle)
        layout.addRow("종료 지연(초)", self.delay_spin)
        layout.addRow("시작 프로그램 등록", self.startup_toggle)
        layout.addRow("테마", self.accent_btn)
        container = QtWidgets.QWidget()
        container.setLayout(layout)
        self.body_layout.addWidget(container)
        self.target_edit.editingFinished.connect(self._persist_targets)
        self.remote_toggle.stateChanged.connect(lambda _: self._persist())
        self.local_toggle.stateChanged.connect(lambda _: self._persist())
        self.startup_toggle.stateChanged.connect(lambda _: self._persist())
        self.delay_spin.valueChanged.connect(lambda _: self._persist())
        self.accent_btn.clicked.connect(self._pick_color)

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


class StatusOverlay(QtWidgets.QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        self.label = QtWidgets.QLabel("작업 중")
        self.label.setStyleSheet("color: white; font-size: 22px; font-weight: 600;")
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
        self.cfg_mgr = cfg_mgr
        self.scheduler = SchedulerEngine(cfg_mgr)
        self.audio_service = AudioService()
        self.overlay = StatusOverlay()
        self._build_palette()
        self._build_ui()
        self._connect_signals()
        self.scheduler.start()

    def _build_palette(self) -> None:
        accent = QtGui.QColor(self.cfg_mgr.config.theme_accent)
        palette = self.palette()
        palette.setColor(QPalette.Window, QtGui.QColor("#1E1E2E"))
        palette.setColor(QPalette.WindowText, Qt.white)
        palette.setColor(QPalette.Base, QtGui.QColor(30, 32, 46, 200))
        palette.setColor(QPalette.Text, Qt.white)
        palette.setColor(QPalette.Button, accent)
        palette.setColor(QPalette.ButtonText, Qt.white)
        self.setPalette(palette)
        self.setStyleSheet(
            """
            QMainWindow { background: #1E1E2E; }
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:1,y2:1, stop:0 rgba(255,255,255,0.15), stop:1 rgba(255,255,255,0.05));
                border: 1px solid rgba(255,255,255,0.2);
                color: white;
                padding: 8px 16px;
                border-radius: 12px;
            }
            QPushButton:hover { background: rgba(255,255,255,0.25); }
            QCheckBox, QLabel { color: white; }
            QListWidget { background: rgba(20, 22, 34, 0.8); color: white; border-radius: 12px; padding: 8px; }
            QLineEdit, QTimeEdit, QSpinBox {
                background: rgba(20, 22, 34, 0.8);
                color: white;
                border-radius: 10px;
                padding: 6px 10px;
                border: 1px solid rgba(255,255,255,0.2);
            }
            QTabWidget::pane {
                border: none;
                background: transparent;
            }
            QTabBar::tab {
                background: rgba(255,255,255,0.08);
                color: white;
                min-width: 120px;
                padding: 10px 14px;
                border-radius: 14px;
            }
            QTabBar::tab:selected { background: rgba(255,255,255,0.22); }
            QListWidget::item:selected { background: rgba(111, 123, 247, 0.6); }
            QMenu {
                background: rgba(20,22,34,0.95);
                color: white;
                border: 1px solid rgba(255,255,255,0.1);
            }
            """
        )

    def _build_ui(self) -> None:
        central = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(central)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(18)
        self.dashboard = DashboardCard(self.cfg_mgr.config.theme_accent)
        layout.addWidget(self.dashboard)
        tabs = QtWidgets.QTabWidget()
        day_widget = QtWidgets.QWidget()
        day_layout = QtWidgets.QGridLayout(day_widget)
        day_layout.setSpacing(16)
        self.day_cards = {}
        for idx, key in enumerate(DAY_KEYS):
            card = DayCard(key, self.cfg_mgr, self.cfg_mgr.config.theme_accent)
            row, col = divmod(idx, 2)
            day_layout.addWidget(card, row, col)
            self.day_cards[key] = card
        tabs.addTab(day_widget, "요일별")
        self.playlist_panel = PlaylistPanel(self.cfg_mgr, self.cfg_mgr.config.theme_accent)
        tabs.addTab(self.playlist_panel, "플레이리스트")
        self.holiday_panel = HolidayPanel(self.cfg_mgr, self.cfg_mgr.config.theme_accent)
        tabs.addTab(self.holiday_panel, "휴일")
        self.settings_panel = SettingsPanel(self.cfg_mgr, self.cfg_mgr.config.theme_accent)
        tabs.addTab(self.settings_panel, "설정")
        layout.addWidget(tabs)
        self.setCentralWidget(central)
        self._create_tray()

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
        self.scheduler.next_run_changed.connect(self.dashboard.update_next_run)
        self.audio_service.playback_started.connect(lambda _: self.overlay.show_message("음성 재생 중"))
        self.audio_service.playback_finished.connect(lambda _: self.overlay.hide())
        self.cfg_mgr.config_changed.connect(lambda cfg: self._apply_theme(cfg.theme_accent))
        for card in self.day_cards.values():
            card.changed.connect(lambda _: self.scheduler._compute_next_run())

    def _apply_theme(self, accent: str) -> None:
        self._build_palette()
        self.tray.setIcon(create_tray_icon(accent))

    def closeEvent(self, event: QtGui.QCloseEvent) -> None:  # pragma: no cover - Qt callback
        event.ignore()
        self.hide()
        self.tray.showMessage(APP_NAME, "프로그램은 계속 백그라운드에서 실행됩니다.")

    def _on_schedule_triggered(self, day_key: str, audio_path: str, allow_remote: bool, allow_local: bool) -> None:
        self.overlay.show_message(f"{DAY_LABEL[day_key]} 일정 실행 중")
        terminate_programs(self.cfg_mgr.config.targets)
        self.audio_service.play(audio_path)
        def _follow_up(_: str) -> None:
            if allow_remote:
                threading.Thread(target=shutdown_remote, args=(self.cfg_mgr.config.remote_hosts,), daemon=True).start()
            if allow_local:
                threading.Thread(target=shutdown_local, args=(self.cfg_mgr.config.shutdown_delay,), daemon=True).start()
        self.audio_service.playback_finished.connect(_follow_up, QtCore.Qt.UniqueConnection)

    def _force_execute(self) -> None:
        now_key = DAY_KEYS[datetime.now().weekday()]
        self._on_schedule_triggered(now_key, self.scheduler._resolve_audio(self.cfg_mgr.config, self.cfg_mgr.config.days[now_key]) or "", True, True)

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