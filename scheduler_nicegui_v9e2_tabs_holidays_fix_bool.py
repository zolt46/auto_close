
# -*- coding: utf-8 -*-
"""
Scheduler — NiceGUI v9d
- 플레이리스트 탭을 넓게(컨테이너 width 확대, 테이블 영역 키움)
- "순차 적용"이 이제 **활성(사용=ON) 요일만**을 기준으로 순환
  · 예) 월~금만 활성: 월=base, 화=base+1, ... 금=base+4, 주말 비활성 → 다음 주 월에서 이어서 base로 루프
- 나머지 UI/동작은 v9c2 기반
"""
import os, json, threading, time, queue
from datetime import datetime, timedelta
from nicegui import ui
import psutil, pygame, paramiko

APP_DIR = os.path.join(os.environ.get('LOCALAPPDATA', os.getcwd()), 'MyScheduler')
os.makedirs(APP_DIR, exist_ok=True)
CONFIG_FILE = os.path.join(APP_DIR, 'scheduler_config.json')

servers = [
    {"host": "220.68.79.81", "user": "user", "password": "0000"},
    {"host": "220.68.79.82", "user": "user", "password": "0000"},
    {"host": "220.68.79.83", "user": "user", "password": "0000"},
    {"host": "220.68.79.84", "user": "user", "password": "0000"},
]
SHUTDOWN_CMD = 'shutdown /s /t 0'

DAYS = ["mon","tue","wed","thu","fri","sat","sun"]
DAY_LABEL = {"mon":"월","tue":"화","wed":"수","thu":"목","fri":"금","sat":"토","sun":"일"}

DEFAULT_CONFIG = {
    "playlist": [],
    "weekly": {"base_index": 0},
    "holidays": {"enabled": True, "dates": [], "ranges": []},
    "schedule": {
        "mon": {"enabled": True,  "mode": "weekly", "time": {"hh": 9, "mm": 0}, "file": None},
        "tue": {"enabled": True,  "mode": "weekly", "time": {"hh": 9, "mm": 0}, "file": None},
        "wed": {"enabled": True,  "mode": "weekly", "time": {"hh": 9, "mm": 0}, "file": None},
        "thu": {"enabled": True,  "mode": "weekly", "time": {"hh": 9, "mm": 0}, "file": None},
        "fri": {"enabled": True,  "mode": "weekly", "time": {"hh": 9, "mm": 0}, "file": None},
        "sat": {"enabled": False, "mode": "weekly", "time": {"hh": 9, "mm": 0}, "file": None},
        "sun": {"enabled": False, "mode": "weekly", "time": {"hh": 9, "mm": 0}, "file": None},
    }
}

def load_config():
    if not os.path.exists(CONFIG_FILE):
        return json.loads(json.dumps(DEFAULT_CONFIG))
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        conf = json.loads(json.dumps(DEFAULT_CONFIG))
        conf["playlist"] = data.get("playlist", [])
        conf["weekly"]["base_index"] = int(data.get("weekly", {}).get("base_index", conf["weekly"]["base_index"]))
        # holidays
        hd = data.get("holidays", {})
        conf.setdefault("holidays", {"enabled": True, "dates": [], "ranges": []})
        conf["holidays"]["enabled"] = bool(hd.get("enabled", conf["holidays"]["enabled"]))
        # normalize date strings
        conf["holidays"]["dates"] = [str(x) for x in hd.get("dates", [])]
        conf["holidays"]["ranges"] = []
        for r in hd.get("ranges", []):
            try:
                conf["holidays"]["ranges"].append({"start": str(r.get("start")), "end": str(r.get("end"))})
            except Exception:
                pass
        for d in DAYS:
            src = data.get("schedule", {}).get(d, {})
            dst = conf["schedule"][d]
            dst["enabled"] = bool(src.get("enabled", dst["enabled"]))
            mode = src.get("mode", dst["mode"])
            if mode not in ("weekly","fixed"):
                mode = "weekly"
            dst["mode"] = mode
            t = src.get("time", {})
            dst["time"]["hh"] = int(t.get("hh", dst["time"]["hh"]))
            dst["time"]["mm"] = int(t.get("mm", dst["time"]["mm"]))
            dst["file"] = src.get("file", dst["file"])
        return conf
    except Exception:
        return json.loads(json.dumps(DEFAULT_CONFIG))

def save_config(conf):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(conf, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print('[설정 저장 오류]', e)

config = load_config()

# ----- System helpers -----
def terminate_processes(target_names):
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] in target_names:
                proc.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

def play_music_blocking(path):
    try:
        if not pygame.get_init():
            pygame.init()
        if not pygame.mixer.get_init():
            pygame.mixer.init()
        pygame.mixer.music.load(path)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            time.sleep(0.2)
    except Exception as e:
        print('[오디오오류]', e)

# ----- Enabled-days sequence helpers -----
def enabled_days():
    """enabled=True 인 요일만 순서대로 반환. 모두 꺼져 있으면 전체 요일 반환(안전)."""
    ed = [d for d in DAYS if config['schedule'].get(d, {}).get('enabled', False)]
    return ed if ed else DAYS[:]

def enabled_position(dkey):
    """dkey의 활성 요일 내 순서를 0..N-1로 반환. 비활성만 있는 경우 기본 DAYS에서 위치."""
    ed = enabled_days()
    return ed.index(dkey) if dkey in ed else DAYS.index(dkey)

def weekly_mapped_index(dkey, playlist_len):
    """활성 요일 기준 순서 + base_index 로 매핑된 playlist 인덱스."""
    if playlist_len <= 0:
        return None
    base = int(config.get('weekly', {}).get('base_index', 0)) % playlist_len
    pos = enabled_position(dkey)
    return (base + pos) % playlist_len

# ----- Playlist helpers -----
def playlist_options_with_labels():
    names = [os.path.basename(p) for p in config['playlist']]
    label_counts, labels = {}, []
    for name in names:
        c = label_counts.get(name, 0) + 1
        label_counts[name] = c
        labels.append(f'{name} (#{c})' if c > 1 else name)
    return [(labels[i], config['playlist'][i]) for i in range(len(config['playlist']))]

def find_path_by_label(label):
    for lab, path in playlist_options_with_labels():
        if lab == label:
            return path
    for p in config['playlist']:
        if os.path.basename(p) == label:
            return p
    return None

# ----- UI state -----
playlist_table = None
weekly_base_select = None
day_controls = {}      # dkey -> control dict
day_selected = {d: False for d in DAYS}
file_pick_queue = queue.Queue()
clipboard_day = None
tabs_ref = None
holiday_table = None
upcoming_table = None
next_run_label = None

# ----- File picker (front-most) -----
def _tk_pick_thread(multiple=True):
    import tkinter as tk
    from tkinter import filedialog
    import sys
    root = None
    try:
        root = tk.Tk()
        root.attributes('-alpha', 0.0)
        root.overrideredirect(True)
        root.geometry('0x0+0+0')
        root.lift()
        try: root.attributes('-topmost', True)
        except Exception: pass
        root.update()
        if sys.platform.startswith('win'):
            try:
                import ctypes
                hwnd = root.winfo_id()
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                ctypes.windll.user32.BringWindowToTop(hwnd)
            except Exception: pass
        if multiple:
            paths = filedialog.askopenfilenames(parent=root, filetypes=[('Audio Files','*.mp3 *.wav')], title='오디오 파일 선택')
        else:
            p = filedialog.askopenfilename(parent=root, filetypes=[('Audio Files','*.mp3 *.wav')], title='오디오 파일 선택')
            paths = (p,) if p else tuple()
    finally:
        try:
            if root is not None: root.destroy()
        except Exception: pass
    file_pick_queue.put(list(paths))

def consume_file_queue():
    try: paths = file_pick_queue.get_nowait()
    except queue.Empty: return
    if not paths: return
    before = len(config["playlist"])
    added = 0
    for p in paths:
        if p and p not in config["playlist"]:
            config["playlist"].append(p); added += 1
    if added:
        save_config(config); ui.notify(f'{added}개 추가됨')
    refresh_playlist_table()
    refresh_day_cards()
    refresh_upcoming_table()
    if before == 0 and added > 0 and tabs_ref is not None:
        try:
            tabs_ref.value = '요일 스케줄'
            ui.notify('요일별 스케줄에서 시간을 설정하세요')
        except Exception:
            pass

def pick_files():
    threading.Thread(target=_tk_pick_thread, kwargs={'multiple': True}, daemon=True).start()


# ----- Holiday table/actions -----
def refresh_holiday_table():
    global holiday_table
    if holiday_table is None: return
    rows = []
    hd = config.get('holidays', {})
    rid = 1
    # single dates
    for d in sorted(hd.get('dates', [])):
        rows.append({'rid': rid, 'type': '단일', 'label': d, 'start': d, 'end': ''}); rid += 1
    # ranges
    for r in hd.get('ranges', []):
        s = str(r.get('start')); e = str(r.get('end'))
        rows.append({'rid': rid, 'type': '기간', 'label': f'{s} ~ {e}', 'start': s, 'end': e}); rid += 1
    holiday_table.rows = rows
    holiday_table.update()

def add_holiday_date(d):
    if not d:
        return ui.notify('날짜를 선택하세요', color='warning')
    ds = str(d)[:10]
    hd = config.setdefault('holidays', {"enabled": True, "dates": [], "ranges": []})
    if ds not in hd['dates']:
        hd['dates'].append(ds)
        save_config(config)
        refresh_holiday_table(); refresh_upcoming_table()
        ui.notify(f'추가: {ds}')

def add_holiday_range(s, e):
    if not s or not e:
        return ui.notify('시작/종료 날짜를 모두 선택하세요', color='warning')
    ss, ee = str(s)[:10], str(e)[:10]
    if ee < ss:
        ss, ee = ee, ss
    hd = config.setdefault('holidays', {"enabled": True, "dates": [], "ranges": []})
    hd['ranges'].append({'start': ss, 'end': ee})
    save_config(config)
    refresh_holiday_table(); refresh_upcoming_table()
    ui.notify(f'추가: {ss} ~ {ee}')

def remove_selected_holiday():
    if holiday_table is None or not holiday_table.selected:
        return ui.notify('삭제할 항목을 선택하세요', color='warning')
    sel = holiday_table.selected[0]
    typ = sel.get('type'); s = sel.get('start'); e = sel.get('end')
    hd = config.setdefault('holidays', {"enabled": True, "dates": [], "ranges": []})
    if typ == '단일':
        try:
            hd['dates'].remove(s)
        except ValueError:
            pass
    else:
        hd['ranges'] = [r for r in hd.get('ranges', []) if not (str(r.get('start'))==s and str(r.get('end'))==e)]
    save_config(config)
    refresh_holiday_table(); refresh_upcoming_table()
    ui.notify('삭제됨')

# ----- Playlist actions -----
def refresh_playlist_table():
    rows = [{'idx': i+1, 'file': os.path.basename(p), 'path': p} for i, p in enumerate(config["playlist"])]
    if playlist_table is not None:
        playlist_table.rows = rows
        playlist_table.update()
    refresh_upcoming_table()

def selected_index():
    if playlist_table is None:
        return None
    sel = playlist_table.selected[0] if playlist_table.selected else None
    if not sel:
        ui.notify('행을 선택하세요', color='negative'); return None
    return sel.get('idx')-1 if isinstance(sel, dict) else int(sel)-1

def select_row_by_path(path):
    try:
        if playlist_table is None: return
        rows = playlist_table.rows or []
        target = next((r for r in rows if r['path'] == path), None)
        if target:
            playlist_table.selected = [target]
            playlist_table.update()
    except Exception: pass

def move_up():
    idx = selected_index()
    if idx is None: return
    if idx > 0:
        moved_path = config["playlist"][idx]
        config["playlist"][idx-1], config["playlist"][idx] = config["playlist"][idx], config["playlist"][idx-1]
        save_config(config); refresh_playlist_table(); select_row_by_path(moved_path)
        refresh_day_cards(); refresh_upcoming_table()

def move_down():
    idx = selected_index()
    if idx is None: return
    if idx < len(config["playlist"])-1:
        moved_path = config["playlist"][idx]
        config["playlist"][idx+1], config["playlist"][idx] = config["playlist"][idx], config["playlist"][idx+1]
        save_config(config); refresh_playlist_table(); select_row_by_path(moved_path)
        refresh_day_cards(); refresh_upcoming_table()

def remove_selected():
    idx = selected_index()
    if idx is None: return
    removed = config["playlist"].pop(idx); save_config(config); refresh_playlist_table()
    next_path = None
    if idx < len(config["playlist"]): next_path = config["playlist"][idx]
    elif len(config["playlist"]) > 0: next_path = config["playlist"][-1]
    if next_path: select_row_by_path(next_path)
    ui.notify(f'삭제: {os.path.basename(removed)}')
    refresh_day_cards(); refresh_upcoming_table()

# ----- Mini Player -----
# ----- Holiday helpers -----
def _date_str(dt):
    # dt: datetime.date | str
    if isinstance(dt, str): return dt[:10]
    return dt.strftime('%Y-%m-%d')

def is_holiday_date(date_str):
    """Return True if date_str(YYYY-MM-DD) is a configured holiday (single or within any range)."""
    try:
        ds = _date_str(date_str)
        # singles
        if ds in set(config.get('holidays', {}).get('dates', [])):
            return True
        # ranges (inclusive)
        for r in config.get('holidays', {}).get('ranges', []):
            s = r.get('start'); e = r.get('end')
            if not s or not e: continue
            if s <= ds <= e:
                return True
        return False
    except Exception:
        return False

_current_play = {"path": None, "paused": False}
def player_play():
    idx = selected_index()
    if idx is None: return
    path = config['playlist'][idx]
    if not os.path.exists(path): return ui.notify('파일이 존재하지 않습니다', color='negative')
    _current_play["path"] = path; _current_play["paused"] = False
    threading.Thread(target=play_music_blocking, args=(path,), daemon=True).start()
    ui.notify(f'재생: {os.path.basename(path)}')
def player_pause_resume():
    if not pygame.mixer.get_init(): pygame.mixer.init()
    if _current_play["path"] is None: return ui.notify('먼저 재생할 항목을 선택하세요', color='warning')
    if _current_play["paused"]: pygame.mixer.music.unpause(); _current_play["paused"] = False; ui.notify('재생 계속')
    else: pygame.mixer.music.pause(); _current_play["paused"] = True; ui.notify('일시정지')
def player_stop():
    if not pygame.mixer.get_init(): pygame.mixer.init()
    pygame.mixer.music.stop(); _current_play["paused"] = False; ui.notify('정지')

# ----- Day helpers (enabled-only mapping) -----
def resolved_file_label(dkey):
    dconf = config['schedule'][dkey]
    if dconf.get('mode','weekly') == 'fixed':
        p = dconf.get('file')
        return os.path.basename(p) if p else '(미지정)'
    pl = config.get('playlist', [])
    if not pl: return '(플레이리스트 비어있음)'
    idx = weekly_mapped_index(dkey, len(pl))
    if idx is None: return '(플레이리스트 비어있음)'
    return f'#{idx} → {os.path.basename(pl[idx])}'

def compute_play_path_for_day(dkey):
    dconf = config['schedule'][dkey]
    if dconf.get('mode','weekly') == 'fixed':
        return dconf.get('file')
    pl = config.get('playlist', [])
    if not pl: return None
    idx = weekly_mapped_index(dkey, len(pl))
    return pl[idx] if idx is not None else None

# ----- Upcoming schedule preview -----
def compute_upcoming_events(limit=8):
    """Return list of upcoming schedule entries skipping holidays and past times."""
    events = []
    now = datetime.now()
    holidays_enabled = config.get('holidays', {}).get('enabled', True)
    lookahead_days = max(limit * 7, 21)
    for offset in range(lookahead_days):
        if len(events) >= limit:
            break
        target_dt = now + timedelta(days=offset)
        dkey = DAYS[target_dt.weekday()]
        dconf = config.get('schedule', {}).get(dkey, {})
        date_str = target_dt.strftime('%Y-%m-%d')
        if holidays_enabled and is_holiday_date(date_str):
            continue
        if not dconf.get('enabled', False):
            continue
        scheduled_dt = target_dt.replace(
            hour=int(dconf.get('time', {}).get('hh', 9)),
            minute=int(dconf.get('time', {}).get('mm', 0)),
            second=0,
            microsecond=0,
        )
        if scheduled_dt <= now:
            continue
        mode = '수동 지정' if dconf.get('mode', 'weekly') == 'fixed' else '플레이리스트 순차'
        events.append({
            'datetime': scheduled_dt,
            'date': scheduled_dt.strftime('%Y-%m-%d'),
            'weekday': DAY_LABEL.get(dkey, dkey.upper()),
            'time': scheduled_dt.strftime('%H:%M'),
            'mode': mode,
            'file': resolved_file_label(dkey),
            'path': compute_play_path_for_day(dkey),
        })
    return events

def refresh_upcoming_table():
    global upcoming_table, next_run_label
    if upcoming_table is None:
        return
    events = compute_upcoming_events()
    rows = [
        {
            'idx': i + 1,
            'date': ev['date'],
            'weekday': ev['weekday'],
            'time': ev['time'],
            'mode': ev['mode'],
            'file': ev['file'],
        }
        for i, ev in enumerate(events)
    ]
    upcoming_table.rows = rows
    upcoming_table.update()
    if next_run_label is not None:
        if events:
            ev = events[0]
            next_run_label.text = f"다음 실행: {ev['date']}({ev['weekday']}) {ev['time']} · {ev['file']}"
        else:
            next_run_label.text = '예정된 실행 일정이 없습니다. 요일 스케줄을 활성화하세요.'
        next_run_label.update()

# ----- Copy/Paste/Test -----
def copy_day(dkey):
    global clipboard_day
    clipboard_day = json.loads(json.dumps(config['schedule'][dkey]))
    ui.notify(f'{DAY_LABEL[dkey]}요일 설정을 복사했습니다')

def paste_day(dkey):
    if clipboard_day is None:
        return ui.notify('복사된 설정이 없습니다', color='warning')
    config['schedule'][dkey] = json.loads(json.dumps(clipboard_day))
    save_config(config); refresh_day_cards()
    ui.notify(f'{DAY_LABEL[dkey]}요일에 붙여넣었습니다')

def test_play_day(dkey):
    path = compute_play_path_for_day(dkey)
    if not path or not os.path.exists(path):
        return ui.notify('재생할 파일이 없습니다', color='warning')
    threading.Thread(target=play_music_blocking, args=(path,), daemon=True).start()
    ui.notify(f'[테스트 재생] {DAY_LABEL[dkey]}: {os.path.basename(path)}')

# ----- Card refresh -----
def refresh_day_cards():
    labels = [lab for lab, _ in playlist_options_with_labels()] or ['(플레이리스트 비어있음)']
    if weekly_base_select is not None:
        base_opts = [str(i) for i in range(max(1, len(config['playlist'])))]
        weekly_base_select.options = base_opts
        weekly_base_select.value = str(int(config['weekly']['base_index']) % max(1, len(config['playlist'])))
        weekly_base_select.update()
    today_key = DAYS[datetime.now().weekday()]
    for d in DAYS:
        ctrls = day_controls.get(d)
        if not ctrls: continue
        dconf = config['schedule'][d]
        ctrls['enabled'].value = dconf.get('enabled', False); ctrls['enabled'].update()
        ctrls['mode'].value = ('플레이리스트 순차 적용' if dconf.get('mode','weekly')=='weekly' else '수동 지정'); ctrls['mode'].update()
        ctrls['hh'].value = f"{int(dconf.get('time',{}).get('hh',9)):02d}"; ctrls['hh'].update()
        ctrls['mm'].value = f"{int(dconf.get('time',{}).get('mm',0)):02d}"; ctrls['mm'].update()
        ctrls['file'].options = labels
        current_path = dconf.get('file'); current_lab = None
        for lab, p in playlist_options_with_labels():
            if p == current_path: current_lab = lab; break
        sel_val = current_lab if (current_lab in labels) else (labels[0] if labels else None)
        ctrls['file'].value = sel_val; ctrls['file'].update()
        is_fixed = (dconf.get('mode','weekly') == 'fixed')
        ctrls['file'].visible = is_fixed
        ctrls['resolved'].text = resolved_file_label(d); ctrls['resolved'].update()
        ctrls['sel'].value = day_selected[d]; ctrls['sel'].update()
        is_today = (d == today_key)
        card_classes = 'w-[300px] transition-all duration-300 '
        if dconf.get('enabled', False):
            card_classes += 'bg-blue-1 shadow-md '
        else:
            card_classes += 'bg-grey-1 '
        card_classes += 'border-2 border-primary' if is_today else 'border border-transparent'
        ctrls['card'].classes(replace=card_classes)
        ctrls['card'].update()
    refresh_upcoming_table()

# ----- Handlers -----
def on_toggle_select_all(on=True):
    for d in DAYS: day_selected[d] = bool(on); refresh_day_cards()
def on_select_weekdays():
    for i,d in enumerate(DAYS): day_selected[d] = (i < 5); refresh_day_cards()
def on_select_weekend():
    for i,d in enumerate(DAYS): day_selected[d] = (i >= 5); refresh_day_cards()
def on_toggle_select_day(dkey, newval):
    day_selected[dkey] = bool(newval)

def on_enabled_change(e, dkey):
    config['schedule'][dkey]['enabled'] = bool(e.value); save_config(config); refresh_day_cards()
def on_mode_change(e, dkey):
    config['schedule'][dkey]['mode'] = ('weekly' if e.value == '플레이리스트 순차 적용' else 'fixed'); save_config(config); refresh_day_cards()
def on_time_change_hh(e, dkey):
    config['schedule'][dkey]['time']['hh'] = int(e.value); save_config(config); refresh_upcoming_table()
def on_time_change_mm(e, dkey):
    config['schedule'][dkey]['time']['mm'] = int(e.value); save_config(config); refresh_upcoming_table()
def on_file_change(e, dkey):
    config['schedule'][dkey]['file'] = find_path_by_label(e.value); save_config(config); refresh_day_cards()
def on_weekly_base_change(e):
    try:
        config['weekly']['base_index'] = int(e.value)
        save_config(config); refresh_day_cards()
    except Exception: pass

def on_holiday_enabled_change(e):
    enabled = bool(e.value)
    config.setdefault('holidays', {"enabled": True, "dates": [], "ranges": []})
    config['holidays']['enabled'] = enabled
    save_config(config)
    refresh_upcoming_table()
    ui.notify('휴일 적용: ' + ('ON' if enabled else 'OFF'))

# ----- Batch apply -----
def selected_days():
    sel = [d for d in DAYS if day_selected[d]]
    return sel if sel else DAYS[:]

def batch_apply_time(hh, mm):
    for d in selected_days():
        config['schedule'][d]['time']['hh'] = int(hh)
        config['schedule'][d]['time']['mm'] = int(mm)
    save_config(config); refresh_day_cards(); ui.notify('시간 일괄 적용 완료')
def batch_apply_mode(mode):
    m = 'weekly' if mode == 'weekly' else 'fixed'
    for d in selected_days(): config['schedule'][d]['mode'] = m
    save_config(config); refresh_day_cards(); ui.notify(('순차 적용' if m=='weekly' else '수동 지정') + ' 일괄 적용 완료')
def batch_apply_enabled(val):
    for d in selected_days(): config['schedule'][d]['enabled'] = bool(val)
    save_config(config); refresh_day_cards(); ui.notify(('사용 켜기' if val else '사용 끄기') + ' 완료')
def batch_copy_from(day_key):
    global clipboard_day
    clipboard_day = json.loads(json.dumps(config['schedule'][day_key]))
    ui.notify(f'{DAY_LABEL[day_key]}요일 설정을 복사했습니다')
def batch_paste_to_selected():
    if clipboard_day is None: return ui.notify('복사된 설정이 없습니다', color='warning')
    for d in selected_days(): config['schedule'][d] = json.loads(json.dumps(clipboard_day))
    save_config(config); refresh_day_cards(); ui.notify('선택된 요일에 붙여넣었습니다')

# ----- UI -----
ACCENT = '#1e88e5'; ui.colors(primary=ACCENT)

with ui.header().classes('items-center'):
    ui.icon('schedule').style('color: white')
    ui.label('주간 재생 스케줄러').classes('text-white text-lg font-semibold')
    ui.space()
    ui.button(icon='save', color='white', on_click=lambda: (save_config(config), ui.notify('저장 완료'))).props('flat round')

# 컨테이너를 더 넓게 (max-w-7xl)
with ui.column().classes('w-full max-w-7xl mx-auto gap-4 p-4'):
    tabs_ref = ui.tabs().classes('w-full')
    with tabs_ref:
        ui.tab('플레이리스트')
        ui.tab('요일 스케줄')
        ui.tab('휴일/예외')
    with ui.tab_panels(tabs_ref, value='플레이리스트'):
        # --- Tab: Holidays ---
        with ui.tab_panel('휴일/예외'):
            with ui.card().classes('w-full'):
                ui.label('휴일/예외 설정').classes('text-base font-semibold')
                with ui.row().classes('items-center gap-3'):
                    ui.label('휴일 적용').classes('text-sm text-blue-grey-7')
                    sw_holiday = ui.switch(value=config.get('holidays',{}).get('enabled', True),
                                           on_change=on_holiday_enabled_change)
                ui.separator()

                # Single date add
                with ui.row().classes('items-center gap-2'):
                    ui.label('단일 날짜 추가').classes('text-sm text-blue-grey-7')
                    date_single = ui.date()
                    ui.button('추가', on_click=lambda: add_holiday_date(date_single.value)).props('outline dense')

                # Range add
                with ui.row().classes('items-center gap-2'):
                    ui.label('기간 추가').classes('text-sm text-blue-grey-7')
                    date_start = ui.date()
                    date_end = ui.date()
                    ui.button('추가', on_click=lambda: add_holiday_range(date_start.value, date_end.value)).props('outline dense')

            with ui.card().classes('w-full'):
                ui.label('휴일 목록').classes('text-base font-semibold')
                cols = [
                    {'name':'rid','label':'#','field':'rid','align':'center'},
                    {'name':'type','label':'종류','field':'type'},
                    {'name':'label','label':'표시','field':'label'},
                    {'name':'start','label':'시작','field':'start'},
                    {'name':'end','label':'종료','field':'end'},
                ]
                # module-level reference
                holiday_table = ui.table(columns=cols, rows=[], row_key='rid', selection='single').classes('w-full')
                with ui.row().classes('items-center gap-2 mt-2'):
                    ui.button('선택 삭제', on_click=remove_selected_holiday).props('outline')
                    ui.button('새로고침', on_click=refresh_holiday_table).props('outline')

        # --- Tab: Playlist (wider & taller) ---
        with ui.tab_panel('플레이리스트'):
            with ui.card().classes('w-full'):
                ui.label('플레이리스트').classes('text-base font-semibold')
                ui.label('여러 MP3/WAV 파일을 선택하여 목록에 추가합니다').classes('text-sm text-blue-grey-7')
                ui.button('파일 추가(다중)', icon='library_music', on_click=pick_files).props('outline')
                with ui.column().classes('w-full max-h-[640px] overflow-y-auto mt-2'):
                    columns = [
                        {'name':'idx','label':'#','field':'idx','align':'center'},
                        {'name':'file','label':'파일','field':'file'},
                        {'name':'path','label':'경로','field':'path'},
                    ]
                    playlist_table = ui.table(columns=columns, rows=[], row_key='idx', selection='single').classes('w-full')
                with ui.row().classes('w-full justify-between items-center mt-2'):
                    with ui.row().classes('items-center gap-2'):
                        ui.button(icon='play_arrow', on_click=player_play).props('round').tooltip('재생')
                        ui.button(icon='pause', on_click=player_pause_resume).props('round').tooltip('일시정지/계속')
                        ui.button(icon='stop', on_click=player_stop).props('round').tooltip('정지')
                    with ui.row().classes('items-center gap-2'):
                        ui.button('위로', icon='arrow_upward', on_click=move_up).props('outline')
                        ui.button('아래로', icon='arrow_downward', on_click=move_down).props('outline')
                        ui.button('선택 삭제', icon='delete_outline', on_click=remove_selected).props('outline')
                with ui.row().classes('justify-end'):
                    ui.button('다음: 요일 스케줄', icon='chevron_right', on_click=lambda: setattr(tabs_ref, 'value', '요일 스케줄')).props('unelevated color=primary')

        # --- Tab: Weekly Schedule ---
        with ui.tab_panel('요일 스케줄'):
            with ui.card().classes('w-full shadow-sm') as upcoming_card:
                upcoming_card.style('background: linear-gradient(135deg, rgba(30,136,229,0.12), rgba(30,136,229,0.04));')
                ui.label('다가오는 실행 일정').classes('text-base font-semibold')
                ui.label('휴일을 제외한 다음 실행 순서를 한눈에 확인하세요.').classes('text-sm text-blue-grey-7')
                next_run_label = ui.label('').classes('text-sm text-primary font-medium mt-1')
                cols = [
                    {'name': 'idx', 'label': '#', 'field': 'idx', 'align': 'center'},
                    {'name': 'date', 'label': '날짜', 'field': 'date'},
                    {'name': 'weekday', 'label': '요일', 'field': 'weekday', 'align': 'center'},
                    {'name': 'time', 'label': '시간', 'field': 'time', 'align': 'center'},
                    {'name': 'mode', 'label': '모드', 'field': 'mode', 'align': 'center'},
                    {'name': 'file', 'label': '재생 항목', 'field': 'file'},
                ]
                upcoming_table = ui.table(columns=cols, rows=[], row_key='idx').classes('w-full mt-2')
                ui.button('일정 새로고침', icon='refresh', on_click=refresh_upcoming_table).props('outline dense').classes('self-end mt-2')
            # 전역 설정 카드
            with ui.card().classes('w-full'):
                ui.label('전역 설정').classes('text-base font-semibold')
                with ui.row().classes('items-center gap-3'):
                    ui.label('주간 순차 시작 인덱스(활성 요일 기준, 월 기준 시작)').classes('text-sm text-blue-grey-7')
                    weekly_base_select = ui.select([str(i) for i in range(max(1,len(config['playlist'])))],
                                                   value=str(int(config['weekly']['base_index']) % max(1,len(config['playlist']))),
                                                   on_change=on_weekly_base_change).classes('min-w-[120px]')
            # 일괄 적용 카드 (범주별 묶음)
            with ui.card().classes('w-full'):
                ui.label('일괄 적용').classes('text-base font-semibold')
                with ui.grid(columns=4).classes('gap-3'):
                    with ui.card().classes('p-3'):
                        ui.label('선택 범위').classes('text-sm text-blue-grey-7 font-semibold')
                        ui.button('전체 선택', on_click=lambda: on_toggle_select_all(True)).props('outline dense')
                        ui.button('전체 해제', on_click=lambda: on_toggle_select_all(False)).props('outline dense')
                        ui.button('평일 선택', on_click=on_select_weekdays).props('outline dense')
                        ui.button('주말 선택', on_click=on_select_weekend).props('outline dense')
                    with ui.card().classes('p-3'):
                        ui.label('시간').classes('text-sm text-blue-grey-7 font-semibold')
                        batch_hh = ui.select([f'{h:02d}' for h in range(24)], value='09').classes('min-w-[90px]').props('label="시"')
                        batch_mm = ui.select([f'{m:02d}' for m in range(0,60,5)], value='00').classes('min-w-[90px]').props('label="분"')
                        ui.button('적용', on_click=lambda: batch_apply_time(batch_hh.value, batch_mm.value)).props('outline dense')
                    with ui.card().classes('p-3'):
                        ui.label('모드').classes('text-sm text-blue-grey-7 font-semibold')
                        ui.button('순차 적용', on_click=lambda: batch_apply_mode('weekly')).props('outline dense')
                        ui.button('수동 지정', on_click=lambda: batch_apply_mode('fixed')).props('outline dense')
                    with ui.card().classes('p-3'):
                        ui.label('사용').classes('text-sm text-blue-grey-7 font-semibold')
                        ui.button('켜기', on_click=lambda: batch_apply_enabled(True)).props('outline dense')
                        ui.button('끄기', on_click=lambda: batch_apply_enabled(False)).props('outline dense')
                with ui.row().classes('items-center gap-2 mt-2'):
                    ui.button('월 설정 복사', on_click=lambda: batch_copy_from('mon')).props('outline')
                    ui.button('선택에 붙여넣기', on_click=batch_paste_to_selected).props('outline')

            # 요일 카드 보드
            with ui.row().classes('gap-3 flex-wrap mt-2'):
                for dkey in DAYS:
                    dconf = config['schedule'][dkey]
                    card = ui.card().classes('w-[300px] bg-grey-1')
                    with card:
                        with ui.row().classes('items-center justify-between w-full'):
                            with ui.row().classes('items-center gap-2'):
                                sel_cb = ui.checkbox(value=day_selected[dkey], on_change=lambda e, k=dkey: on_toggle_select_day(k, e.value))
                                ui.chip(DAY_LABEL[dkey])
                            sw_enabled = ui.switch('사용', value=dconf.get('enabled', False), on_change=lambda e, k=dkey: on_enabled_change(e, k))
                        with ui.row().classes('gap-2 w-full'):
                            hh_sel = ui.select([f'{h:02d}' for h in range(24)],
                                               value=f"{int(dconf.get('time',{}).get('hh',9)):02d}",
                                               on_change=lambda e, k=dkey: on_time_change_hh(e, k)).classes('w-1/2').props('label="시"')
                            mm_sel = ui.select([f'{m:02d}' for m in range(0,60,5)],
                                               value=f"{int(dconf.get('time',{}).get('mm',0)):02d}",
                                               on_change=lambda e, k=dkey: on_time_change_mm(e, k)).classes('w-1/2').props('label="분"')
                        mode_sel = ui.select(['플레이리스트 순차 적용','수동 지정'],
                                             value=('플레이리스트 순차 적용' if dconf.get('mode','weekly')=='weekly' else '수동 지정'),
                                             on_change=lambda e, k=dkey: on_mode_change(e, k)).props('label="모드"').classes('w-full')
                        labels = [lab for lab, _ in playlist_options_with_labels()] or ['(플레이리스트 비어있음)']
                        current_path = dconf.get('file'); current_lab = None
                        for lab, p in playlist_options_with_labels():
                            if p == current_path: current_lab = lab; break
                        sel_val = current_lab if current_lab in labels else (labels[0] if labels else None)
                        file_sel = ui.select(labels, value=sel_val, on_change=lambda e, k=dkey: on_file_change(e, k)).props('label="수동 파일"').classes('w-full')
                        resolved = ui.label(resolved_file_label(dkey)).classes('text-blue-grey-7 text-sm')
                        with ui.row().classes('justify-between w-full'):
                            ui.button('복사', icon='content_copy', on_click=lambda k=dkey: copy_day(k)).props('flat')
                            ui.button('붙여넣기', icon='content_paste', on_click=lambda k=dkey: paste_day(k)).props('flat')
                            ui.button('테스트 재생', icon='play_arrow', on_click=lambda k=dkey: test_play_day(k)).props('flat')
                        day_controls[dkey] = {
                            'card': card,
                            'sel': sel_cb,
                            'enabled': sw_enabled,
                            'hh': hh_sel,
                            'mm': mm_sel,
                            'mode': mode_sel,
                            'file': file_sel,
                            'resolved': resolved,
                        }
                        file_sel.visible = (dconf.get('mode','weekly') == 'fixed')

# 초기 UI 동기화
refresh_playlist_table()
refresh_day_cards()
refresh_upcoming_table()

# timers / worker
ui.timer(0.3, consume_file_queue)
ui.timer(0.6, refresh_holiday_table)
ui.timer(60, refresh_upcoming_table)

def scheduler_loop():
    while True:
        try:
            dkey = DAYS[datetime.now().weekday()]
            dconf = config['schedule'].get(dkey, {})
            if config.get('holidays',{}).get('enabled', True):
                today = datetime.now().strftime('%Y-%m-%d')
                if is_holiday_date(today):
                    time.sleep(30); continue
            if not dconf.get('enabled'):
                time.sleep(30); continue
            hh, mm = datetime.now().hour, datetime.now().minute
            if hh == int(dconf.get('time',{}).get('hh', 9)) and mm == int(dconf.get('time',{}).get('mm', 0)):
                if dconf.get('mode','weekly') == 'fixed':
                    play_path = dconf.get('file')
                else:
                    pl = config.get('playlist', [])
                    play_path = None
                    if pl:
                        idx = weekly_mapped_index(dkey, len(pl))
                        play_path = pl[idx] if idx is not None else None
                if play_path and os.path.exists(play_path):
                    terminate_processes(['chrome.exe','msedge.exe','firefox.exe'])
                    play_music_blocking(play_path)
                    for s in servers:
                        try:
                            ssh = paramiko.SSHClient(); ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                            ssh.connect(s['host'], username=s['user'], password=s['password'])
                            ssh.exec_command(SHUTDOWN_CMD); ssh.close()
                        except Exception as e:
                            print(f"[{s['host']}] 실패:", e)
                    os.system(SHUTDOWN_CMD); return
                else:
                    print('[경고] 재생 파일이 설정되지 않았거나 존재하지 않습니다:', play_path)
            time.sleep(15)
        except Exception as e:
            print('[스케줄 오류]', e); time.sleep(15)

threading.Thread(target=scheduler_loop, daemon=True).start()

if __name__ in {'__main__','__mp_main__'}:
    try:
        if not pygame.get_init(): pygame.init(); pygame.mixer.init()
    except Exception: pass
    ui.run(title='주간 재생 스케줄러', native=True, window_size=(1100, 800))