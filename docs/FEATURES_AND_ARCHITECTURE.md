# 기능 및 아키텍처 문서 (개발자용)

## 1. 개요

`desktop_scheduler_qt.py`는 단일 파일 기반 PySide6 애플리케이션으로, 설정 저장/스케줄 계산/오디오 재생/종료 시퀀스를 모두 포함합니다.

핵심 설계 포인트:
- **설정 중심 구조**: `SchedulerConfig` + `ConfigManager`
- **백그라운드 스케줄러**: `SchedulerEngine`의 별도 스레드 루프
- **오디오 완료 기반 후속 동작**: `AudioService` 시그널로 종료 단계 연결
- **GUI 상태 분리**: 사용자/관리자 모드, 잠금 상태, 트레이 숨김 동작

---

## 2. 모듈 단위 역할

## 2.1 설정/도메인 모델

- `DaySchedule`: 요일별 1개 일정 상태
  - `enabled`, `time`, `auto_assign`, `audio_path`, `allow_remote`, `allow_local_shutdown`, `last_ran`
- `SchedulerConfig`: 앱 전체 상태
  - 플레이리스트, 대상 프로세스 목록, 원격 호스트, 휴일 정책, 테마, 비밀번호 해시, wakeup 옵션 등

설정 JSON 직렬화/역직렬화는 각 모델의 `as_dict()` / `from_dict()`로 수행됩니다.

## 2.2 설정 저장 계층

- `ConfigLocator`
  - 기본 설정 경로: `%APPDATA%` 또는 `$XDG_CONFIG_HOME` 또는 `~/.config`
  - 포인터 파일(`storage_location.json`)로 실제 설정 폴더 위치를 추적
- `ConfigManager`
  - `settings.json` + `settings.json.bak` + 임시파일 원자적 치환(`.tmp`)
  - 로드 실패 시 백업 복구 시도
  - 마이그레이션 포인트(`_apply_migrations`) 제공

## 2.3 스케줄 계산/트리거

- `compute_upcoming_runs()`
  - 미래 일정(기본 28일)을 계산
  - 휴일/주말 제외 정책 반영
  - `last_ran`으로 중복 실행 방지
  - 자동 할당 시 플레이리스트 회전 인덱스 적용
- `SchedulerEngine`
  - 약 15초 주기로 다음 실행 계산 + 트리거 조건 확인
  - 트리거 시 `schedule_triggered(day_key, audio, allow_remote, allow_local)` 시그널 송출

## 2.4 실행 액션

- `terminate_programs(targets)`
  - 지정 프로세스 종료
- `AudioService`
  - Qt Multimedia로 음원 재생/중지/오류 시그널 처리
- `shutdown_remote(hosts)`
  - `winrm` 또는 `ssh(paramiko)` 방식 종료 명령
- `shutdown_local(delay)`
  - Windows: `shutdown /s /t <delay>`
  - Linux/macOS: 대기 후 `shutdown -h now`

종료 시퀀스는 “트리거 → 프로세스 종료 → 음원 재생 완료 이벤트 → 원격/로컬 종료” 형태입니다.

## 2.5 UI 계층

- 커스텀 카드 컴포넌트(`FancyCard`, `DayCard`, `PlaylistPanel`, `HolidayPanel`, `SettingsPanel` 등)
- `MainWindow`
  - 화면 전환, 잠금 상태, 트레이 메뉴, 강제 실행, 로그 갱신
- `App`
  - 단일 인스턴스 보장(`QSharedMemory`)
  - 로그인 다이얼로그 라우팅, 초기 트레이 숨김

---

## 3. 실행 흐름(런타임 시퀀스)

1. `App` 시작 → 단일 인스턴스 검사
2. `ConfigManager`가 설정 로드/복구
3. `MainWindow` 초기화 후 트레이로 숨김
4. 사용자가 트레이에서 열기 → 비밀번호 인증
5. `SchedulerEngine`가 다음 실행시각 계산 및 대기
6. 스케줄 시각 도달 시 실행 체인 시작
   - 대상 프로세스 종료
   - 음원 재생
   - 재생 완료 시 원격/로컬 종료 실행
7. 실행 결과는 로그 카드 및 트레이 메시지로 노출

---

## 4. 설정 파일 구조 요약

저장 파일: `settings.json`

대표 키:
- `days.mon ~ days.sun`: 요일별 시간/옵션
- `playlist`, `playlist_rotation`
- `targets`(종료 대상 프로세스)
- `remote_hosts`(host, username, password, method, command 등)
- `holidays`, `holiday_ranges`, `auto_skip_weekends`
- `shutdown_delay`, `enable_remote_shutdown`, `enable_local_shutdown`
- `user_password_hash`, `admin_password_hash`

운영 권장:
- JSON 수동 편집보다는 UI 편집을 기본으로 사용
- 원격 종료 계정 정보는 운영망 보안 정책에 맞게 관리

---

## 5. 확장/개선 포인트

- 현재 단일 파일 구조를 패키지 구조로 분리
  - `core/config.py`, `core/scheduler.py`, `ui/main_window.py` 등
- 원격 종료 어댑터 인터페이스화
  - `WinRMAdapter`, `SSHAdapter` 분리 및 테스트 가능 구조
- 설정 스키마 버전 도입
  - 마이그레이션 이력 관리 고도화
- CI에서 정적 분석/유닛 테스트 자동화

---

## 6. 주의 사항(코드 관찰 기반)

- 진입점에 `--saver-only` 옵션이 존재하며 `_run_screensaver_only(...)`를 호출하도록 되어 있으나, 현재 파일에서 동일 이름 함수 정의를 확인하기 어렵습니다.
- 배포 전 해당 옵션의 실제 사용 여부를 점검하거나, 함수 구현/가드 로직을 정리하는 것이 안전합니다.

