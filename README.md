# AutoClose Studio

AutoClose Studio는 **요일/시간 기반 자동 종료 스케줄러**입니다. 기본 동작은 아래 순서로 실행됩니다.

1. 지정된 시각에 스케줄 트리거
2. 대상 프로세스 강제 종료
3. 음성 파일 재생
4. 원격 PC 종료(옵션)
5. 로컬 PC 종료(옵션)

앱은 PySide6 GUI 기반으로 작성되어 있으며, 메인 창을 닫아도 트레이에서 계속 동작합니다.

---

## 1) 핵심 기능

- 요일별 스케줄(활성/비활성, 시간, 자동/수동 음원 지정)
- 플레이리스트 기반 자동 음원 할당(회전 인덱스)
- 휴일/기간 휴일 등록 및 휴일 동작 정책 분리
- 원격 종료(WinRM/SSH), 로컬 종료 지연 설정
- 일반 사용자/관리자 비밀번호 기반 모드 분리
- 시작 프로그램 등록(Windows)
- 설정 JSON 자동 저장/백업 및 저장 위치 변경

자세한 기능 설명은 [`docs/FEATURES_AND_ARCHITECTURE.md`](docs/FEATURES_AND_ARCHITECTURE.md)에서 확인할 수 있습니다.

---

## 2) 저장소 구조

```text
auto_close/
├─ desktop_scheduler_qt.py      # 메인 애플리케이션(단일 엔트리)
├─ AutoCloseStudio.spec         # PyInstaller onefile 빌드 스펙
├─ assets/
│  ├─ app_icon.ico              # 앱 아이콘
│  └─ topbar_logo.png           # 상단 로고
└─ docs/
   ├─ FEATURES_AND_ARCHITECTURE.md
   ├─ EXE_PACKAGING_GUIDE.md
   └─ USER_MANUAL.md
```

---

## 3) 빠른 실행(개발 환경)

> 권장: Python 3.10+

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install PySide6 psutil paramiko pyinstaller
python desktop_scheduler_qt.py
```

앱 첫 진입 시 기본 비밀번호:
- 일반 사용자: `0000`
- 관리자: `000000`

---

## 4) EXE 패키징 핵심 요약

이 프로젝트는 `AutoCloseStudio.spec`를 사용해 onefile exe를 생성합니다.

```bash
pyinstaller --clean AutoCloseStudio.spec
```

결과물:
- `dist/AutoCloseStudio.exe`

중요 포인트:
- spec 파일 내 `ROOT`, `SRC_DIR`는 현재 개발 환경 경로로 조정해야 합니다.
- PyInstaller와 충돌하는 `pathlib`(pip backport)가 설치되어 있으면 제거해야 합니다.

상세 절차와 체크리스트는 [`docs/EXE_PACKAGING_GUIDE.md`](docs/EXE_PACKAGING_GUIDE.md)를 참고하세요.

---

## 5) 사용자 운영 매뉴얼

실사용자 기준 화면별 설정/운영 절차는 [`docs/USER_MANUAL.md`](docs/USER_MANUAL.md)에 정리되어 있습니다.

포함 내용:
- 최초 실행
- 요일 스케줄 등록
- 플레이리스트/미리듣기
- 관리자 모드 진입 및 휴일/고급설정
- 트레이 동작, 잠금, 강제 실행
- 장애 대응(음원/원격 종료/권한 문제)

---

## 6) 개발자 문서

- 기능/아키텍처/주요 클래스/이벤트 흐름: [`docs/FEATURES_AND_ARCHITECTURE.md`](docs/FEATURES_AND_ARCHITECTURE.md)
- 배포/패키징/운영 체크리스트: [`docs/EXE_PACKAGING_GUIDE.md`](docs/EXE_PACKAGING_GUIDE.md)
- 사용자 입장에서의 화면 운영 절차: [`docs/USER_MANUAL.md`](docs/USER_MANUAL.md)

