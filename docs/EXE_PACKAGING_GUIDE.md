# EXE 패키징 가이드 (개발/배포 담당자용)

이 문서는 `desktop_scheduler_qt.py`를 Windows 실행 파일(`AutoCloseStudio.exe`)로 빌드하고 배포하는 표준 절차를 설명합니다.

---

## 1. 사전 준비

- OS: Windows 권장(타깃 환경과 동일 버전 권장)
- Python: 3.10+
- 필수 패키지:
  - `PySide6`
  - `psutil`
  - `paramiko` (SSH 원격 종료 사용 시)
  - `pyinstaller`

설치 예시:

```powershell
python -m venv .venv
.\.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install PySide6 psutil paramiko pyinstaller
```

---

## 2. Spec 파일 핵심 이해

프로젝트는 `AutoCloseStudio.spec`를 기준으로 빌드합니다.

중요 항목:
- `ROOT`, `SRC_DIR`, `ASSET_DIR`
  - 현재 로컬 경로에 맞게 수정 필요
- `datas`
  - `assets` 폴더 포함
  - Qt 플러그인(`platforms`, `styles`, `audio`, `multimedia`) 포함
- `hiddenimports`
  - `PySide6.QtMultimedia` 등 명시
- `EXE(..., console=False, upx=False)`
  - GUI 앱, onefile 빌드

> 주의: spec 파일에는 `pathlib` pip backport 충돌 방지 가드가 있어, 문제 패키지가 있으면 빌드를 중단합니다.

---

## 3. 빌드 절차

### 3.1 경로 수정

`AutoCloseStudio.spec`에서 아래 값을 실제 저장소 위치로 변경:

- `ROOT`
- `SRC_DIR`

예시(개념):

```python
ROOT = Path(r"C:\work")
SRC_DIR = ROOT / "auto_close"
ASSET_DIR = SRC_DIR / "assets"
```

### 3.2 빌드 실행

```powershell
pyinstaller --clean AutoCloseStudio.spec
```

### 3.3 결과물 확인

- 최종 exe: `dist/AutoCloseStudio.exe`
- 중간 산출물: `build/AutoCloseStudio/`

권장 검증:
1. exe 실행
2. 트레이 아이콘 노출 확인
3. 로그인/화면 진입 확인
4. 테스트 음원 재생 확인
5. 강제 실행으로 종료 시퀀스 확인

---

## 4. 배포 패키지 구성 권장

onefile 기준 최소 구성:

```text
release/
└─ AutoCloseStudio.exe
```

운영 문서 포함 권장:

```text
release/
├─ AutoCloseStudio.exe
├─ USER_MANUAL.md
└─ 운영체크리스트.txt
```

---

## 5. 운영 환경 체크리스트

- Windows 종료 권한(로컬/원격) 확보 여부
- 방화벽/네트워크 정책으로 원격 종료 명령 허용 여부
- 오디오 출력 장치 존재 여부
- 시작 프로그램 등록 시 사용자 권한/정책 충돌 여부
- 백신/보안 솔루션의 exe 격리 여부

---

## 6. 트러블슈팅

## 6.1 빌드 시 `pathlib` 관련 오류

증상:
- spec 가드에서 `pip backport 'pathlib'` 충돌 메시지

조치:

```powershell
python -m pip uninstall pathlib
```

## 6.2 exe 실행 시 Qt 플랫폼 플러그인 오류

원인:
- Qt 플러그인 누락

조치:
- spec의 `qt_plugin_datas` 항목 확인
- clean 빌드 재시도

## 6.3 원격 종료 실패

점검:
- 대상 호스트/계정/비밀번호
- SSH 키 경로
- 네트워크 접근/방화벽
- WinRM 또는 SSH 서비스 활성화

## 6.4 오디오 재생 실패

점검:
- 음원 경로 유효성
- 코덱/파일 형식 지원
- 오디오 장치 상태

---

## 7. 릴리스 절차(권장)

1. 기능 검증/회귀 테스트
2. 버전 문자열 갱신(`APP_VERSION`, `BUILD_DATE`)
3. clean 빌드
4. smoke test(실행/로그인/재생/강제실행)
5. 문서 포함 릴리스 패키지 생성
6. 사내 배포/서명(필요 시)

