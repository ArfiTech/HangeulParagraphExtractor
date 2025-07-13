# 한글 프로그램 찾는 단어 포함 문단 추출 README

## 개요

한글 프로그램 내에서 원하는 단어가 포함된 문단을 추출하여, 문단별 HWPML 파일(.hwpml)로 생성하는 자동화 스크립트를 제공합니다. 생성된 `.hwpml` 파일은 한글(한컴오피스)에서 열어 **다른 이름으로 저장 → HWP(.hwp)** 로 변환하면 최종 `.hwp` 파일을 얻을 수 있습니다.

---

## Setting

- **Windows 10/11** 전용
- 한컴오피스 한글 2020 이상 설치 (HWPX 저장 기능 필요)
- *Python과 한글 프로그램의 비트 수를 일치* (예: Python 64bit ↔ HWP 64bit)
  - COM 자동화(`pywin32`)를 사용하는 경우 필수
- **Python 3.7 이상** 설치 (3.10 권장)
- 가상환경 생성(venv) 및 활성화 권장
  ```powershell
  python -m venv .venv
  .\.venv\Scripts\Activate.ps1   # PowerShell
  # 또는  .\.venv\Scripts\activate.bat   # CMD
  ```
- 표준 라이브러리만 사용 (별도 설치 불필요):
  - `zipfile`, `xml.etree.ElementTree`, `os`, `sys`
- (구 버전 COM 방식 사용 시)
  ```powershell
  pip install pywin32
  ```

---

## 사용 방법

1. 한글 프로그램에서 ``hwp → hwpx`` 저장

   - 한글 열기 → 파일 > 다른 이름으로 저장 > **HWPX (\*.hwpx)** 선택
   - 예: `한글파일.hwpx`

2. 코드 실행

   - 스크립트 파일: `extract.py`
   - 명령:
     ```powershell
     python extract.py 한글파일.hwpx
     ```
   - 실행 결과: 작업 디렉터리에 `OUTPUT_DIR`에 입력한 이름의 폴더가 생성되고, 내부에 `CHAR_NAMES`에 입력한 단어로 `단어.hwpml` 파일이 저장됩니다.

3. 최종 ``hwpml → hwpx`` 변환

   - 한글 프로그램에서 각 `.hwpml` 파일 열기
   - 파일 > 다른 이름으로 저장 > **HWP (\*.hwp)** 선택

---

## 참고
- 추출 시 **옛한글 자모**, **한자 병기**, **각주**가 모두 **원본 그대로** 보존됩니다.
- 비트 수 불일치 또는 한컴오피스 미설치 시 스크립트가 정상 동작하지 않을 수 있습니다.