# Word → PNG 변환기

Word 문서(`.docx`, `.doc`)를 페이지별 PNG 이미지로 변환하는 Windows 데스크톱 앱입니다.

## 주요 기능

- **파일 추가** — 파일 탐색기 또는 드래그 앤 드롭으로 여러 파일 한 번에 추가
- **저장 폴더 지정** — 직접 입력하거나 폴더 선택 다이얼로그 사용 (미입력 시 원본 파일 위치에 `output/` 폴더 자동 생성)
- **진행률 표시** — 파일/페이지 단위로 진행 상태 실시간 표시
- **강제종료** — 변환 중 취소 시 백그라운드 Word 프로세스까지 즉시 종료

## 요구 사항

- Windows 10 / 11
- **Microsoft Word** 설치 필수 (COM 자동화 방식으로 변환)
- Python 3.9 이상 (소스 실행 시)

## 설치 및 실행

### 소스에서 실행

```bash
pip install -r requirements.txt
python main.py
```

### EXE 빌드

```bash
build.bat
```

빌드 완료 후 `dist/word_to_image.exe` 가 생성됩니다.

## 출력 파일 형식

```
{원본 파일명}_page_001.png
{원본 파일명}_page_002.png
...
```

## 기술 스택

| 항목 | 내용 |
|------|------|
| GUI | tkinter + tkinterdnd2 |
| Word → PDF | pywin32 (win32com.client) |
| PDF → PNG | PyMuPDF (fitz) |
| EXE 패키징 | PyInstaller |

## 주의사항

- 변환 시 Microsoft Word가 백그라운드에서 잠시 실행됩니다.
- 파일명에 특수문자(`()`, `..` 등)가 포함되어 있어도 정상 동작합니다.
- EXE 실행 시 최초 변환이 다소 느릴 수 있습니다.
