@echo off
chcp 65001 > nul
echo ============================================
echo   Word to PNG 변환기 - EXE 빌드 스크립트
echo ============================================
echo.

echo [1/3] 패키지 설치 중...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo 패키지 설치 실패! Python과 pip이 설치되어 있는지 확인하세요.
    pause
    exit /b 1
)

echo.
echo [2/3] EXE 빌드 중 (시간이 걸릴 수 있습니다)...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "word_to_image" ^
    --icon NONE ^
    --hidden-import=win32com ^
    --hidden-import=win32com.client ^
    --hidden-import=win32com.server ^
    --hidden-import=win32timezone ^
    --hidden-import=pythoncom ^
    --hidden-import=pywintypes ^
    main.py

if %errorlevel% neq 0 (
    echo 빌드 실패!
    pause
    exit /b 1
)

echo.
echo [3/3] 완료!
echo.
echo 생성된 파일: dist\Word_to_PNG.exe
echo.
echo ※ 주의사항:
echo   - 변환 시 Microsoft Word가 설치되어 있어야 합니다.
echo   - .doc/.docx 파일 변환을 위해 Word가 백그라운드에서 실행됩니다.
echo.
pause
