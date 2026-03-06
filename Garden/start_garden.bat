@echo off
setlocal
chcp 65001 >nul
title Home Vegetable Garden System

REM ==== venv311（実体の場所）====
call "C:\Users\spax2\venv311\Scripts\activate.bat"
if errorlevel 1 (
  echo [ERROR] venv311 の activate に失敗しました。
  pause
  exit /b 1
)

REM ==== 農作業フォルダへ移動（OneDrive配下）====
cd /d "%OneDrive%\ドキュメント\PythonWork\農作業"
if errorlevel 1 (
  echo [ERROR] 農作業フォルダへ移動できません: %OneDrive%\ドキュメント\PythonWork\農作業
  pause
  exit /b 1
)

echo.
echo =======================================
echo  Starting garden apps (run_all.py) ...
echo  Open: http://localhost:8501  (start_page)
echo =======================================
echo.

REM ==== run_all.py を起動（Streamlit画面配下）====
python "%OneDrive%\ドキュメント\PythonWork\農作業\Streamlit画面\run_all.py"

echo.
echo To stop all apps, press Ctrl + C in this window.
pause
endlocal



