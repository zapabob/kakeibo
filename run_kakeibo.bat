@echo off
chcp 65001 > nul
echo ======================================
echo    家計簿アプリケーション起動ツール
echo ======================================

:: Python実行ファイルを探す
set PYTHON_PATH=

:: 一般的なインストール場所を確認
if exist "C:\Python39\python.exe" (
    set PYTHON_PATH=C:\Python39\python.exe
) else if exist "C:\Python38\python.exe" (
    set PYTHON_PATH=C:\Python38\python.exe
) else if exist "C:\Python37\python.exe" (
    set PYTHON_PATH=C:\Python37\python.exe
) else if exist "C:\Program Files\Python39\python.exe" (
    set PYTHON_PATH="C:\Program Files\Python39\python.exe"
) else if exist "C:\Program Files\Python38\python.exe" (
    set PYTHON_PATH="C:\Program Files\Python38\python.exe"
) else if exist "C:\Program Files\Python37\python.exe" (
    set PYTHON_PATH="C:\Program Files\Python37\python.exe"
) else if exist "C:\Program Files (x86)\Python39\python.exe" (
    set PYTHON_PATH="C:\Program Files (x86)\Python39\python.exe"
) else if exist "C:\Program Files (x86)\Python38\python.exe" (
    set PYTHON_PATH="C:\Program Files (x86)\Python38\python.exe"
) else if exist "C:\Program Files (x86)\Python37\python.exe" (
    set PYTHON_PATH="C:\Program Files (x86)\Python37\python.exe"
) else if exist "%LOCALAPPDATA%\Programs\Python\Python39\python.exe" (
    set PYTHON_PATH="%LOCALAPPDATA%\Programs\Python\Python39\python.exe"
) else if exist "%LOCALAPPDATA%\Programs\Python\Python38\python.exe" (
    set PYTHON_PATH="%LOCALAPPDATA%\Programs\Python\Python38\python.exe"
) else if exist "%LOCALAPPDATA%\Programs\Python\Python37\python.exe" (
    set PYTHON_PATH="%LOCALAPPDATA%\Programs\Python\Python37\python.exe"
)

:: Python Launcherを使用
if "%PYTHON_PATH%"=="" (
    echo Python Launcherで実行します...
    py -3 kakeibo.py
    goto check_error
)

:: 見つかったPythonで実行
echo Pythonパス: %PYTHON_PATH%
%PYTHON_PATH% kakeibo.py

:check_error
if %ERRORLEVEL% neq 0 (
    echo.
    echo エラーが発生しました (エラーコード: %ERRORLEVEL%)
    echo [必要なライブラリ]
    echo  - PyQt6
    echo  - pandas
    echo  - schedule
    echo.
    echo [インストール方法]
    echo  pip install PyQt6 pandas schedule
    echo.
    pause
) else (
    echo アプリケーションが終了しました
) 