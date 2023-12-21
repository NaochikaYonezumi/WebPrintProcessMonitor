@echo off
setlocal

:: ネットワークドライブのパスを指定します
set NETWORK_DRIVE_PATH=W:\

:: 試行回数を設定します
set /a RETRY_COUNT=6

:: 試行間隔（ミリ秒単位）を設定します
set INTERVAL=10000

:: アプリケーションのパスを指定します
set APP_PATH="C:\Program Files\PaperCut MF\providers\web-print\win\pc-web-print.exe"

:RETRY
:: ネットワークドライブにアクセスできるかどうかを確認します
dir %NETWORK_DRIVE_PATH% >nul 2>nul
if %ERRORLEVEL% equ 0 (
    :: アクセスできた場合、特定のアプリケーションが既に実行されているかどうかを確認します
    tasklist | find /i "%APP_PATH%" >nul
    if %ERRORLEVEL% equ 0 (
        :: アプリケーションが実行されている場合、それを終了します
        taskkill /f /im "pc-web-print.exe" >nul
        :: 少し待ちます
        timeout /t 3 /nobreak >nul
    )
    :: アプリケーションを起動します
    start "" %APP_PATH%
    goto END
)

:: 試行回数をデクリメントします
set /a RETRY_COUNT-=1

:: 試行回数が0になるまでループします
if %RETRY_COUNT% gtr 0 (
    :: 10秒間待ちます
    timeout /t 10 /nobreak >nul
    goto RETRY
)

:END
echo Unable to access network drive after 6 attempts.
endlocal
