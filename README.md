# 概要
  - サンドボックス・サーバ上の特定のプロセスを監視し、もし監視しているプロセスが立ち上がり続けている場合は終了し管理者に通知メールを送るスクリプトです。<br>
  　プロセスが終了しない場合は、retryLimitに設定している回数分、プロセス終了コマンドを実行し、それでも終了しない場合はサーバを再起動します。<br>
    - デフォルトでは以下のプロセスが監視されます。<br>
        - "AcroRd32":32bitのAcrobatReader
        - "Acrobat":64bitのAcrobatReader
        - "WINWORD":Microsoft Office Word
        - "EXCEL":Microsoft Office Excel
        - "POWERPNT": Microsoft Office PowerPoint
# 設定方法
  設定は基本的にconfig.jsonファイルにて編集します。
  WebPrintStatusMonitorを任意のディレクトリに配置し、メモ帳などで編集するようにしてください。
## config.jsonの編集
```
{
    "smtpServer": "******", //SMTPサーバのホスト名 
    "smtpPort": "25", //SMTPサーバへ接続する際に使用するポート番号
    "from": "****@*****", //送信元のメールアドレス
    "recipients": [
        "****@*****"  //送信先の目メールアドレス(カンマ区切りで複数メールアドレス設定可能)
    ],
    "subject": {
        "processStopTrue": "WARN: ServerName: {0} ProcessName: {1} ID: {2} has been running for over {3} minutes.", //プロセスの時間が過ぎた場合に送信するメールの件名("stopProcess": "True"の場合)
        "processStopFalse": "WARN: ServerName: {0} ProcessName: {1} ID: {2} has been running for over {3} minutes.", //プロセスの時間が過ぎた場合に送信するメールの件名("stopProcess": "False"の場合)
        {0} サーバ名
        {1} プロセス名
        {2} プロセスID
        {3} 経過した分数

        "processStopSuccess" : "SUCCESS: ServerName: {0} ProcessName: {1} ID: {2} is Stopped.",//プロセスが正常に停止できた場合のメールの件名
        {0} サーバ名
        {1} プロセス名

        "rebootTrue": "ERROR: Failed to stop the process. Restart Server: {0} ProcessName: {1} ProcessID: {2} Retry count: {3}",//プロセス終了失敗後のサーバ再起動通知メールの件名("restartServer": "True"の場合)
        "rebootFalse": "ERROR: Failed to stop the process.The server will not reboot because the server restart setting is set to false.Server: {0} ProcessName: {1} ProcessID: {2} Retry count: {3}" //プロセス終了失敗後に通知メールの件名("restartServer": "False"の場合)
        {0} サーバ名
        {1} プロセス名
        {2} プロセスID
        {3} 再試行カウント
    },
    "body": {
        "processStopTrue": "ServerName:{0}\r\nProcessName:{1}\r\nProcessID:{2}\r\nhas been running for over {3} minutes.\r\nProcess is being stopped...", //プロセスの時間が過ぎた場合に送信するメールの本文("stopProcess": "True"の場合)
        "processStopFalse": "ServerName:{0}\r\nProcessName:{1}\r\nProcessID:{2}\r\nhas been running for over {3} minutes.",//プロセスの時間が過ぎた場合に送信するメールの本文("stopProcess": "False"の場合)
        {0} サーバ名
        {1} プロセス名
        {2} プロセスID
        {3} 経過した分数

        "processStopSuccess" : "ServerName:{0}\r\nProcessName:{1}\r\nProcessID:{2}\r\nThe process was terminated successfully.",//プロセスが正常に停止できた場合のメールの本文
        {0} サーバ名
        {1} プロセス名

        "rebootTrue": "ServerName:{0}\r\nProcessName:{1}\r\nProcessID:{2}\r\nRetry count:{3}\r\nFailed to stop the process.\r\nRestart Server.",//プロセス終了失敗後のサーバ再起動通知メールの本文("restartServer": "True"の場合)
        "rebootFalse": "ServerName:{0}\r\nProcessName:{1}\r\nProcessID:{2}\r\nRetry count:{3}\r\nFailed to stop the process."//プロセス終了失敗後に通知メールの本文("restartServer": "False"の場合)
        {0} サーバ名
        {1} プロセス名
        {2} プロセスID
        {3} 再試行カウント

    },
    "log":{ //ログに記録する内容
        "info": "ProcessName: {0} ProcessID: {1} has been running for about {2} minutes.",
        "processtimeOver": "ProcessName:{0} ID:{1} has been running for over {2} minutes.",
        "processStopTrue": "ProcessName:{0} ID:{1} has been running for over {2} minutes.Process is being stopped...",
        "processStopFalse": "ProcessName:{0} ID:{1} has been running for over {2} minutes.",
        "processStopRetry": "ProcessName:{0} ID:{1} Retry count:{2}/{3} Failed to stop the process... Retrying. ", 
        "processStopSuccess" : "ProcessName:{0} ID:{1} The process was terminated successfully.",
        "rebootTrue": "ProcessName:{0} ProcessID:{1} Retry count:{2} Failed to stop the process. Restart Server.",
        "rebootFalse": "ProcessName:{0} ProcessID:{1} Retry count:{2}Failed to stop the process."
    },
    "runningTimeThreshold": 10, //プロセスを終了するまでの経過時間
    "retryInterval": 10, //再試行までのインターバル(秒)
    "retryLimit": 6, //再試行回数
    "stopProcess": "True", //プロセスを終了するか
    "restartServer": "True", //再試行回数実施後、サーバを再起動するか
    "logSettings": { //ログの設定
        "logFileBaseName": "log", //ログ名
        "logFileExtension": ".txt", //ログの拡張子
        "logFileMaxSize": "10MB", //1世代のログの大きさ
        "logGenerations": 10 //ログの世代数
    },
    "processNames": [ //監視するプロセス
        "AcroRd32",
        "Acrobat",
        "WINWORD",
        "EXCEL",
        "POWERPNT"
    ]
}

```
## タスクスケジューラへの登録
  WebPrintProcessMonitorは「StartUpAcPM.vbs」をタスクスケジューラに登録することで定期的にステータスを監視することが可能です。<br>
  タスクスケジューラの登録は下記方法で実施します。

  1. タスクスケジューラを開きます。
  2. [タスクスケジューラ] - [タスクスケジューラライブラリ]の順番に開きます。
  3. 画面右にある[タスクの作成]をクリックします。
  4. 「タスクの作成」ウィンドウが開きます。<全般>を以下のように設定します。
     ```
     名前:任意の名前
     説明:任意の説明
     タスクの実行時に使うユーザアカウント:管理者権限を持つアカウント
     ユーザがログインしているかどうかにかかわらず実行する。
     最上位の特権で実行する:有効化
     ```
  5.  <トリガー>タブをクリックし、[新規]ボタンをクリックし、以下のように設定します。
     ```
    設定:1回
    開始:任意の日時
    詳細設定:
      ・繰り返し間隔:1分間(任意の時間間隔)
      ・継続時間:無期限
     ```
  6. ページ下部にある[OK]ボタンをクリックします。
  7. <操作>タブをクリックします。ページ下部にある[新規]ボタンをクリックします。
  8. 「新しい操作」ウィンドウが表示されます。以下のように設定します。
     ```
     操作:プログラムの開始
     プログラム/スクリプト:[App_Path]\WebPrintProcessMonitor\StartUpAcPM.vbs
     引数:なし
     開始(オプション):[App_Path]\WebPrintProcessMonitor\
     ```
  9. ページ下部にある[OK]ボタンをクリックします。
  10. <条件>タブをクリックします。以下のように設定します。
      ```
      電源:
        - コンピュータをAC電源で使用している場合のみタスクを開始する。:無効化
      ```
  11. <設定>タブをクリックします。以下のように設定します。
      ```
        - タスクを停止するまでの時間:無効化
      ```
  12. ページ下部にある[OK]ボタンをクリックします。
      
## StartUpWebPrintの設定
  "restartServer": "True"を設定している場合、再起動直後、Wドライブに接続できず、<br>
  Webプリントが正常に起動できない場合があります。その場合、同封している「StartUpWebPrint.bat」をタスクスケジューラに登録することにより、<br>
  Wドライブ接続後、Webプリントを起動することができます。<br>
  なお、Webプリントのインストール先がデフォルトと異なっている場合はbatの以下の行を書き換えてください。
  ```
@echo off
setlocal

:: ネットワークドライブのパスを指定します
set NETWORK_DRIVE_PATH=W:\

:: 試行回数を設定します
set /a RETRY_COUNT=6

:: 試行間隔（ミリ秒単位）を設定します
set INTERVAL=10000

:: アプリケーションのパスを指定します
set APP_PATH="C:\Program Files\PaperCut MF\providers\web-print\win\pc-web-print.exe" //この部分をWebプリントのインストール先に変更する。

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
```


  
  
