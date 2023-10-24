Option Explicit

Dim objShell, strFolder, strCmd

' スクリプトのフォルダを取得
strFolder = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1)

' Shell オブジェクトを作成
Set objShell = CreateObject("Shell.Application")

' コマンドを定義（Powershellスクリプトをサイレントで実行）
strCmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strFolder & "\process.ps1"" "

' コマンドを実行（管理者権限で実行）
objShell.ShellExecute "powershell.exe", "-ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strFolder & "\process.ps1"" ", "", "runas", 0

' オブジェクトを解放
Set objShell = Nothing
