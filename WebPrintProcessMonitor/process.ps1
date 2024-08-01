# JSONから設定を取得する
$config = Get-Content -Path "$($PSScriptRoot)\config.json" | ConvertFrom-Json
$serverName = hostname
$logSettings = $config.logSettings
$subjectMessages = $config.subject
$bodyMessages = $config.body
$logMessages = $config.log
$stopProcess = $config.stopProcess
$restartServer = $config.restartServer
$LogDir = Join-Path $PSScriptRoot 'logs' # ログフォルダのパスを定義
$LogFileName = $logSettings.logFileBaseName + $logSettings.logFileExtension
$LogFilePath = Join-Path $LogDir $LogFileName 
$logGenerations = $logSettings.logGenerations

# ログを記録する関数
function Write-Log {
    param (
        [string]$level,
        [string]$message
    )
    # タイムスタンプの取得
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # ログメッセージの作成
    $logMessage = "$timestamp - $level - $message"
    
    # ログファイルにメッセージを書き込む
    Add-Content -Path $LogFilePath -Value $logMessage

    # ログファイルサイズと世代の管理
    if ((Get-Item $LogFilePath).Length -gt $logSettings.logFileMaxSize) {
        for ($i = $logGenerations; $i -ge 0; $i--) {
            $old = "$LogDir\log$i.txt"
            if (Test-Path $old) {
                if ($i -eq $logGenerations) {
                    # 最古のログファイルを削除
                    Remove-Item -Path $old -ErrorAction SilentlyContinue
                } else {
                    # 古いログファイルの名前を変更
                    $new = "$LogDir\log$($i + 1).txt"
                    Rename-Item -Path $old -NewName $new -ErrorAction SilentlyContinue
                }
            }
        }
        # 現在のログファイルを新しい世代としてリネームし、新しいログファイルを作成
        Rename-Item -Path $LogFilePath -NewName "$LogDir\log0.txt" -ErrorAction SilentlyContinue
        New-Item -Path $LogFilePath -ItemType File -Force -ErrorAction SilentlyContinue
    }
}

# メールを送信するための関数
function Send-Mail {
    param (
        [string]$subject,
        [string]$body
    )
    $smtpServer = $config.smtpServer
    $smtpPort = $config.smtpPort
    $smtpUser = $config.smtpUser
    $smtpPassword = $config.smtpPassword
    $fromAddress = $config.from
    $toAddresses = [string]::Join(',', $config.recipients)
        
    # 相対パスでPythonスクリプトを指定
    $scriptPath = $PSScriptRoot
    $pythonScriptPath = Join-Path -Path $scriptPath -ChildPath "sendemail.exe"

    try {
        # Pythonスクリプトを呼び出し
        $result = & $pythonScriptPath $smtpServer $smtpPort $fromAddress $toAddresses $subject $body
        Write-Output $result
        if ($result -like "*successfully sent*") {
            Write-Log -level "INFO" -message "The email has been successfully sent."
        } else {
            Write-Log -level "ERROR" -message "The email could not be sent. Error: $result"
        }
    } catch {
        Write-Log -level "ERROR" -message "The email could not be sent. Exception: $_"
    }
}

# main

# ログフォルダがない場合は作成
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir }

# プロセス名からプロセスの取得
$allProcesses = @()
foreach ($processName in $config.processNames) {
    $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
    $allProcesses += $processes
}

# プロセスの取得
foreach ($proc in $allProcesses) {
    $runningTime = ((Get-Date) - $proc.StartTime).TotalMinutes

    if ($runningTime -ge $config.runningTimeThreshold) {
        Write-Log -level "ERROR" -message ($logMessages.processtimeOver -f $proc.Name, $proc.Id, $runningTime)
        if ($stopProcess -eq "True") {
            $bodyMessage = $bodyMessages.processStopTrue -f $serverName, $proc.Name, $proc.Id, $runningTime
            $subjectMessage = $subjectMessages.processStopTrue -f $serverName, $proc.Name, $proc.Id, $runningTime
            Write-Log -level "INFO" -message ($logMessages.processStopTrue -f $proc.Name, $proc.Id, $runningTime)
            # メールの送信
            Send-Mail -subject $subjectMessage -body $bodyMessage
            # プロセスの停止
            #Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        } else {
            $bodyMessage = $bodyMessages.processStopFalse -f $serverName, $proc.Name, $proc.Id, $runningTime
            $subjectMessage = $subjectMessages.processStopFalse -f $serverName, $proc.Name, $proc.Id, $runningTime
            Write-Log -level "INFO" -message ($logMessages.processStopFalse -f $proc.Name, $proc.Id, $runningTime)
            # メールの送信
            Send-Mail -subject $subjectMessage -body $bodyMessage
            exit
        }
        $retryCount = 1
        $retryLimit = $config.retryLimit  # 再試行限界回数
        $retryInterval = $config.retryInterval  # 再試行間隔

        while (($stoppedProcess = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue) -and ($retryCount -le $retryLimit)) {
            Write-Log -level "INFO" -message ($logMessages.processStopRetry -f $proc.Name, $proc.Id, $retryCount, $retryLimit)
            Write-Output $logMessages.processStopRetry -f $proc.Name, $proc.Id, $retryCount, $retryLimit
            Start-Sleep -Seconds $retryInterval

            # インクリメント
            $retryCount++

            # プロセス停止（再試行）
            #Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        }

        if ($retryCount -gt $retryLimit) {
            if ($restartServer -eq "True") {
                Write-Output "Failed to stop the process. ProcessName: $($proc.Name) ProcessID: $($proc.Id) Retry count: $retryLimit"
                $subjectMessage = $subjectMessages.rebootTrue -f $serverName, $proc.Name, $proc.Id, $retryLimit
                $windowslogMessage = $bodyMessages.rebootTrue -f $serverName, $proc.Name, $proc.Id, $retryLimit
                Write-Log -level "ERROR" -message ($logMessages.rebootTrue -f $serverName, $proc.Name, $proc.Id, $retryLimit)

                # メールの送信
                Send-Mail -subject $subjectMessage -body $windowslogMessage

                #アプリを停止する。
                try {
                    Write-Log -level "INFO" -message "Attempting to stop process: pc-web-print"
                    $pcweb = Get-Process -Name pc-web-print -ErrorAction SilentlyContinue
                    Write-Log -level "INFO" -message "pc-web: $pcweb"
                    Stop-Process -Name pc-web-print -Force -ErrorAction SilentlyContinue
                    Write-Log -level "INFO" -message "Process pc-web-print stopped successfully."
                } catch {
                    Write-Log -level "ERROR" -message "Failed to stop process pc-web-print. Error: $_"
                }
                
                # 30秒待機
                Start-Sleep -Seconds 30
                
                Start-Process PowerShell -ArgumentList "Restart-Computer -Force" -Verb RunAs
            } else {
                $subjectMessage = $subjectMessages.rebootFalse -f $serverName, $proc.Name, $proc.Id, $retryLimit
                $windowslogMessage = $bodyMessages.rebootFalse -f $serverName, $proc.Name, $proc.Id, $retryLimit
                Write-Log -level "ERROR" -message ($logMessages.rebootFalse -f $serverName, $proc.Name, $proc.Id, $retryLimit)

                # メールの送信
                Send-Mail -subject $subjectMessage -body $windowslogMessage
                exit
            }
        } else {
            # プロセスが正常に停止した場合
            Write-Log -level "INFO" -message ($logMessages.processStopSuccess -f $proc.Name, $proc.Id)
            $subject = $subjectMessages.processStopSuccess -f $serverName, $proc.Name, $proc.Id, $retryLimit
            $body = $bodyMessages.processStopSuccess -f $serverName, $proc.Name, $proc.Id, $retryLimit
            # メールを送信
            Send-Mail -subject $subject -body $body
        }
    } else {
        Write-Log -level "INFO" -message ($logMessages.info -f $proc.Name, $proc.Id, $runningTime)
    }
}
