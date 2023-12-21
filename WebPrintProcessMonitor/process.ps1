#JSONから設定を取得する。
$config = Get-Content -Path "$($PSScriptRoot)\config.json" | ConvertFrom-Json
$serverName = hostname
$logSettings = $config.logSettings
$subjectMessages = $config.subject
$bodyMessages = $config.body
$logMessages = $config.log
$stopProcess = $config.stopProcess
$restartServer = $config.restartServer
$stopProcess = $config.stopProcess
$restartServer = $config.restartServer

#ログパスの生成
function Get-LogFileName {
    $logFileNumber = 0
    while (Test-Path -Path "$($PSScriptRoot)\logs\$($logSettings.logFileBaseName)$logFileNumber$($logSettings.logFileExtension)") {
        $logFileNumber++
    }
    return "$($PSScriptRoot)\logs\$($logSettings.logFileBaseName)$($logFileNumber-1)$($logSettings.logFileExtension)"
}

# logsフォルダが存在しない場合、作成する
if (!(Test-Path -Path "$($PSScriptRoot)\logs")) {
    New-Item -Path "$($PSScriptRoot)\logs" -ItemType Directory
}

$logFile = Get-LogFileName

#ログを保存するディレクトリがなければ作成する。
if (!(Test-Path $logFile)) {
    New-Item -Path $logFile -ItemType File
}

#ログロテートの設定
$logFileMaxSizeBytes = [int]([double]::Parse(($logSettings.logFileMaxSize -replace 'MB', '')) * 1MB)
if ((Get-Item -Path $logFile).Length -gt $logFileMaxSizeBytes) {
    $logFiles = Get-ChildItem -Path "$($PSScriptRoot)\logs" -Filter "$($logSettings.logFileBaseName)*$($logSettings.logFileExtension)" | Sort-Object Name
    if ($logFiles.Count -ge $logSettings.logGenerations) {
        Remove-Item -Path $logFiles[0].FullName
    }
    $logFile = "$($PSScriptRoot)\logs\$($logSettings.logFileBaseName)$logFileNumber$($logSettings.logFileExtension)"
    if (!(Test-Path $logFile)) {
        New-Item -Path $logFile -ItemType File
    }
}

# メールを送信するための関数
function Send-Mail($subject, $body) {
    $smtpServer = $config.smtpServer
    $smtpPort = $config.smtpPort
    $message = New-Object System.Net.Mail.MailMessage
    $message.From = $config.from
    $message.Subject = $subject
    $message.Body = $body

    foreach ($recipient in $config.recipients) {
        $message.To.Add($recipient)
    }

    # SMTP クライアントオブジェクトの作成
    $smtpClient = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)

    # メール送信
    try {
        $smtpClient.Send($message)
        Write-Log -level "INFO" -message "The email has been successfully sent."
    } catch {
        Write-Log -level "ERROR" -message "The email could not be sent."
    }
}

#プロセス名からプロセスの取得
$allProcesses = @()
foreach ($processName in $config.processNames) {
    # Get processes matching the process name
    $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
    $allProcesses += $processes
}

#main
#プロセスの取得
foreach ($proc in $allProcesses) {
    $runningTime = ((Get-Date) - $proc.StartTime).TotalMinutes
    if ($runningTime -ge $config.runningTimeThreshold) {
        $logLevel = "ERROR"
        $logMessage = $logMessages.processtimeOver -f $proc.Name, $proc.Id, $runningTime
        Add-Content -Path $logFile -Value ("$(Get-Date) - " + $logLevel + " - " + $logMessage)
        if ($stopProcess -eq "True"){
            $bodyMessage = $bodyMessages.processStopTrue -f $serverName, $proc.Name, $proc.Id, $runningTime
            $subjectMessage = $subjectMessages.processStopTrue -f $serverName, $proc.Name, $proc.Id, $runningTime
            $logMessage = $logMessages.processStopTrue -f $proc.Name, $proc.Id, $runningTime
            Add-Content -Path $logFile -Value ("$(Get-Date) - " + $logLevel + " - " + $logMessage)
            Write-Output $logMessage
            #メールの送信
            Send-Mail $subjectMessage $bodyMessage
            #プロセスの停止
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        } else {
            $bodyMessage = $bodyMessages.processStopFalse -f $serverName, $proc.Name, $proc.Id, $runningTime
            $subjectMessage = $subjectMessages.processStopFalse -f $serverName, $proc.Name, $proc.Id, $runningTime
            $logMessage = $logMessages.processStopFalse -f $proc.Name, $proc.Id, $runningTime
            Add-Content -Path $logFile -Value ("$(Get-Date) - " + $logLevel + " - " + $logMessage)
            Write-Output $logMessage
            #メールの送信
            Send-Mail $subjectMessage $bodyMessage
            exit
        }
        $retryCount = 1
        $retryLimit = $config.retryLimit  #再試行限界回数
        $retryInterval = $config.retryInterval  #再試行間隔

        while (($stoppedProcess = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue) -and ($retryCount -le $retryLimit)) {
            $logMessage = $logMessages.processStopRetry -f $proc.name, $proc.Id, $retryCount, $retryLimit
            Add-Content -Path $logFile -Value ("$(Get-Date) - " + $logLevel + " - " + $logMessage)
            Write-Output $logMessage

            # インターバル
            Start-Sleep -Seconds $retryInterval

            # Increment retry count
            $retryCount++

            # プロセス停止（再試行)
            Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        }

        if ($retryCount -gt $retryLimit) {
            if($restartServer -eq "True") {
                Write-Output "Failed to stop the process. ProcessName:$($proc.name) ProcessID:$($proc.Id) Retry count:$($retryLimit)"
                $subjectMessage = $subjectMessages.rebootTrue -f $serverName, $proc.Name, $proc.Id, $retryLimit
                $windowslogMessage = $bodyMessages.rebootTrue -f $serverName, $proc.Name, $proc.Id, $retryLimit
                $logMessage = $logMessages.rebootTrue -f $proc.Name, $proc.Id, $retryLimit
                Add-Content -Path $logFile -Value ("$(Get-Date) - ERROR - " + $logMessage)

                #メールの送信
                Send-Mail $subjectMessage $windowslogMessage

                #アプリを停止する。
                Stop-Process -Name pc-web-print　

                #10秒待機
                Start-Sleep -Seconds 30
                
                #プロセスが終了できない場合サーバを再起動する｡
                Start-Process PowerShell -ArgumentList "Restart-Computer -Force" -Verb RunAs

            } else {
                $subjectMessage = $subjectMessages.rebootFalse -f $serverName, $proc.Name, $proc.Id, $retryLimit
                $windowslogMessage = $bodyMessages.rebootFalse -f $serverName, $proc.Name, $proc.Id, $retryLimit
                $logMessage = $logMessages.rebootFalse -f $proc.Name, $proc.Id, $retryLimit
                Add-Content -Path $logFile -Value ("$(Get-Date) - ERROR - " + $logMessage)
                Write-Output $logMessage

                #メールの送信
                Send-Mail $subjectMessage $windowslogMessage
                exit
            }
            
        } else {
         # プロセスが正常に停止した場合
            $logMessage = $logMessages.processStopSuccess -f $proc.Name, $proc.Id
            Add-Content -Path $logFile -Value ("$(Get-Date) - INFO - " + $logMessage)
            $subject = $subjectMessages.processStopSuccess -f $serverName, $proc.Name, $proc.Id, $retryLimit
            $body = $bodyMessages.processStopSuccess -f $serverName, $proc.Name, $proc.Id, $retryLimit
            # メールを送信
            Send-Mail $subject $body
            Write-Output "Process $($proc.Name) $($proc.Id) has been successfully stopped."
        }
    } else {
        $logMessage = $logMessages.info -f $proc.name, $proc.Id, $runningTime
        Write-Output $logMessage
        Add-Content -Path $logFile -Value ("$(Get-Date) - INFO - " + $logMessage)
    }
}
