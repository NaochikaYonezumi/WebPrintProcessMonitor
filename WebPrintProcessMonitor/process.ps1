# JSONから設定を取得する
$config = Get-Content -Path "$($PSScriptRoot)\config.json" | ConvertFrom-Json
$serverName = hostname
$logSettings = $config.logSettings
$logDir = Join-Path $PSScriptRoot 'logs'
$logFilePath = Join-Path $logDir ($logSettings.logFileBaseName + $logSettings.logFileExtension)
$logGenerations = $logSettings.logGenerations
$statusDir = Join-Path $PSScriptRoot 'Status'

# フォルダが存在しない場合は作成する関数
function Ensure-FolderExists {
    param (
        [string]$path
    )
    if (-not (Test-Path $path)) {
        New-Item -ItemType Directory -Path $path
    }
}

# ログを記録する関数
function Write-Log {
    param (
        [string]$level,
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $level - $message"
    Add-Content -Path $logFilePath -Value $logMessage
    Manage-LogFiles
}

# ログファイルの管理関数
function Manage-LogFiles {
    if ((Get-Item $logFilePath).Length -gt $logSettings.logFileMaxSize) {
        for ($i = $logGenerations; $i -ge 0; $i--) {
            $old = "$logDir\log$i.txt"
            if (Test-Path $old) {
                if ($i -eq $logGenerations) {
                    Remove-Item -Path $old -ErrorAction SilentlyContinue
                } else {
                    Rename-Item -Path $old -NewName "$logDir\log$($i + 1).txt" -ErrorAction SilentlyContinue
                }
            }
        }
        Rename-Item -Path $logFilePath -NewName "$logDir\log0.txt" -ErrorAction SilentlyContinue
        New-Item -Path $logFilePath -ItemType File -Force -ErrorAction SilentlyContinue
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
    $scriptPath = $PSScriptRoot
    $pythonScriptPath = Join-Path -Path $scriptPath -ChildPath "sendemail.exe"

    try {
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

# ステータスを確認するための関数
function CheckStatus {
    param (
        [System.Diagnostics.Process]$proc
    )
    $statusFileName = "$($proc.Id).status"
    $statusFilePath = Join-Path $statusDir $statusFileName
    $previousStatus = if (Test-Path $statusFilePath) { Get-Content $statusFilePath } else { $null }
    $data = [PSCustomObject]@{
        serverName   = $serverName
        Name     = $proc.Name
        Id       = $proc.Id
        runningTime  = $runningTime
    }
    $data | Export-Csv -Path $statusFilePath -NoTypeInformation
    return $previousStatus
}

# ステータスファイルをクリーンアップする関数
function Cleanup-StatusFiles {
    $allProcessIds = $allProcesses.Id
    Get-ChildItem -Path $statusDir -File | ForEach-Object {
        $fileName = $_.Name
        if ($fileName -match '^(\d+)\.status$') {
            $fileProcessId = [int]$matches[1]
            if (-not ($allProcessIds -contains $fileProcessId)) {
                $Proc = Import-Csv -Path $_.FullName
                Remove-Item $_.FullName -Force
                Write-Log -level "INFO" -message ($config.log.processStopSuccess -f $proc.Name, $proc.Id)
                $subject = $config.subject.processStopSuccess -f $serverName, $proc.Name, $proc.Id, $retryLimit
                $body = $config.body.processStopSuccess -f $serverName, $proc.Name, $proc.Id, $retryLimit
                Send-Mail -subject $subject -body $body
                Write-Log -level "INFO" -message "Remove status file: $fileName"
            }
        }
    }
}

# タスク自動終了時、ステータスファイルをクリーンアップする関数
function Cleanup-StatusFilesAfterProcessStopSuccess {
    param (
        [System.Diagnostics.Process]$proc
            )
    Get-ChildItem -Path $statusDir -File | ForEach-Object {
        $fileName = $_.Name
        if ($fileName -match '^(\d+)\.status$') {
            $fileProcessId = [int]$matches[1]
            if ($proc.Id -eq $fileProcessId){
                Remove-Item $_.FullName -Force
                Write-Log -level "INFO" -message "Remove status file: $fileName"
            }
        }
    }
}

# プロセスを取得する関数
function Get-AllProcesses {
    $allProcesses = @()
    foreach ($processName in $config.processNames) {
        $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
        $allProcesses += $processes
    }
    return $allProcesses
}

# プロセスが設定時間を超えた場合の処理を行う関数
function Handle-ProcessOverTime {
    param (
        [System.Diagnostics.Process]$proc,
        [double]$runningTime
    )
    Write-Log -level "ERROR" -message ($config.log.processtimeOver -f $proc.Name, $proc.Id, $runningTime)
    if ($config.stopProcess -eq "True") {
        Stop-ProcessAction -proc $proc -runningTime $runningTime
    } else {
        $previousStatus = CheckStatus -proc $proc -runningTime $runningTime
        if ($previousStatus -eq $null) {
            Send-ProcessStopFalseMail -proc $proc -runningTime $runningTime
        }
    }
}

# プロセス停止処理を行う関数
function Stop-ProcessAction {
    param (
        [System.Diagnostics.Process]$proc,
        [double]$runningTime
    )
    $bodyMessage = $config.body.processStopTrue -f $serverName, $proc.Name, $proc.Id, $runningTime
    $subjectMessage = $config.subject.processStopTrue -f $serverName, $proc.Name, $proc.Id, $runningTime
    Write-Log -level "INFO" -message ($config.log.processStopTrue -f $proc.Name, $proc.Id, $runningTime)
    $previousStatus = CheckStatus -proc $proc
    if ($previousStatus -eq $null) {
                Send-Mail -subject $subjectMessage -body $bodyMessage
    }
    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    Retry-StopProcess -proc $proc
}

# プロセス停止再試行処理を行う関数
function Retry-StopProcess {
    param (
        [System.Diagnostics.Process]$proc
    )
    $retryCount = 1
    $retryLimit = $config.retryLimit
    $retryInterval = $config.retryInterval

    while (($stoppedProcess = Get-Process -Id $proc.Id -ErrorAction SilentlyContinue) -and ($retryCount -le $retryLimit)) {
        Write-Log -level "INFO" -message ($config.log.processStopRetry -f $proc.Name, $proc.Id, $retryCount, $retryLimit)
        Start-Sleep -Seconds $retryInterval
        $retryCount++
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
    }

    if ($retryCount -gt $retryLimit) {
        if ($config.restartServer -eq "True") {
            Handle-RestartServer -proc $proc -retryLimit $retryLimit
        } else {
            $previousStatus = CheckStatus -proc $proc
            if ($previousStatus -eq $null) {
                Send-RebootFalseMail -proc $proc -retryLimit $retryLimit
            }
        }
    } else {
        Handle-ProcessStopSuccess -proc $proc -retryLimit $retryLimit
    }
}

# プロセス停止成功時のメール送信を行う関数
function Handle-ProcessStopSuccess {
    param (
        [System.Diagnostics.Process]$proc,
        [int]$retryLimit
    )
    Write-Log -level "INFO" -message ($config.log.processStopSuccess -f $proc.Name, $proc.Id)
    $subject = $config.subject.processStopSuccess -f $serverName, $proc.Name, $proc.Id, $retryLimit
    $body = $config.body.processStopSuccess -f $serverName, $proc.Name, $proc.Id, $retryLimit
    Send-Mail -subject $subject -body $body
    Cleanup-StatusFilesAfterProcessStopSuccess $proc
}

# プロセス停止失敗時のメール送信を行う関数
function Send-ProcessStopFalseMail {
    param (
        [System.Diagnostics.Process]$proc,
        [double]$runningTime
    )
    $bodyMessage = $config.body.processStopFalse -f $serverName, $proc.Name, $proc.Id, $runningTime
    $subjectMessage = $config.subject.processStopFalse -f $serverName, $proc.Name, $proc.Id, $runningTime
    Write-Log -level "INFO" -message ($config.log.processStopFalse -f $proc.Name, $proc.Id, $runningTime)
    Send-Mail -subject $subjectMessage -body $bodyMessage
}

# サーバ再起動処理を行う関数
function Handle-RestartServer {
    param (
        [System.Diagnostics.Process]$proc,
        [int]$retryLimit
    )
    Write-Output "Failed to stop the process. ProcessName: $($proc.Name) ProcessID: $($proc.Id) Retry count: $retryLimit"
    $subjectMessage = $config.subject.rebootTrue -f $serverName, $proc.Name, $proc.Id, $retryLimit
    $windowslogMessage = $config.body.rebootTrue -f $serverName, $proc.Name, $proc.Id, $retryLimit
    Write-Log -level "ERROR" -message ($config.log.rebootTrue -f $proc.Name, $proc.Id, $retryLimit)
    Send-Mail -subject $subjectMessage -body $windowslogMessage
    Try-StopPcWebPrint
    Start-Sleep -Seconds 30
    Start-Process PowerShell -ArgumentList "Restart-Computer -Force" -Verb RunAs
}

# pc-web-printプロセス停止を試みる関数
function Try-StopPcWebPrint {
    try {
        Write-Log -level "INFO" -message "Attempting to stop process: pc-web-print"
        $pcweb = Get-Process -Name pc-web-print -ErrorAction SilentlyContinue
        Write-Log -level "INFO" -message "pc-web: $pcweb"
        Stop-Process -Name pc-web-print -Force -ErrorAction SilentlyContinue
        Write-Log -level "INFO" -message "Process pc-web-print stopped successfully."
    } catch {
        Write-Log -level "ERROR" -message "Failed to stop process pc-web-print. Error: $_"
    }
}

# サーバ再起動設定をしていない時のメール送信を行う関数
function Send-RebootFalseMail {
    param (
        [System.Diagnostics.Process]$proc,
        [int]$retryLimit
    )
    $subjectMessage = $config.subject.rebootFalse -f $serverName, $proc.Name, $proc.Id, $retryLimit
    $windowslogMessage = $config.body.rebootFalse -f $serverName, $proc.Name, $proc.Id, $retryLimit
    Write-Log -level "ERROR" -message ($config.log.rebootFalse -f $proc.Name, $proc.Id, $retryLimit)
    Send-Mail -subject $subjectMessage -body $windowslogMessage
}

# メイン処理
function Main {
    Ensure-FolderExists -path $logDir
    Ensure-FolderExists -path $statusDir

    $allProcesses = Get-AllProcesses
    Cleanup-StatusFiles

    
    foreach ($proc in $allProcesses) {
        $runningTime = ((Get-Date) - $proc.StartTime).TotalMinutes

        if ($runningTime -ge $config.runningTimeThreshold) {
            Handle-ProcessOverTime -proc $proc -runningTime $runningTime
        } else {
            Write-Log -level "INFO" -message ($config.log.info -f $proc.Name, $proc.Id, $runningTime)
        }
    }
}

# メイン処理の実行
Main
