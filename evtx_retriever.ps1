Import-Module CredentialManager
$script:reportMsg = ""

$filePath = "C:\admin\evtx-retriever"
$outDir = Join-Path $filePath "evtx"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }


function Send-NotificationEmail($subject, $body) {
    $mailParams = @{
        SmtpServer = "ccmail.kyoto-su.ac.jp"
        Port       = 25
        From       = "center-windows@cc.kyoto-su.ac.jp"
        To         = "kshinjo@cc.kyoto-su.ac.jp"
        Subject    = $subject
        Body       = $body
    }

    try {
        Send-MailMessage @mailParams -ErrorAction Stop
        Write-Host "Notification email sent: $subject"
    } catch {
        Write-Warning "Failed to send notification email: $($_.Exception.Message)"
    }
}

function ProcessServerListFile($path) {
    $servers  = (Get-Content $path) -as [String[]]
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($path)
    Write-Host "[$(Get-Date)] Processing server list file: $path. Found $($servers.Count) entries. (type: $baseName)"
    foreach ($server in $servers) {
        Write-Host "[$(Get-Date)] -> Queueing server: $server"
        logHandler $server $baseName
    }
}

function logHandler($server, $baseName) {
    $cred      = Get-StoredCredential -Target $baseName
    $serverDir = Join-Path $outDir $server

    if (-not (Test-Path $serverDir)) {
        New-Item -ItemType Directory -Path $serverDir | Out-Null
        Write-Host "Created directory: $serverDir"
    }

    # Security ログのみ、直近24時間分を取得
    $timestamp    = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logsToExport = @('Security')

    foreach ($logName in $logsToExport) {
        $remoteFile = "C:\Windows\Temp\evtx_export_${timestamp}_${($logName.Replace('/','_'))}.evtx"
        $localFile  = Join-Path $serverDir ("${($logName.Replace('/','_'))}_${timestamp}.evtx")

        try {
            # 直近24時間のクエリ（86400000ミリ秒）。※HTMLエスケープなしで <= を使用。
            Invoke-Command -ComputerName $server -Credential $cred -ArgumentList $logName, $remoteFile -ScriptBlock {
                param($logName, $remotePath)
                $query = "*[System[TimeCreated[timediff(@SystemTime) <= 86400000]]]"
                Write-Output "Exporting $logName to $remotePath with time-only query: $query"
                wevtutil.exe epl $logName $remotePath /q:"$query"
            } -ErrorAction Stop
            Write-Output "[$(Get-Date)] Successfully exported $logName from ${server} to ${remoteFile}"

            # PSSessionでコピー（ストリーム転送）
            $session = $null
            try {
                try {
                    $session = New-PSSession -ComputerName $server -Credential $cred -ErrorAction Stop
                } catch {
                    $msg = "Failed to establish PSSession to ${server}: $($_.Exception.Message)"
                    Write-Warning "[$(Get-Date)] $msg"
                    Send-NotificationEmail "EVTX session error on $server" "$msg`nLog: $logName`nTime: $(Get-Date)"
                    continue
                }

                try {
                    Copy-Item -FromSession $session -Path $remoteFile -Destination $localFile -Force -ErrorAction Stop
                    Write-Host "[$(Get-Date)] Saved EVTX to $localFile"
                } catch {
                    $msg = "Failed to copy remote file $remoteFile from $server to ${localFile}: $($_.Exception.Message)"
                    Write-Warning "[$(Get-Date)] $msg"
                    Send-NotificationEmail "EVTX copy failure on $server" "$msg`nLog: $logName`nTime: $(Get-Date)"
                    continue
                }

                # リモート一時ファイルの削除（PSSession経由で1回のみ）
                try {
                    Invoke-Command -Session $session -ArgumentList $remoteFile -ScriptBlock {
                        param($p)
                        Remove-Item -LiteralPath $p -Force -ErrorAction Stop
                    } -ErrorAction Stop
                    Write-Host "[$(Get-Date)] Removed remote temp file $remoteFile via PSSession on $server"
                } catch {
                    $msg = "Failed to remove remote temp file $remoteFile via PSSession on ${server}: $($_.Exception.Message)"
                    Write-Warning "[$(Get-Date)] $msg"
                    Send-NotificationEmail "EVTX remote deletion failure (PSSession) on $server" "$msg`nLog: $logName`nLocalFile: $localFile`nTime: $(Get-Date)"
                }
            } finally {
                if ($session) { Remove-PSSession -Session $session -ErrorAction SilentlyContinue }
            }
        } catch {
            Write-Warning "[$(Get-Date)] Failed exporting $logName from ${server}: $($_.Exception.Message)"
        }
    }
}

# ホストリストに対して処理
Get-ChildItem -Path ${filePath}\hosts | ForEach-Object { ProcessServerListFile $_.FullName }

# watchdogから過去のログを削除
Get-ChildItem -Path "C:\admin\evtx-watchdog\evtx" -Directory -ErrorAction Stop | ForEach-Object {
    $dir = $_.FullName
    Write-Host "[$(Get-Date)] Removing directory and contents: $dir"
    Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction Stop
}

# 収集結果をwatchdog側へ移動
robocopy $outDir "C:\admin\evtx-watchdog\evtx" /E /MOVE /R:1 /W:1

# 半年以上前のレポートディレクトリを削除
Get-ChildItem -Path 'C:\admin\evtx-watchdog\reports' -Directory |
    Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-180) } |
    Remove-Item -Recurse -Force -ErrorAction Continue
