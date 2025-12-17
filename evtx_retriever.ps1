Import-Module CredentialManager
$script:reportMsg = ""

$filePath = "C:\admin\evtx-retriever"
$encPath = "${filePath}\encrypted"
$outDir = Join-Path $filePath "evtx"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$smtpServer = "ccmail.kyoto-su.ac.jp"
$smtpPort = 25
$smtpUseSSL = $false
$smtpFrom = "center-windows@cc.kyoto-su.ac.jp"
$smtpTo = "kshinjo@cc.kyoto-su.ac.jp"
$smtpUser = ""                      # e.g. smtp_user
$smtpPasswordFile = "${encPath}\smtp_password.txt"  # optional
$smtpCred = $null
if ($smtpServer -and $smtpUser -and (Test-Path $smtpPasswordFile)) {
    try {
        $securePass = Get-Content $smtpPasswordFile | ConvertTo-SecureString -key (1..16)
        $smtpCred = New-Object System.Management.Automation.PSCredential($smtpUser, $securePass)
    } catch {
        Write-Warning "Failed to build SMTP credential from ${smtpPasswordFile}: $($_.Exception.Message)"
    }
}

function Send-NotificationEmail($subject, $body) {
    if (-not $smtpServer -or -not $smtpTo) {
        Write-Warning "SMTP settings not configured, skipping notification: $subject"
        return
    }

    $mailParams = @{
        SmtpServer = $smtpServer
        Port = $smtpPort
        From = $smtpFrom
        To = $smtpTo
        Subject = $subject
        Body = $body
    }
    if ($smtpUseSSL) { $mailParams['UseSsl'] = $true }
    if ($smtpCred) { $mailParams['Credential'] = $smtpCred }

    try {
        Send-MailMessage @mailParams -ErrorAction Stop
        Write-Host "Notification email sent: $subject"
    } catch {
        Write-Warning "Failed to send notification email: $($_.Exception.Message)"
    }
}

function ProcessServerListFile($path) {
    $servers = (Get-Content $path) -as [String[]]
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($path)
    Write-Host "[$(Get-Date)] Processing server list file: $path. Found $($servers.Count) entries. (type: $baseName)"
    foreach ($server in $servers) {
        Write-Host "[$(Get-Date)] -> Queueing server: $server"
        logHandler $server $baseName
    }
}


function logHandler($server, $baseName) {
    $cred = Get-StoredCredential -Target $baseName

    $serverDir = Join-Path $outDir $server # 出力ディレクトリ（サーバごと）
    if (-not (Test-Path $serverDir)) {
        New-Item -ItemType Directory -Path $serverDir | Out-Null
        Write-Host "Created directory: $serverDir"
    }

    # Export only the Security log for the last 24 hours
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logsToExport = @('Security')

    foreach ($logName in $logsToExport) {
        $remoteFile = "C:\Windows\Temp\evtx_export_${timestamp}_${logName.Replace('/','_')}.evtx"
        $localFile  = Join-Path $serverDir ("${logName.Replace('/','_')}_${timestamp}.evtx")

        try {
            Invoke-Command -ComputerName $server -Credential $cred -ArgumentList $logName, $remoteFile -ScriptBlock {
                param($logName, $remotePath)
                $query = "*[System[TimeCreated[timediff(@SystemTime) <= 86400000]]]"
                Write-Output "Exporting $logName to $remotePath with time-only query: $query"
                wevtutil.exe epl $logName $remotePath /q:"$query"
            } -ErrorAction Stop
            Write-Output "[$(Get-Date)] Successfully exported $logName from ${server} to ${remoteFile}"

            # Use a PSSession and Copy-Item -FromSession to stream the file to the collector
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

                # Clean up remote temporary file (preferred using the active PSSession)
                try {
                    Invoke-Command -Session $session -ArgumentList $remoteFile -ScriptBlock { param($p) Remove-Item -LiteralPath $p -Force -ErrorAction Stop } -ErrorAction Stop
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
        } finally {
            # Ensure remote temp file is removed if we successfully fetched it locally.
            if ($localFile -and (Test-Path $localFile)) {
                try {
                    Invoke-Command -ComputerName $server -Credential $cred -ArgumentList $remoteFile -ScriptBlock {
                        param($p)
                        if (Test-Path $p) { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue }
                    } -ErrorAction SilentlyContinue
                    Write-Host "[$(Get-Date)] Confirmed removal of remote file $remoteFile on $server"
                } catch {
                    Write-Warning "[$(Get-Date)] Could not confirm removal of remote file $remoteFile on ${server}: $($_.Exception.Message)"
                }
            }
        }
    }
}

Get-ChildItem -Path ${filePath}\hosts | ForEach-Object { ProcessServerListFile $_.FullName }

robocopy $outDir "C:\admin\evtx-watchdog\evtx" /E /MOVE /R:1 /W:1
# Clean up old files in each server subdirectory (older than 1 days)
Get-ChildItem -Path "C:\admin\evtx-watchdog\evtx" -Directory | ForEach-Object {
    $serverDir = $_.FullName
    Get-ChildItem -Path $serverDir -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-1) } | Remove-Item -Force
}
