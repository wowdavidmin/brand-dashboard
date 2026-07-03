# Jake 메일 실시간 모니터 + 데스크탑 알림 (지속 실행)
# 30초마다 Outlook 확인, Jake 메일 발견 시 알림
$ErrorActionPreference = "Stop"
$SignalFile = "C:\Users\user\.openclaw\workspace\jake_mail_signal.txt"
$LogFile = "C:\Users\user\.openclaw\workspace\jake_mail_log.txt"

Write-Host "=== JAKE MAIL MONITOR STARTED ==="
Write-Host "Email: david.min@yakjin.com"
Write-Host "Target: Jake (jake@yakjin.com)"
Write-Host "Interval: 30 sec"

while ($true) {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)
        
        $items = $inbox.Items
        $items.Sort("[ReceivedTime]", $true)
        
        $limit = [Math]::Min(20, $items.Count)
        for ($i = 1; $i -le $limit; $i++) {
            $mail = $items.Item($i)
            if ($mail.Unread) {
                $senderName = $mail.SenderName
                $senderEmail = $mail.SenderEmailAddress
                
                if ($senderName -like "*Jake*" -or $senderEmail -like "*jake*") {
                    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    $subject = $mail.Subject
                    $received = $mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
                    
                    $logMsg = "[$timestamp] JAKE|$senderName/$senderEmail|$subject|$received"
                    Add-Content -Path $LogFile -Value $logMsg
                    Write-Host $logMsg
                    
                    # 시그널 파일
                    "$timestamp||$senderName/$senderEmail||$subject||$received" | Out-File -FilePath $SignalFile -Encoding UTF8 -Force
                    
                    # 데스크탑 알림 (팝업)
                    $popup = New-Object -ComObject Wscript.Shell
                    $popup.Popup("Jake 메일 도착!`n제목: $subject`n수신: $received", 10, "📬 Jake 메일 알림", 48) | Out-Null
                    
                    # 읽음 처리
                    $mail.Unread = $false
                    $mail.Save()
                }
            }
        }
        
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
    catch {
        $errMsg = "[ERROR] $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $($_.Exception.Message)"
        Add-Content -Path $LogFile -Value $errMsg
    }
    
    Start-Sleep -Seconds 30
}
