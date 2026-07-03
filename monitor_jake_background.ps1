# Jake 메일 실시간 모니터링 (NewMailEx 이벤트 방식)
# 이 스크립트는 백그라운드에서 실행되며 새 메일이 오면 이벤트로 감지합니다.
$ErrorActionPreference = "Stop"
$LogFile = "C:\Users\user\.openclaw\workspace\jake_mail_log.txt"
$SignalFile = "C:\Users\user\.openclaw\workspace\jake_mail_signal.txt"

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    Write-Host "MONITOR_STARTED: Jake 메일 모니터링 시작됨"
    
    # 이벤트 등록
    $eventIdentifier = Register-ObjectEvent -InputObject $outlook -EventName "NewMailEx" -Action {
        $entryId = $Event.SourceArgs[0]
        try {
            $ol = $Event.MessageData.Outlook
            $ns = $ol.GetNamespace("MAPI")
            $mail = $ns.GetItemFromID($entryId)
            $senderName = $mail.SenderName
            $senderEmail = $mail.SenderEmailAddress
            $subject = $mail.Subject
            $received = $mail.ReceivedTime.ToString("MM-dd HH:mm")
            
            if ($senderName -like "*Jake*" -or $senderEmail -like "*jake*") {
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $logMsg = "[$timestamp] JAKE|$senderName/$senderEmail|$subject|$received"
                Add-Content -Path $Event.MessageData.LogFile -Value $logMsg
                
                # 시그널 파일 생성 (크론이 감지)
                "$timestamp||$senderName||$senderEmail||$subject||$received" | Out-File -FilePath $Event.MessageData.SignalFile -Encoding UTF8 -Force
            }
        }
        catch {
            $errMsg = "[EVENT_ERROR] $($_.Exception.Message)"
            Add-Content -Path $Event.MessageData.LogFile -Value $errMsg
        }
    } -MessageData @{ Outlook = $outlook; LogFile = $LogFile; SignalFile = $SignalFile }
    
    # 스크립트가 종료되지 않도록 유지
    while ($true) {
        Start-Sleep -Seconds 5
        # 30초마다 한 번씩 로그에 heartbeat
        $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "RUNNING:$now"
        Start-Sleep -Seconds 25
    }
}
catch {
    $errMsg = "[FATAL] $($_.Exception.Message)"
    Add-Content -Path $LogFile -Value $errMsg
    Write-Host "FATAL_ERROR:$errMsg"
}
finally {
    if ($eventIdentifier) { Unregister-Event -SourceIdentifier $eventIdentifier.Name -ErrorAction SilentlyContinue }
    if ($outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
}
