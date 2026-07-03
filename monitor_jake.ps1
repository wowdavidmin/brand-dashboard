# Jake 메일 체크 (크론용 - 최근 30개만 빠르게 확인)
$ErrorActionPreference = "Stop"
$SignalFile = "C:\Users\user\.openclaw\workspace\jake_mail_signal.txt"
$LogFile = "C:\Users\user\.openclaw\workspace\jake_mail_log.txt"

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    
    # 최근 30개만 확인 (정렬)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    
    $limit = [Math]::Min(30, $items.Count)
    for ($i = 1; $i -le $limit; $i++) {
        $mail = $items.Item($i)
        if ($mail.Unread) {
            $senderName = $mail.SenderName
            $senderEmail = $mail.SenderEmailAddress
            
            if ($senderName -like "*Jake*" -or $senderEmail -like "*jake*") {
                $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                $subject = $mail.Subject
                $received = $mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
                
                # 시그널 파일 작성
                "$timestamp||$senderName/$senderEmail||$subject||$received" | Out-File -FilePath $SignalFile -Encoding UTF8 -Force
                
                # 로그
                Add-Content -Path $LogFile -Value "[$timestamp] FOUND|$senderName/$senderEmail|$subject|$received"
                
                # 읽음 처리
                $mail.Unread = $false
                $mail.Save()
                
                Write-Host "JAKE_MAIL_FOUND"
                Write-Host "SENDER: $senderName / $senderEmail"
                Write-Host "SUBJECT: $subject"
                exit 0
            }
        }
    }
    
    Write-Host "NO_NEW_JAKE_MAIL"
}
catch {
    Add-Content -Path $LogFile -Value "[ERROR] $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $($_.Exception.Message)"
    Write-Host "CHECK_ERROR:$($_.Exception.Message)"
}
finally {
    if ($outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
}
