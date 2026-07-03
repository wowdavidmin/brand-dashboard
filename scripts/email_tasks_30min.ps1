# 30분마다 실행되는 메일 작업
$ErrorActionPreference = "Stop"
$SignalFile = "C:\Users\user\.openclaw\workspace\jake_mail_signal.txt"
$LogFile = "C:\Users\user\.openclaw\workspace\jake_mail_log.txt"
$Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Task 1: Jake mail check
try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $inbox = $namespace.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]", $true)
    $limit = [Math]::Min(30, $items.Count)
    $foundJake = $false
    for ($i = 1; $i -le $limit; $i++) {
        $mail = $items.Item($i)
        if ($mail.Unread) {
            $senderName = $mail.SenderName
            $senderEmail = $mail.SenderEmailAddress
            if ($senderName -like "*Jake*" -or $senderEmail -like "*jake*") {
                $subject = $mail.Subject
                $received = $mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
                "$Timestamp||$senderName/$senderEmail||$subject||$received" | Out-File -FilePath $SignalFile -Encoding UTF8 -Force
                Add-Content -Path $LogFile -Value "[$Timestamp] JAKE|$senderName/$senderEmail|$subject|$received"
                $mail.Unread = $false
                $mail.Save()
                $foundJake = $true
                break
            }
        }
    }
    if (-not $foundJake) {
        Add-Content -Path $LogFile -Value "[$Timestamp] NO_NEW_JAKE_MAIL"
    }
}
catch {
    Add-Content -Path $LogFile -Value "[$Timestamp] JAKE_ERR: $($_.Exception.Message)"
}
finally {
    if ($outlook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null }
}

# Task 2: Brand folder mail moving
try {
    Start-Sleep -Seconds 2
    $outlook2 = New-Object -ComObject Outlook.Application
    $ns2 = $outlook2.GetNamespace('MAPI')
    $store = $null
    try { $store = $ns2.Folders.Item('david.min@yakjin.com') } catch {}
    if (-not $store) { try { $store = $ns2.Folders | Select-Object -First 1 } catch {} }
    if ($store) {
        $inbox2 = $null
        try { $inbox2 = $store.Folders.Item('Received') } catch {}
        if (-not $inbox2) { try { $inbox2 = $ns2.GetDefaultFolder(6) } catch {} }
        function Find-Folder($folder, $name) {
            if ($folder.Name -eq $name) { return $folder }
            foreach ($sub in $folder.Folders) {
                $result = Find-Folder $sub $name
                if ($result) { return $result }
            }
            return $null
        }
        $targetFolder = Find-Folder $store '브랜드별 마감보고'
        if ($targetFolder) {
            $patterns = @( '*월마감*', '*브랜드*마감*', '*바이어*월마감*'            )
            $moved = 0
            $cutoff = (Get-Date).AddDays(-7)
            $filter = "[ReceivedTime] >= '" + $cutoff.ToString('yyyy-MM-dd') + "'"
            try {
                $filtered = $inbox2.Items.Restrict($filter)
                for ($i = $filtered.Count; $i -ge 1; $i--) {
                    $mail = $filtered.Item($i)
                    if (-not $mail) { continue }
                    $subj = if ($mail.Subject) { $mail.Subject.ToString() } else { '' }
                    foreach ($p in $patterns) {
                        if ($subj -like $p) {
                            try { $mail.Move($targetFolder) | Out-Null; $moved++ } catch { }
                            break
                        }
                    }
                }
            }
            catch {
                $total = $inbox2.Items.Count
                for ($i = $total; $i -ge 1; $i--) {
                    $mail = $inbox2.Items.Item($i)
                    if (-not $mail) { continue }
                    $received = try { $mail.ReceivedTime } catch { $null }
                    if ($received -and $received -lt $cutoff) { continue }
                    $subj = if ($mail.Subject) { $mail.Subject.ToString() } else { '' }
                    foreach ($p in $patterns) {
                        if ($subj -like $p) {
                            try { $mail.Move($targetFolder) | Out-Null; $moved++ } catch { }
                            break
                        }
                    }
                }
            }
            if ($moved -gt 0) {
                Add-Content -Path $LogFile -Value "[$Timestamp] MOVED=$moved"
            }
        }
    }
}
catch {
    Add-Content -Path $LogFile -Value "[$Timestamp] MOVE_ERR: $($_.Exception.Message)"
}
finally {
    if ($outlook2) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook2) | Out-Null }
}
Write-Host "DONE:$Timestamp"