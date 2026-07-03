$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('MAPI')
$store = $ns.Folders.Item('david.min@yakjin.com')
$inbox = $store.Folders.Item('받은 편지함')

function Find-Folder($folder, $name) {
    if ($folder.Name -eq $name) { return $folder }
    foreach ($sub in $folder.Folders) {
        $result = Find-Folder $sub $name
        if ($result) { return $result }
    }
    return $null
}

$targetFolder = Find-Folder $store '브랜드별 마감보고'
if (-not $targetFolder) { Write-Output 'FOLDER NOT FOUND'; exit 1 }

$patterns = @('*월마감*', '*브랜드*마감*', '*바이어*월마감*')
$moved = 0

# Use Restrict to get only last 7 days
$cutoff = (Get-Date).AddDays(-7)
$filter = "[ReceivedTime] >= '" + $cutoff.ToString('yyyy-MM-dd') + "'"

try {
    $filtered = $inbox.Items.Restrict($filter)
    Write-Output ("Recent items: " + $filtered.Count)
    
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
} catch {
    Write-Output ("Restrict failed, scanning all...")
    $total = $inbox.Items.Count
    for ($i = $total; $i -ge 1; $i--) {
        $mail = $inbox.Items.Item($i)
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
Write-Output ("Done. Moved: " + $moved)
