# Python script to generate PowerShell script with proper UTF-8 BOM
ps1_script = """$outlook = New-Object -ComObject Outlook.Application
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

$target = Find-Folder $store '브랜드별 마감보고'
Write-Output "Target: $($target.FolderPath)"

$keywords = @('월마감')
$count = 0
$total = $inbox.Items.Count
Write-Output "Total inbox items: $total"

for ($i = 1; $i -le $total; $i++) {
    $mail = $inbox.Items.Item($i)
    if (-not $mail) { continue }
    $subj = if ($mail.Subject) { $mail.Subject.ToString() } else { '' }
    
    $matched = $false
    foreach ($kw in $keywords) {
        if ($subj -like "*$kw*") {
            $matched = $true
            break
        }
    }
    if ($matched) {
        $count++
        Write-Output ("[" + $mail.ReceivedTime.ToString('yyyy-MM-dd') + "] " + $subj)
    }
}
Write-Output "`nMatching: $count"
"""

with open(r'C:\Users\user\.openclaw\workspace\_outlook_search.ps1', 'w', encoding='utf-8-sig') as f:
    f.write(ps1_script)

print("Written OK")
