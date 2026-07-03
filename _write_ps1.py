import sys
import os

# Write PowerShell script with proper UTF-8 BOM
script = """$outlook = New-Object -ComObject Outlook.Application
$ns = $outlook.GetNamespace('MAPI')
$store = $ns.Folders.Item('david.min@yakjin.com')

function Find-Folder($folder, $name) {
    if ($folder.Name -eq $name) { return $folder }
    foreach ($sub in $folder.Folders) {
        $result = Find-Folder $sub $name
        if ($result) { return $result }
    }
    return $null
}

$f1 = Find-Folder $store '\ube0c\ub79c\ub4dc\ubcc4 \ub9c8\uac10 \ubcf4\uace0'
$f2 = Find-Folder $store '\ube0c\ub79c\ub4dc\ubcc4 \ub9c8\uac10\ubcf4\uace0'
if ($f1) { Write-Output "FOUND with space: $($f1.FolderPath)" }
if ($f2) { Write-Output "FOUND without space: $($f2.FolderPath)" }
if (-not $f1 -and -not $f2) { Write-Output 'NEITHER exists' }
"""

# Write with UTF-8 BOM (byte order mark 0xEF 0xBB 0xBF)
with open(r'C:\Users\user\.openclaw\workspace\_outlook9.ps1', 'w', encoding='utf-8-sig') as f:
    f.write(script)

print("Written OK")
