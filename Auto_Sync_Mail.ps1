[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Stop"

function Get-UStr($arr) {
    return -join ($arr | ForEach-Object { [char]$_ })
}

$workDir = $PSScriptRoot
if (-not $workDir) {
    $workDir = "D:\Antigravity Project\Outlook圖檢異常mail自動彙整"
}
$keyword = Get-UStr @(0x5716, 0x6AA2, 0x7570, 0x5E38) # 圖檢異常
$folderSuffix = Get-UStr @(0x005F, 0x5BC4, 0x4EF6, 0x5099, 0x4EFD) # _寄件備份
$listSuffix = Get-UStr @(0x005F, 0x5BC4, 0x4EF6, 0x5099, 0x4EFD, 0x6E05, 0x55AE) # _寄件備份清單

$targetDir = Join-Path $workDir ($keyword + $folderSuffix)
$targetFolderName = $keyword + $folderSuffix

$excelPath = Join-Path $workDir ($keyword + $listSuffix + ".xlsx")
$csvPath = Join-Path $workDir ($keyword + $listSuffix + ".csv")
$logPath = Join-Path $workDir "sync_log.txt"

function Write-Log($msg) {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] $msg"
    Write-Host $line
    Add-Content -Path $logPath -Value $line -Encoding UTF8
}

Write-Log "========== Starting Auto Sync Mail Task =========="

# 1. Connect to Outlook MAPI
$outlook = $null
try {
    $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    Write-Log "Connected to running Outlook instance."
} catch {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        Write-Log "Started new Outlook COM instance."
    } catch {
        Write-Log "ERROR: Failed to connect or start Outlook: $_"
        exit 1
    }
}

$namespace = $outlook.GetNamespace("MAPI")
$folderSent = Get-UStr @(0x5BC4, 0x4EF6, 0x5099, 0x4EFD) # 寄件備份

function Find-SentFoldersRecursive($folder) {
    $res = @()
    if ($folder.Name -eq $folderSent) {
        $res += $folder
    }
    foreach ($sub in $folder.Folders) {
        $res += (Find-SentFoldersRecursive $sub)
    }
    return $res
}

# Scan archive.pst (primary default mailbox) and default Sent Mail
$sentFoldersToScan = @()

foreach ($store in $namespace.Stores) {
    $filePath = ""
    try { $filePath = $store.FilePath } catch {}
    $dispName = $store.DisplayName

    # Primary target is archive.pst
    if ($filePath -like "*archive.pst*" -or $dispName -like "*archive*") {
        Write-Log "Found primary archive.pst store: $dispName ($filePath)"
        try {
            $root = $store.GetRootFolder()
            $found = Find-SentFoldersRecursive $root
            foreach ($f in $found) {
                Write-Log "  Added Sent folder: $($f.FullFolderPath) (Items: $($f.Items.Count))"
                $sentFoldersToScan += $f
            }
        } catch {
            Write-Log "  Warning: Unable to scan root of store ${dispName} - $_"
        }
    }
}

# Also ensure default Sent Folder is checked (in case Outlook default MAPI folder points to it)
try {
    $defaultSent = $namespace.GetDefaultFolder(5) # olFolderSentMail
    $alreadyIn = $false
    foreach ($sf in $sentFoldersToScan) {
        if ($sf.EntryID -eq $defaultSent.EntryID) {
            $alreadyIn = $true; break
        }
    }
    if (-not $alreadyIn) {
        Write-Log "Added Default Sent Folder: $($defaultSent.FullFolderPath) (Items: $($defaultSent.Items.Count))"
        $sentFoldersToScan += $defaultSent
    }
} catch {
    Write-Log "Warning: Could not get default Sent folder: $_"
}

if ($sentFoldersToScan.Count -eq 0) {
    Write-Log "ERROR: No sent folders found to scan!"
    exit 1
}

# 2. Read existing keys from Excel
if (-not (Test-Path $excelPath)) {
    Write-Log "ERROR: Target Excel file not found: $excelPath"
    exit 1
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($excelPath)
} catch {
    Write-Log "ERROR: Failed to open workbook (may be in use): $_"
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 1
}

$ws = $wb.Sheets.Item(1)
$existingRows = $ws.UsedRange.Rows.Count
Write-Log "Current existing rows in Excel: $existingRows"

function Get-MailKey($sentOn, $subj) {
    $dt = [DateTime]$sentOn
    $dateKey = $dt.ToString("yyyyMMdd_HHmm")
    return "${dateKey}_${subj}".Trim()
}

$existingDict = @{}
for ($r = 2; $r -le $existingRows; $r++) {
    $sTime = $ws.Cells.Item($r, 1).Text
    $sSubj = $ws.Cells.Item($r, 2).Text
    if ($sTime -and $sSubj) {
        $k = Get-MailKey $sTime $sSubj
        $existingDict[$k] = $true
    }
}
Write-Log "Existing keys loaded in memory: $($existingDict.Count)"

# 3. Helpers
function Clean-FileName ($name) {
    $invalidChars = [IO.Path]::GetInvalidFileNameChars()
    $clean = $name
    foreach ($char in $invalidChars) {
        $clean = $clean.Replace($char, '_')
    }
    $clean = $clean -replace '[:\*\?"<>\|/\\+]', '_'
    $clean = $clean -replace '#', ''
    $clean = $clean -replace '\[', '('
    $clean = $clean -replace '\]', ')'
    if ($clean.Length -gt 85) {
        $clean = $clean.Substring(0, 85)
    }
    return $clean.Trim()
}

function Parse-MailLines($body) {
    $issueLines = [System.Collections.Generic.List[string]]::new()
    $resLines = [System.Collections.Generic.List[string]]::new()
    $beforeLines = [System.Collections.Generic.List[string]]::new()

    if (-not $body) {
        return @{ Issue = ""; Resolution = "" }
    }

    $lines = $body -split "\r?\n"
    $state = "BEFORE"

    $pStop = '(?i)^\s*(?:From:|\u5BC4\u4EF6\u8005:|-----Original)'
    $pSig = '(?i)^\s*(?:Best Regards|Quality Management|\+886-|-{5,}|_{5,})'
    $pIssue = '(?i)^\s*\[?Issue\]?(?::|\uFF1A)?\s*(.*)$'
    $pRes = '(?i)^\s*\[?(?:(?:\u6703\u8B70|\u65E9\u6703)?\u6C7A\u8B70|\u5716\u6AA2\u5224\u5B9A|\u5224\u5B9A)\]?(?::|\uFF1A)?\s*(.*)$'
    $pSkipBefore = '(?i)^\s*(?:Hi\s+All,?|Product:?|LAC\d+.*)\s*$'

    foreach ($rawLine in $lines) {
        $line = $rawLine.Trim()

        if ($line -match $pStop -or $line -match $pSig) {
            break
        }

        if ($line -match $pIssue) {
            $state = "ISSUE"
            $inlineText = $matches[1].Trim()
            if ($inlineText -and -not ($inlineText -match $pRes)) {
                $issueLines.Add($inlineText)
            }
            continue
        }

        if ($line -match $pRes) {
            $state = "RESOLUTION"
            $inlineText = $matches[1].Trim()
            if ($line -match '(?i)^\s*(?:\u5224\u5B9A(?::|\uFF1A).*)') {
                $resLines.Add($line)
            } elseif ($inlineText) {
                if (-not ($inlineText -match $pRes)) {
                    $resLines.Add($inlineText)
                }
            }
            continue
        }

        if ($state -eq "BEFORE") {
            if ($line -and -not ($line -match $pSkipBefore)) {
                $beforeLines.Add($line)
            }
        } elseif ($state -eq "ISSUE") {
            if ($line) {
                $issueLines.Add($line)
            }
        } elseif ($state -eq "RESOLUTION") {
            if ($line) {
                $resLines.Add($line)
            }
        }
    }

    $finalIssue = ""
    if ($issueLines.Count -gt 0) {
        $finalIssue = ($issueLines -join "`r`n").Trim()
    } elseif ($beforeLines.Count -gt 0) {
        $finalIssue = ($beforeLines -join "`r`n").Trim()
    }

    $finalRes = ($resLines -join "`r`n").Trim()

    return @{
        Issue = $finalIssue
        Resolution = $finalRes
    }
}

# 4. Scan for new emails (only 2026 emails)
$newEmails = @()

foreach ($sf in $sentFoldersToScan) {
    try {
        $items = $sf.Items
        $items.Sort("[SentOn]", $false)

        foreach ($item in $items) {
            if ($item.MessageClass -eq "IPM.Note" -or $item.MessageClass -like "IPM.Note*") {
                $subj = ""
                try { $subj = $item.Subject } catch {}
                if ($subj -and $subj.Contains($keyword)) {
                    $sentOn = $item.SentOn
                    # Filter only current year (2026)
                    if ($sentOn.Year -eq 2026) {
                        $k = Get-MailKey $sentOn $subj

                        if (-not $existingDict.ContainsKey($k)) {
                            $sentOnStr = $sentOn.ToString("yyyy-MM-dd HH:mm:ss")
                            $datePrefix = $sentOn.ToString("yyyyMMdd_HHmmss")
                            $cleanSubj = Clean-FileName $subj
                            $fileName = "${datePrefix}_${cleanSubj}.msg"
                            $destFile = Join-Path $targetDir $fileName

                            if (-not (Test-Path $destFile)) {
                                try {
                                    $item.SaveAs($destFile, 9)
                                } catch {
                                    $item.SaveAs($destFile, 3)
                                }
                            }

                            $parsed = Parse-MailLines $item.Body

                            $newEmails += [PSCustomObject]@{
                                SentTime   = $sentOnStr
                                Subject    = $subj
                                Issue      = $parsed.Issue
                                Resolution = $parsed.Resolution
                                FileName   = $fileName
                            }

                            $existingDict[$k] = $true
                            Write-Log "Found NEW email: $sentOnStr | $subj"
                        }
                    }
                }
            }
        }
    } catch {
        Write-Log "Warning scanning folder: $_"
    }
}

Write-Log "Total NEW emails found to append: $($newEmails.Count)"

if ($newEmails.Count -eq 0) {
    Write-Log "No new emails found. Sync completed."
    $wb.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    exit 0
}

# 5. Append to Excel
$tipText = Get-UStr @(0x9EDE, 0x64CA, 0x958B, 0x555F, 0x90F5, 0x4EF6) # 點擊開啟郵件
$displayText = Get-UStr @(0x958B, 0x555F, 0x90F5, 0x4EF6)             # 開啟郵件

$currentRow = $existingRows + 1

foreach ($m in $newEmails) {
    # Col 1: 寄件時間
    $c1 = $ws.Cells.Item($currentRow, 1)
    $c1.Value2 = $m.SentTime
    $c1.HorizontalAlignment = -4108
    $c1.VerticalAlignment = -4160

    # Col 2: 主旨
    $c2 = $ws.Cells.Item($currentRow, 2)
    $c2.Value2 = $m.Subject
    $c2.VerticalAlignment = -4160
    $c2.WrapText = $true

    # Col 3: [Issue]
    $c3 = $ws.Cells.Item($currentRow, 3)
    $c3.Value2 = $m.Issue
    $c3.VerticalAlignment = -4160
    $c3.WrapText = $true

    # Col 4: [會議決議]
    $c4 = $ws.Cells.Item($currentRow, 4)
    $c4.Value2 = $m.Resolution
    $c4.VerticalAlignment = -4160
    $c4.WrapText = $true

    # Col 5: 開啟郵件
    $c5 = $ws.Cells.Item($currentRow, 5)
    $relLink = "${targetFolderName}/$($m.FileName)"
    $ws.Hyperlinks.Add($c5, $relLink, [Type]::Missing, $tipText, $displayText) | Out-Null
    $c5.HorizontalAlignment = -4108
    $c5.VerticalAlignment = -4160

    $currentRow++
}

$totalRows = $currentRow - 1
Write-Log "Total rows after append: $totalRows. Sorting..."

# Sort A2:E$totalRows ascending by Column A
$dataRange = $ws.Range("A2", "E$totalRows")
$keyCell = $ws.Range("A2")
$dataRange.Sort($keyCell, 1)

# Borders
$fullRange = $ws.Range("A1", "E$totalRows")
$fullRange.Borders.LineStyle = 1
$fullRange.Borders.Color = 13421772

# Column widths
$ws.Columns.Item(1).ColumnWidth = 20
$ws.Columns.Item(2).ColumnWidth = 40
$ws.Columns.Item(3).ColumnWidth = 45
$ws.Columns.Item(4).ColumnWidth = 55
$ws.Columns.Item(5).ColumnWidth = 14

$wb.Save()
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Log "Excel successfully saved."

# 6. Update CSV
$csvLines = @()
$csvHeaders = @("寄件時間", "主旨", "[Issue]", "[會議決議]", "檔案連結")
$csvLines += ($csvHeaders -join ",")

$excel2 = New-Object -ComObject Excel.Application
$excel2.Visible = $false
try {
    $wb2 = $excel2.Workbooks.Open($excelPath)
    $ws2 = $wb2.Sheets.Item(1)
    for ($r = 2; $r -le $totalRows; $r++) {
        $t1 = $ws2.Cells.Item($r, 1).Text
        $t2 = $ws2.Cells.Item($r, 2).Text
        $t3 = $ws2.Cells.Item($r, 3).Text
        $t4 = $ws2.Cells.Item($r, 4).Text
        $t5 = ""
        try { $t5 = $ws2.Cells.Item($r, 5).Hyperlinks.Item(1).Address } catch {}

        $csvLines += ('"' + ($t1 -replace '"', '""') + '",' +
                      '"' + ($t2 -replace '"', '""') + '",' +
                      '"' + ($t3 -replace '"', '""') + '",' +
                      '"' + ($t4 -replace '"', '""') + '",' +
                      '"' + ($t5 -replace '"', '""') + '"')
    }
    $wb2.Close($false)
} finally {
    $excel2.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel2) | Out-Null
}

[System.IO.File]::WriteAllLines($csvPath, $csvLines, [System.Text.Encoding]::UTF8)
Write-Log "CSV successfully updated."

# 7. Notification
try {
    $wshell = New-Object -ComObject Wscript.Shell
    $popTitle = Get-UStr @(0x004F, 0x0075, 0x0074, 0x006C, 0x006F, 0x006F, 0x006B, 0x0020, 0x5716, 0x6AA2, 0x7570, 0x5E38, 0x90F5, 0x4EF6, 0x540C, 0x6B65, 0x901A, 0x77E5)
    $popBody = (Get-UStr @(0x5DF2, 0x5B8C, 0x6210, 0x5716, 0x6AA2, 0x7570, 0x5E38, 0x90F5, 0x4EF6, 0x540C, 0x6B65, 0xFF01, 0x000A, 0x65B0, 0x589E, 0x7B46, 0x6578, 0xFF1A)) + "$($newEmails.Count)"
    $wshell.Popup($popBody, 5, $popTitle, 64) | Out-Null
} catch {}

Write-Log "========== Sync Completed Successfully =========="
