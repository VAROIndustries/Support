# ==============================================================
# KnotDo - Microsoft To-Do Local Export
# ==============================================================
# Reads directly from the Microsoft To-Do app's local database.
# No Microsoft account sign-in required.
# No Azure app registration required.
# Works completely offline.
#
# REQUIREMENTS:
#   - Microsoft To-Do app installed (from Microsoft Store)
#   - The app must have synced at least once so data is local
#
# USAGE:
#   1. Open PowerShell
#   2. Run:  .\export-ms-todo.ps1
#   3. Upload Downloads\ms-todo-export.json to KnotDo > Import
# ==============================================================

$ErrorActionPreference = "Stop"

# -- 1. Find the To-Do local database ------------------------------------------
Write-Host ""
Write-Host "Looking for Microsoft To-Do database..." -ForegroundColor Cyan

$accountsRoot = Join-Path $env:LOCALAPPDATA "Packages\Microsoft.Todos_8wekyb3d8bbwe\LocalState\AccountsRoot"

if (-not (Test-Path $accountsRoot)) {
    Write-Host "ERROR: Microsoft To-Do app not found." -ForegroundColor Red
    Write-Host "Install it from the Microsoft Store, open it, and let it sync first." -ForegroundColor Yellow
    exit 1
}

$dbPath = Get-ChildItem $accountsRoot -Recurse -Filter "todosqlite.db" | Select-Object -First 1 -ExpandProperty FullName

if (-not $dbPath) {
    Write-Host "ERROR: To-Do database file not found. Open the To-Do app and wait for it to sync." -ForegroundColor Red
    exit 1
}

$dbSizeMB = [math]::Round((Get-Item $dbPath).Length / 1MB, 1)
Write-Host "Found database ($dbSizeMB MB)" -ForegroundColor Green

# -- 2. Download sqlite3.exe if not already present ----------------------------
$sqlitePath = Join-Path $env:TEMP "knotdo_sqlite3.exe"

if (-not (Test-Path $sqlitePath)) {
    Write-Host "Downloading sqlite3 (one-time, ~2MB)..." -ForegroundColor Cyan

    # Fetch the sqlite.org download page to get the current version URL
    $dlPage  = (Invoke-WebRequest -Uri "https://www.sqlite.org/download.html" -UseBasicParsing).Content
    $match   = [regex]::Match($dlPage, '(20\d\d/sqlite-tools-win-x64-\d+\.zip)')

    if (-not $match.Success) {
        Write-Host "ERROR: Could not find sqlite3 download URL from sqlite.org." -ForegroundColor Red
        exit 1
    }

    $zipUrl  = "https://www.sqlite.org/" + $match.Value
    $zipPath = Join-Path $env:TEMP "knotdo_sqlite_tools.zip"

    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip   = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
    $entry = $zip.Entries | Where-Object { $_.Name -eq "sqlite3.exe" } | Select-Object -First 1
    if (-not $entry) {
        $zip.Dispose()
        Write-Host "ERROR: sqlite3.exe not found inside the downloaded zip." -ForegroundColor Red
        exit 1
    }
    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $sqlitePath, $true)
    $zip.Dispose()
    Remove-Item $zipPath -Force

    Write-Host "sqlite3 ready." -ForegroundColor Green
}

# -- 3. Copy DB to temp (never lock the live file) -----------------------------
$tempDb = Join-Path $env:TEMP "knotdo_todo_export.db"
Copy-Item $dbPath $tempDb -Force

# -- 4. Read all lists ---------------------------------------------------------
Write-Host "Reading lists..." -ForegroundColor Cyan

$listSql = @"
.mode json
SELECT local_id, name, folder_type FROM task_folders WHERE deleted=0;
"@

$listsRaw = $listSql | & $sqlitePath $tempDb
$lists    = @($listsRaw | ConvertFrom-Json) | Where-Object { $_.name -ne "Flagged Email" -and $_.name -ne "Flagged Emails" }

Write-Host "Found $($lists.Count) list(s)." -ForegroundColor Green

# -- 5. Read all tasks via CSV (most reliable cross-platform approach) ----------
Write-Host "Reading tasks..." -ForegroundColor Cyan

$tempTasksCsv   = Join-Path $env:TEMP "knotdo_tasks.csv"
$tempScriptFile = Join-Path $env:TEMP "knotdo_query.sql"

# Write sqlite3 script to a file - CSV avoids JSON parsing issues
$csvPath = $tempTasksCsv.Replace('\', '/')
Set-Content -Path $tempScriptFile -Encoding ASCII -Value ".separator `",`"
.headers on
.output $csvPath
SELECT t.local_id, t.task_folder_local_id, f.name as list_name, t.subject, t.status, t.importance, t.due_date, t.completed_datetime, t.body_content FROM tasks t LEFT JOIN task_folders f ON t.task_folder_local_id = f.local_id WHERE t.deleted=0;
.quit"

# Run sqlite3 with stdin redirected from the script file (avoids pipeline size limits)
$proc = Start-Process -FilePath $sqlitePath -ArgumentList $tempDb `
    -RedirectStandardInput $tempScriptFile `
    -RedirectStandardError "$env:TEMP\knotdo_err.txt" `
    -Wait -NoNewWindow -PassThru

Remove-Item $tempScriptFile -Force -ErrorAction SilentlyContinue

$allTasks = @()
if (Test-Path $tempTasksCsv) {
    $allTasks = @(Import-Csv $tempTasksCsv)
    Remove-Item $tempTasksCsv -Force -ErrorAction SilentlyContinue
}
Remove-Item "$env:TEMP\knotdo_err.txt" -Force -ErrorAction SilentlyContinue

Write-Host "Found $($allTasks.Count) task(s) total." -ForegroundColor Green

# Tasks already have list_name from JOIN - no grouping needed, just index by local_id
$tasksByList = @{}
foreach ($t in $allTasks) {
    $fid = $t.task_folder_local_id
    if (-not $fid) { continue }
    if (-not $tasksByList.ContainsKey($fid)) { $tasksByList[$fid] = [System.Collections.ArrayList]::new() }
    [void]$tasksByList[$fid].Add($t)
}

# -- 6. Build export -----------------------------------------------------------
$export = [ordered]@{
    exportedAt = (Get-Date -Format "o")
    source     = "microsoft-graph"
    lists      = [System.Collections.ArrayList]::new()
}

$totalTasks = 0

foreach ($list in $lists) {
    $tasks   = if ($tasksByList -and $tasksByList.ContainsKey($list.local_id)) { @($tasksByList[$list.local_id]) } else { @() }
    $taskArr = [System.Collections.ArrayList]::new()

    foreach ($t in $tasks) {
        $status = switch ($t.status) {
            "NotStarted" { "notStarted" }
            "InProgress" { "inProgress" }
            "Completed"  { "completed"  }
            default      { "notStarted" }
        }
        $importance = if ([int]$t.importance -ge 1) { "high" } else { "normal" }

        [void]$taskArr.Add([ordered]@{
            id                = $t.local_id
            title             = $t.subject
            status            = $status
            importance        = $importance
            dueDateTime       = if ($t.due_date)            { [ordered]@{ dateTime = $t.due_date;            timeZone = "UTC" } } else { $null }
            completedDateTime = if ($t.completed_datetime)  { [ordered]@{ dateTime = $t.completed_datetime;  timeZone = "UTC" } } else { $null }
            body              = if ($t.body_content)        { [ordered]@{ content = $t.body_content; contentType = "text" } } else { $null }
        })
    }

    [void]$export.lists.Add([ordered]@{
        id          = $list.local_id
        displayName = $list.name
        isOwner     = $true
        tasks       = $taskArr
    })

    $totalTasks += $tasks.Count
    Write-Host "  $($list.name): $($tasks.Count) tasks" -ForegroundColor Gray
}

# -- 7. Clean up and save ------------------------------------------------------
Remove-Item $tempDb -Force -ErrorAction SilentlyContinue

$outputPath = Join-Path $env:USERPROFILE "Downloads\ms-todo-export.json"
$export | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath $outputPath -Encoding UTF8

$outSizeMB = [math]::Round((Get-Item $outputPath).Length / 1MB, 2)

Write-Host ""
Write-Host "==============================================================" -ForegroundColor Green
Write-Host "Export complete!" -ForegroundColor Green
Write-Host "  Lists : $($lists.Count)" -ForegroundColor White
Write-Host "  Tasks : $($totalTasks.ToString('N0'))" -ForegroundColor White
Write-Host "  Size  : $outSizeMB MB" -ForegroundColor White
Write-Host "  Saved : $outputPath" -ForegroundColor White
Write-Host "==============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next: upload ms-todo-export.json to KnotDo > Import > Microsoft Graph Export" -ForegroundColor Cyan
Write-Host ""
