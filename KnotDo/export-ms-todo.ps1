# ==============================================================
# KnotDo - Microsoft To-Do Graph Export
# ==============================================================
# Exports all your To-Do lists and tasks to a JSON file
# using the Microsoft Graph PowerShell SDK.
#
# No Azure app registration required.
# Uses Microsoft's own first-party Graph Command Line Tools client.
#
# USAGE:
#   1. Open PowerShell (Windows Terminal or Start Menu)
#   2. Run:  .\scripts\export-ms-todo.ps1
#   3. You will get a code to enter at microsoft.com/devicelogin
#      Sign in with the Microsoft account that has your To-Do tasks
#   4. The export saves to your Downloads folder as ms-todo-export.json
#   5. Upload that file to KnotDo > Import > Microsoft Graph Export
# ==============================================================

$ErrorActionPreference = "Stop"

# -- 1. Install Microsoft.Graph.Tasks module if needed -------------------------
Write-Host ""
Write-Host "Checking for Microsoft.Graph.Tasks module..." -ForegroundColor Cyan

if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph.Tasks" -ErrorAction SilentlyContinue)) {
    Write-Host "Installing Microsoft.Graph.Tasks (this may take a minute)..." -ForegroundColor Yellow
    Install-Module "Microsoft.Graph.Tasks" -Scope CurrentUser -Force -AllowClobber
    Write-Host "Installed." -ForegroundColor Green
} else {
    Write-Host "Module already installed." -ForegroundColor Green
}

Import-Module "Microsoft.Graph.Tasks" -ErrorAction Stop

# -- 2. Connect to Microsoft Graph ---------------------------------------------
Write-Host ""
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Write-Host "You will receive a code to enter at: https://microsoft.com/devicelogin" -ForegroundColor DarkGray
Write-Host "Sign in with the Microsoft account that contains your To-Do tasks." -ForegroundColor DarkGray
Write-Host ""

Connect-MgGraph -Scopes "Tasks.Read" -UseDeviceAuthentication -NoWelcome

Write-Host ""
Write-Host "Connected!" -ForegroundColor Green

# -- 3. Fetch all To-Do lists --------------------------------------------------
Write-Host ""
Write-Host "Fetching task lists..." -ForegroundColor Cyan

$mgLists = Get-MgUserTodoList -UserId "me" -All
Write-Host "Found $($mgLists.Count) list(s)." -ForegroundColor Green

# -- 4. Fetch tasks for each list ----------------------------------------------
$export = @{
    exportedAt = (Get-Date -Format "o")
    source     = "microsoft-graph"
    lists      = [System.Collections.Generic.List[object]]::new()
}

$totalTasks = 0

foreach ($list in $mgLists) {
    Write-Host "  Fetching: $($list.DisplayName)..." -ForegroundColor DarkCyan -NoNewline

    $mgTasks = Get-MgUserTodoListTask -TodoTaskListId $list.Id -UserId "me" -All

    $taskArr = [System.Collections.Generic.List[object]]::new()
    foreach ($t in $mgTasks) {
        $taskArr.Add(@{
            id                   = $t.Id
            title                = $t.Title
            status               = $t.Status
            importance           = $t.Importance
            dueDateTime          = if ($t.DueDateTime)       { @{ dateTime = $t.DueDateTime.DateTime;       timeZone = $t.DueDateTime.TimeZone       } } else { $null }
            completedDateTime    = if ($t.CompletedDateTime) { @{ dateTime = $t.CompletedDateTime.DateTime; timeZone = $t.CompletedDateTime.TimeZone } } else { $null }
            body                 = if ($t.Body)              { @{ content = $t.Body.Content; contentType = $t.Body.ContentType } } else { $null }
            createdDateTime      = $t.CreatedDateTime
            lastModifiedDateTime = $t.LastModifiedDateTime
        })
    }

    $export.lists.Add(@{
        id          = $list.Id
        displayName = $list.DisplayName
        isOwner     = $list.IsOwner
        tasks       = $taskArr
    })

    $totalTasks += $mgTasks.Count
    Write-Host " $($mgTasks.Count) tasks" -ForegroundColor Gray
}

# -- 5. Save to Downloads ------------------------------------------------------
$outputPath = Join-Path $env:USERPROFILE "Downloads\ms-todo-export.json"

Write-Host ""
Write-Host "Saving export..." -ForegroundColor Cyan
$export | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath $outputPath -Encoding UTF8

$fileSizeMB = [math]::Round((Get-Item $outputPath).Length / 1MB, 1)

Write-Host ""
Write-Host "==============================================================" -ForegroundColor Green
Write-Host "Export complete!" -ForegroundColor Green
Write-Host "  Lists : $($mgLists.Count)" -ForegroundColor White
Write-Host "  Tasks : $($totalTasks.ToString('N0'))" -ForegroundColor White
Write-Host "  Size  : $fileSizeMB MB" -ForegroundColor White
Write-Host "  Saved : $outputPath" -ForegroundColor White
Write-Host "==============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next: upload ms-todo-export.json to KnotDo > Import > Microsoft Graph Export" -ForegroundColor Cyan
Write-Host ""

Disconnect-MgGraph | Out-Null
