# ==============================================================
# KnotDo - Microsoft To-Do Graph Export
# ==============================================================
# Exports all your To-Do lists and tasks to a JSON file.
#
# No Azure app registration required.
# No module installation required - uses raw Graph API REST calls.
# Works on Windows PowerShell 5.1+ out of the box.
#
# USAGE:
#   1. Open PowerShell (Windows Terminal or Start Menu)
#   2. Run:  .\export-ms-todo.ps1
#   3. Open the URL shown, enter the code, sign in with the
#      Microsoft account that has your To-Do tasks
#   4. Export saves to Downloads\ms-todo-export.json
#   5. Upload to KnotDo > Import > Microsoft Graph Export
# ==============================================================

$ErrorActionPreference = "Stop"

# Microsoft Graph Command Line Tools (first-party Microsoft app - no registration needed)
$ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
$Scope    = "Tasks.Read offline_access"
$Tenant   = "common"

# -- 1. Request device code ----------------------------------------------------
Write-Host ""
Write-Host "Requesting sign-in code from Microsoft..." -ForegroundColor Cyan

$deviceCodeBody = @{
    client_id = $ClientId
    scope     = $Scope
}

$deviceCode = Invoke-RestMethod `
    -Method POST `
    -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/devicecode" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body $deviceCodeBody

Write-Host ""
Write-Host "================================================================" -ForegroundColor Yellow
Write-Host $deviceCode.message -ForegroundColor Yellow
Write-Host "================================================================" -ForegroundColor Yellow
Write-Host ""

# -- 2. Poll for token ---------------------------------------------------------
Write-Host "Waiting for you to sign in..." -ForegroundColor DarkGray

$tokenBody = @{
    grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
    client_id   = $ClientId
    device_code = $deviceCode.device_code
}

$token       = $null
$interval    = [int]$deviceCode.interval
$expiresSecs = [int]$deviceCode.expires_in
$waited      = 0

while ($waited -lt $expiresSecs) {
    Start-Sleep -Seconds $interval
    $waited += $interval

    try {
        $token = Invoke-RestMethod `
            -Method POST `
            -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" `
            -ContentType "application/x-www-form-urlencoded" `
            -Body $tokenBody
        break
    } catch {
        $raw = $_.ErrorDetails.Message
        if ($raw) {
            try {
                $errObj = $raw | ConvertFrom-Json
                if ($errObj.error -eq "authorization_pending") { continue }
                if ($errObj.error -eq "authorization_declined") { throw "Sign-in was declined. Run the script again." }
                if ($errObj.error -eq "expired_token")          { throw "The code expired. Run the script again." }
            } catch [System.Management.Automation.RuntimeException] { throw }
            catch { }
        }
        throw
    }
}

if (-not $token) {
    Write-Host "ERROR: Sign-in timed out. Run the script again." -ForegroundColor Red
    exit 1
}

Write-Host "Signed in!" -ForegroundColor Green
Write-Host ""

$headers = @{ Authorization = "Bearer $($token.access_token)" }

# -- 3. Helper: paginated Graph GET --------------------------------------------
function Get-GraphAll($url) {
    $all     = [System.Collections.Generic.List[object]]::new()
    $nextUrl = $url
    while ($nextUrl) {
        $resp    = Invoke-RestMethod -Uri $nextUrl -Headers $headers
        if ($resp.value) { $all.AddRange($resp.value) }
        $nextUrl = $resp.'@odata.nextLink'
    }
    return $all
}

# -- 4. Fetch all To-Do lists --------------------------------------------------
Write-Host "Fetching task lists..." -ForegroundColor Cyan
$lists = Get-GraphAll "https://graph.microsoft.com/v1.0/me/todo/lists"
Write-Host "Found $($lists.Count) list(s)." -ForegroundColor Green
Write-Host ""

# -- 5. Fetch tasks for each list ----------------------------------------------
$export = [ordered]@{
    exportedAt = (Get-Date -Format "o")
    source     = "microsoft-graph"
    lists      = [System.Collections.Generic.List[object]]::new()
}

$totalTasks = 0

foreach ($list in $lists) {
    Write-Host "  Fetching: $($list.displayName)..." -ForegroundColor DarkCyan -NoNewline

    $tasks = Get-GraphAll "https://graph.microsoft.com/v1.0/me/todo/lists/$($list.id)/tasks"

    $taskArr = [System.Collections.Generic.List[object]]::new()
    foreach ($t in $tasks) {
        $taskArr.Add([ordered]@{
            id                   = $t.id
            title                = $t.title
            status               = $t.status
            importance           = $t.importance
            dueDateTime          = $t.dueDateTime
            completedDateTime    = $t.completedDateTime
            body                 = $t.body
            createdDateTime      = $t.createdDateTime
            lastModifiedDateTime = $t.lastModifiedDateTime
        })
    }

    $export.lists.Add([ordered]@{
        id          = $list.id
        displayName = $list.displayName
        isOwner     = $list.isOwner
        tasks       = $taskArr
    })

    $totalTasks += $tasks.Count
    Write-Host " $($tasks.Count) tasks" -ForegroundColor Gray
}

# -- 6. Save to Downloads ------------------------------------------------------
$outputPath = Join-Path $env:USERPROFILE "Downloads\ms-todo-export.json"

Write-Host ""
Write-Host "Saving..." -ForegroundColor Cyan
$export | ConvertTo-Json -Depth 10 -Compress | Out-File -FilePath $outputPath -Encoding UTF8

$fileSizeMB = [math]::Round((Get-Item $outputPath).Length / 1MB, 2)

Write-Host ""
Write-Host "==============================================================" -ForegroundColor Green
Write-Host "Export complete!" -ForegroundColor Green
Write-Host "  Lists : $($lists.Count)" -ForegroundColor White
Write-Host "  Tasks : $($totalTasks.ToString('N0'))" -ForegroundColor White
Write-Host "  Size  : $fileSizeMB MB" -ForegroundColor White
Write-Host "  Saved : $outputPath" -ForegroundColor White
Write-Host "==============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Next: upload ms-todo-export.json to KnotDo > Import > Microsoft Graph Export" -ForegroundColor Cyan
Write-Host ""
