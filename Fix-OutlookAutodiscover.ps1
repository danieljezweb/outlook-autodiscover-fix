<#
.SYNOPSIS
    Fixes new Outlook autodiscover issues after migrating an email account to Axigen (ax.email).

.DESCRIPTION
    New Outlook for Windows does not query the mail server directly for autodiscover -
    it goes through a Microsoft cloud endpoint (prod-autodetect.outlookmobile.com) which
    caches domain configuration. After a domain is migrated to a new mail server (e.g.
    Rackspace -> Axigen), that cache can be stale, causing new Outlook to fail to add
    the account, loop on sign-in, or ask for an "app password".

    Run with no parameters to get an interactive menu covering all the tool's options.
    Parameters can also be passed directly for scripted/RMM use, which skips the menu.

.PARAMETER EmailAddress
    The email address to refresh. If omitted in direct-parameter mode, the script
    will prompt for it.

.PARAMETER TestMode
    Safe validation mode for use on a test VM or before a client rollout.
    - Checks that autodiscover DNS records resolve correctly for the domain
    - Sends the Microsoft cache-refresh request (non-destructive, safe even on
      production domains)
    - Never touches Outlook (no process kill, no reset) - just reports what it finds
    - Does not require a real mailbox to exist

.PARAMETER WhatIf
    Shows what the script WOULD do at the Outlook-reset step, without actually
    closing Outlook or resetting it. The DNS check and cache-refresh request still
    run for real, since those are non-destructive.

.PARAMETER Menu
    Forces the interactive menu to show even if other parameters were passed.

.EXAMPLE
    .\Fix-OutlookAutodiscover.ps1
    Shows the interactive main menu.

.EXAMPLE
    .\Fix-OutlookAutodiscover.ps1 -TestMode -EmailAddress test@yourtestdomain.com.au
    Skips the menu, runs test mode directly. Useful for scripting.

.EXAMPLE
    .\Fix-OutlookAutodiscover.ps1 -EmailAddress name@client-domain.com.au
    Skips the menu, refreshes the autodiscover cache only (Outlook reset is a
    separate action - use -Menu to access it interactively).
#>

param(
    [string]$EmailAddress,
    [switch]$TestMode,
    [switch]$WhatIf,
    [switch]$Menu
)

# ===================================================================
# Shared functions
# ===================================================================

function Show-Header {
    param([string]$Subtitle)
    Clear-Host
    Write-Host "===================================================" -ForegroundColor Cyan
    Write-Host " Outlook Autodiscover Fix Tool (Jezweb)" -ForegroundColor Cyan
    if ($Subtitle) { Write-Host " $Subtitle" -ForegroundColor Magenta }
    Write-Host "===================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Read-EmailOrDomain {
    param(
        [string]$Prompt,
        [switch]$AllowBareDomain
    )
    $value = Read-Host $Prompt

    if ($AllowBareDomain -and ($value -notmatch "@")) {
        $value = "autodiscover-test@$value"
    }

    if ([string]::IsNullOrWhiteSpace($value) -or ($value -notmatch "^[^@\s]+@[^@\s]+\.[^@\s]+$")) {
        Write-Host "That doesn't look like a valid email address (or domain). Please try again." -ForegroundColor Red
        return $null
    }
    return $value
}

function Test-AutodiscoverDns {
    param([string]$Domain)

    Write-Host "--- DNS check: autodiscover.$Domain ---" -ForegroundColor Cyan
    try {
        $cname = Resolve-DnsName -Name "autodiscover.$Domain" -Type CNAME -ErrorAction Stop
        Write-Host "CNAME found:" -ForegroundColor Green
        $cname | Select-Object Name, NameHost | Format-Table -AutoSize | Out-String | Write-Host
    }
    catch {
        Write-Host "No CNAME record found for autodiscover.$Domain (or it's an A record instead)." -ForegroundColor Yellow
        try {
            $arec = Resolve-DnsName -Name "autodiscover.$Domain" -Type A -ErrorAction Stop
            Write-Host "A record found:" -ForegroundColor Green
            $arec | Select-Object Name, IPAddress | Format-Table -AutoSize | Out-String | Write-Host
        }
        catch {
            Write-Host "No autodiscover record resolves at all for $Domain - this needs fixing in DNS before Outlook autodiscover can work." -ForegroundColor Red
        }
    }

    Write-Host "--- DNS check: SRV records ---" -ForegroundColor Cyan
    foreach ($srv in @("_imaps._tcp.$Domain", "_submission._tcp.$Domain", "_autodiscover._tcp.$Domain")) {
        try {
            $srvRec = Resolve-DnsName -Name $srv -Type SRV -ErrorAction Stop
            Write-Host "$srv -> found" -ForegroundColor Green
            $srvRec | Select-Object Name, NameTarget, Port, Priority | Format-Table -AutoSize | Out-String | Write-Host
        }
        catch {
            Write-Host "$srv -> not found (may be fine depending on your Axigen autodiscover setup)" -ForegroundColor DarkYellow
        }
    }
    Write-Host ""
}

function Invoke-AutodiscoverCacheRefresh {
    param([string]$EmailAddress)

    Write-Host "Requesting Microsoft to refresh its cached autodiscover data for:" -ForegroundColor Yellow
    Write-Host "  $EmailAddress" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "(This only tells Microsoft to re-check the domain's settings -" -ForegroundColor DarkGray
    Write-Host " it does not delete or clear any mailbox/account data.)" -ForegroundColor DarkGray
    Write-Host ""

    # Microsoft has changed this endpoint's hostname before without notice - the original
    # prod-autodetect.outlookmobile.com host now returns NXDOMAIN. Confirmed working
    # endpoint (as of testing 2026-08-19) is listed first, with older/alternate hostnames
    # kept as fallback in case Microsoft moves it again.
    $candidateHosts = @(
        "prod.autodetect.outlook.cloud.microsoft",
        "prod-autodetect.outlookmobile.com",
        "autodetect.outlookmobile.com"
    )

    $succeeded = $false

    foreach ($hostName in $candidateHosts) {
        Write-Host "Trying endpoint: $hostName ..." -ForegroundColor DarkGray
        try {
            Resolve-DnsName -Name $hostName -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Host "  DNS lookup failed for $hostName (name does not exist) - trying next candidate." -ForegroundColor DarkYellow
            continue
        }

        try {
            $uri = "https://$hostName/detect?services=office365,outlook,google,icloud,yahoo&protocols=rest-cloud,rest-outlook,rest-office365,eas,imap,smtp"
            $response = Invoke-WebRequest -Uri $uri -Headers @{ "x-email" = $EmailAddress } -UseBasicParsing -ErrorAction Stop
            Write-Host "  Success via $hostName (HTTP $($response.StatusCode))." -ForegroundColor Green
            Write-Host ""
            Write-Host "Response headers:" -ForegroundColor DarkGray
            $response.Headers | Format-Table -AutoSize | Out-String | Write-Host
            $succeeded = $true
            break
        }
        catch {
            Write-Host "  Request to $hostName failed: $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }

    Write-Host ""
    if ($succeeded) {
        Write-Host "Cache refresh request complete." -ForegroundColor Green
    }
    else {
        Write-Host "None of the known Microsoft autodetect endpoints responded." -ForegroundColor Red
        Write-Host "This can happen if Microsoft has changed the endpoint again. Try the" -ForegroundColor Yellow
        Write-Host "manual tool instead: https://aka.ms/autodetect" -ForegroundColor Yellow
        Write-Host "You can still continue and try adding the account again - this step is often silent even when it works." -ForegroundColor Yellow
    }
    Write-Host ""
}

function Backup-OutlookLocalData {
    <#
        New Outlook doesn't use a PST/OST file the way classic Outlook does - all mail
        lives server-side (synced via IMAP/Graph). The closest thing to a local "data
        file" is the Olk cache folder under LocalAppData, which can hold local drafts,
        rules, and cached items. This copies that folder (and the classic signatures
        folder, in case it's in use) to a timestamped backup location before any reset.

        This is a safety net, not a full mail backup - actual mailbox content is on
        the mail server and is unaffected by any of this.
    #>
    param(
        [string]$DestinationRoot = (Join-Path $env:USERPROFILE "Desktop")
    )

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupFolder = Join-Path $DestinationRoot "OutlookBackup_$timestamp"

    Write-Host "Backing up local Outlook data before making changes..." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Note: your actual mailbox (emails, contacts, calendar) lives on the" -ForegroundColor DarkGray
    Write-Host "mail server and is not affected by a reset. This backs up local" -ForegroundColor DarkGray
    Write-Host "cache, drafts, and settings only, as an extra safety net." -ForegroundColor DarkGray
    Write-Host ""

    try {
        New-Item -Path $backupFolder -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "Could not create backup folder at $backupFolder ($($_.Exception.Message))." -ForegroundColor Red
        return $null
    }

    $olkPath = Join-Path $env:LOCALAPPDATA "Microsoft\Olk"
    $signaturesPath = Join-Path $env:APPDATA "Microsoft\Signatures"

    $copied = @()

    if (Test-Path $olkPath) {
        try {
            Write-Host "Copying new Outlook local data (Olk folder)..." -ForegroundColor Yellow
            Copy-Item -Path $olkPath -Destination (Join-Path $backupFolder "Olk") -Recurse -Force -ErrorAction Stop
            $copied += "Olk (new Outlook local cache/drafts)"
            Write-Host "  Done." -ForegroundColor Green
        }
        catch {
            Write-Host "  Could not copy Olk folder: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "  This can happen if Outlook is currently running and has files locked -" -ForegroundColor Yellow
            Write-Host "  close Outlook first and try again if this backup is important." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "No new Outlook local data folder found at:" -ForegroundColor DarkYellow
        Write-Host "  $olkPath" -ForegroundColor DarkYellow
    }

    if (Test-Path $signaturesPath) {
        try {
            Write-Host "Copying signatures folder..." -ForegroundColor Yellow
            Copy-Item -Path $signaturesPath -Destination (Join-Path $backupFolder "Signatures") -Recurse -Force -ErrorAction Stop
            $copied += "Signatures"
            Write-Host "  Done." -ForegroundColor Green
        }
        catch {
            Write-Host "  Could not copy Signatures folder: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "No local Signatures folder found (normal for new Outlook, which stores" -ForegroundColor DarkYellow
        Write-Host "signatures in the cloud rather than locally)." -ForegroundColor DarkYellow
    }

    Write-Host ""
    if ($copied.Count -gt 0) {
        $sizeBytes = (Get-ChildItem -Path $backupFolder -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        $sizeMb = if ($sizeBytes) { [math]::Round($sizeBytes / 1MB, 1) } else { 0 }
        Write-Host "Backup complete: $($copied -join ', ')" -ForegroundColor Green
        Write-Host "Location: $backupFolder" -ForegroundColor Green
        Write-Host "Size: $sizeMb MB" -ForegroundColor Green
        return $backupFolder
    }
    else {
        Write-Host "Nothing was backed up (no matching local folders found)." -ForegroundColor Yellow
        return $null
    }
}

function Invoke-OutlookReset {
    param([switch]$DryRun)

    if ($DryRun) {
        Write-Host "[WhatIf] Would stop any running 'olk' processes." -ForegroundColor Magenta
        Write-Host "[WhatIf] Would wait 2 seconds." -ForegroundColor Magenta
        Write-Host "[WhatIf] Would run: olk.exe --reset" -ForegroundColor Magenta
        Write-Host "No changes were actually made (WhatIf mode)." -ForegroundColor Green
        return
    }

    Write-Host "Closing Outlook if it's running..." -ForegroundColor Yellow
    Get-Process -Name "olk" -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2

    Write-Host "Resetting new Outlook (local profile/cache only - nothing server-side is affected)..." -ForegroundColor Yellow
    try {
        Start-Process "olk.exe" -ArgumentList "--reset"
        Write-Host "Reset command sent. A confirmation dialog should appear in Outlook - follow the prompts." -ForegroundColor Green
    }
    catch {
        Write-Host "Could not launch olk.exe --reset automatically ($($_.Exception.Message))." -ForegroundColor Red
        Write-Host "You can run this manually: Win+R -> olk.exe --reset" -ForegroundColor Yellow
    }
}

function Show-Footer {
    Write-Host ""
    Write-Host "===================================================" -ForegroundColor Cyan
    Write-Host " If the account still fails to add:" -ForegroundColor Cyan
    Write-Host "  1. Wait 5-10 minutes (cache refresh isn't always instant)" -ForegroundColor White
    Write-Host "  2. Try again, choosing 'Advanced options' -> manual IMAP setup" -ForegroundColor White
    Write-Host "  3. Contact Jezweb support with a screenshot of the error" -ForegroundColor White
    Write-Host "===================================================" -ForegroundColor Cyan
}

function Confirm-BackupBeforeReset {
    $backupAnswer = Read-Host "Back up local Outlook data first (recommended)? (Y/n)"
    if ($backupAnswer -notmatch "^[Nn]") {
        Write-Host ""
        Backup-OutlookLocalData | Out-Null
        Write-Host ""
    }
}

# ===================================================================
# Menu actions
# ===================================================================

function Run-CacheRefreshOnly {
    Show-Header -Subtitle "Refresh Autodiscover Cache"
    Write-Host "This tells Microsoft to refresh its cached autodiscover settings for" -ForegroundColor White
    Write-Host "a failing account. It does NOT touch Outlook itself in any way." -ForegroundColor White
    Write-Host "If you need to reset Outlook as well, use the separate 'Reset new" -ForegroundColor White
    Write-Host "Outlook' option on the main menu." -ForegroundColor White
    Write-Host ""

    $email = $null
    while (-not $email) {
        $email = Read-EmailOrDomain -Prompt "Enter the full email address that is failing to add (e.g. name@yourdomain.com.au)"
    }

    Write-Host ""
    Invoke-AutodiscoverCacheRefresh -EmailAddress $email

    Write-Host "Try adding/re-adding the account in new Outlook now." -ForegroundColor Cyan
    Show-Footer
}

function Run-TestMode {
    Show-Header -Subtitle "Test Mode (safe - no changes to Outlook)"
    Write-Host "Checks autodiscover DNS records and sends the Microsoft cache-refresh" -ForegroundColor White
    Write-Host "request. Never touches Outlook. Safe to run against real domains too." -ForegroundColor White
    Write-Host ""

    $email = $null
    while (-not $email) {
        $email = Read-EmailOrDomain -Prompt "Enter a test email address, or just a domain (e.g. test@yourtestdomain.com.au)" -AllowBareDomain
    }

    $domain = $email.Split("@")[1]
    Write-Host ""
    Write-Host "Domain: $domain" -ForegroundColor Yellow
    Write-Host ""

    Test-AutodiscoverDns -Domain $domain
    Invoke-AutodiscoverCacheRefresh -EmailAddress $email

    Write-Host "===================================================" -ForegroundColor Cyan
    Write-Host " TEST MODE COMPLETE - Outlook was not touched." -ForegroundColor Cyan
    Write-Host " If DNS looks correct, try adding a real test mailbox in new" -ForegroundColor Cyan
    Write-Host " Outlook on this VM now to confirm the fix actually works." -ForegroundColor Cyan
    Write-Host "===================================================" -ForegroundColor Cyan
}

function Run-WhatIfMode {
    Show-Header -Subtitle "WhatIf Mode (dry run of the Outlook reset step)"
    Write-Host "Runs the full flow for real, except the Outlook reset step is only" -ForegroundColor White
    Write-Host "simulated (nothing closed or reset). Good for demos/training." -ForegroundColor White
    Write-Host ""

    $email = $null
    while (-not $email) {
        $email = Read-EmailOrDomain -Prompt "Enter the full email address that is failing to add (e.g. name@yourdomain.com.au)"
    }

    Write-Host ""
    Invoke-AutodiscoverCacheRefresh -EmailAddress $email

    $resetAnswer = Read-Host "Do you also want to reset new Outlook now (SIMULATED - nothing will actually change)? (y/N)"
    Write-Host ""
    if ($resetAnswer -match "^[Yy]") {
        Invoke-OutlookReset -DryRun
    }
    else {
        Write-Host "Skipping Outlook reset." -ForegroundColor Cyan
    }

    Show-Footer
}

function Run-BackupOnly {
    Show-Header -Subtitle "Backup Local Outlook Data"
    Write-Host "Copies new Outlook's local data folder (drafts/cache) and the" -ForegroundColor White
    Write-Host "signatures folder to a timestamped backup on the Desktop." -ForegroundColor White
    Write-Host ""
    Write-Host "Your actual mailbox (emails, contacts, calendar) lives on the mail" -ForegroundColor White
    Write-Host "server and is unaffected either way - this is just a local safety net." -ForegroundColor White
    Write-Host ""
    Backup-OutlookLocalData | Out-Null
}

function Run-OutlookResetOnly {
    Show-Header -Subtitle "Reset New Outlook Only"
    Write-Host "This clears new Outlook's LOCAL profile/cache and removes accounts" -ForegroundColor White
    Write-Host "from the profile. Nothing on any mail server is affected." -ForegroundColor White
    Write-Host "Use this if the account is already correctly set up but Outlook" -ForegroundColor White
    Write-Host "itself is stuck or misbehaving." -ForegroundColor White
    Write-Host ""

    $confirm = Read-Host "Are you sure you want to reset new Outlook now? (y/N)"
    Write-Host ""
    if ($confirm -match "^[Yy]") {
        Confirm-BackupBeforeReset
        Invoke-OutlookReset
    }
    else {
        Write-Host "Cancelled - no changes made." -ForegroundColor Cyan
    }
}

function Show-About {
    Show-Header -Subtitle "About This Tool"
    Write-Host "New Outlook doesn't query the mail server directly for autodiscover -" -ForegroundColor White
    Write-Host "it goes through a Microsoft cloud endpoint that caches domain settings." -ForegroundColor White
    Write-Host "After a domain moves to Axigen, that cache can go stale, causing:" -ForegroundColor White
    Write-Host "  - New Outlook failing to add the account" -ForegroundColor White
    Write-Host "  - Sign-in loops" -ForegroundColor White
    Write-Host "  - A false 'you might need an app password' error" -ForegroundColor White
    Write-Host ""
    Write-Host "This tool tells Microsoft to refresh that cached data. It does NOT" -ForegroundColor White
    Write-Host "delete or clear any mailbox/account data anywhere - server-side or" -ForegroundColor White
    Write-Host "otherwise. Refreshing the cache and resetting Outlook are two separate" -ForegroundColor White
    Write-Host "actions on the main menu - the cache refresh never touches Outlook," -ForegroundColor White
    Write-Host "and the reset only clears the LOCAL Outlook profile/cache on this PC." -ForegroundColor White
    Write-Host ""
    Write-Host "New Outlook has no PST/OST file like classic Outlook - all mail lives" -ForegroundColor White
    Write-Host "on the server. The backup option copies the local Olk cache/drafts" -ForegroundColor White
    Write-Host "folder as an extra safety net before any reset, not a full mail backup." -ForegroundColor White
    Write-Host ""
    Write-Host "The Microsoft autodetect endpoint used by cache refresh has changed" -ForegroundColor White
    Write-Host "hostname before without notice. This tool tries several known" -ForegroundColor White
    Write-Host "candidate hostnames automatically; if all fail, try https://aka.ms/autodetect" -ForegroundColor White
    Write-Host "manually." -ForegroundColor White
    Write-Host ""
    Write-Host "Repo: https://github.com/danieljezweb/outlook-autodiscover-fix" -ForegroundColor DarkGray
    Write-Host ""
    Read-Host "Press Enter to return to the menu"
}

# ===================================================================
# Main menu
# ===================================================================

function Show-MainMenu {
    while ($true) {
        Show-Header
        Write-Host "What would you like to do?" -ForegroundColor White
        Write-Host ""
        Write-Host "  1. Refresh autodiscover cache (fix a failing account - no Outlook changes)" -ForegroundColor White
        Write-Host "  2. Reset new Outlook (local profile/cache only)" -ForegroundColor White
        Write-Host "  3. Backup local Outlook data (recommended before a reset)" -ForegroundColor White
        Write-Host "  4. Test Mode - check DNS + refresh cache only (safe, no Outlook changes)" -ForegroundColor White
        Write-Host "  5. WhatIf Mode - dry run of a full fix + reset flow (safe, for demos)" -ForegroundColor White
        Write-Host "  6. About this tool" -ForegroundColor White
        Write-Host "  7. Exit" -ForegroundColor White
        Write-Host ""
        $choice = Read-Host "Enter a number (1-7)"
        Write-Host ""

        switch ($choice) {
            "1" { Run-CacheRefreshOnly; Read-Host "`nPress Enter to return to the menu" }
            "2" { Run-OutlookResetOnly; Read-Host "`nPress Enter to return to the menu" }
            "3" { Run-BackupOnly;       Read-Host "`nPress Enter to return to the menu" }
            "4" { Run-TestMode;         Read-Host "`nPress Enter to return to the menu" }
            "5" { Run-WhatIfMode;       Read-Host "`nPress Enter to return to the menu" }
            "6" { Show-About }
            "7" { Write-Host "Bye!" -ForegroundColor Cyan; return }
            default { Write-Host "Please enter a number from 1 to 7." -ForegroundColor Red; Start-Sleep -Seconds 1 }
        }
    }
}

# ===================================================================
# Entry point
# ===================================================================

$hasDirectParams = $EmailAddress -or $TestMode -or $WhatIf

if ($Menu -or -not $hasDirectParams) {
    Show-MainMenu
}
else {
    # Backward-compatible direct-parameter mode (for scripted/RMM use)
    if (-not $EmailAddress) {
        if ($TestMode) {
            $EmailAddress = Read-Host "Enter a test email address, or just a domain (e.g. test@yourtestdomain.com.au)"
            if ($EmailAddress -notmatch "@") { $EmailAddress = "autodiscover-test@$EmailAddress" }
        }
        else {
            $EmailAddress = Read-Host "Enter the full email address that is failing to add (e.g. name@yourdomain.com.au)"
        }
    }

    if ([string]::IsNullOrWhiteSpace($EmailAddress) -or ($EmailAddress -notmatch "^[^@\s]+@[^@\s]+\.[^@\s]+$")) {
        Write-Host "That doesn't look like a valid email address. Please re-run and try again." -ForegroundColor Red
        exit 1
    }

    $domain = $EmailAddress.Split("@")[1]

    if ($TestMode) {
        Test-AutodiscoverDns -Domain $domain
        Invoke-AutodiscoverCacheRefresh -EmailAddress $EmailAddress
        Write-Host "TEST MODE COMPLETE - Outlook was not touched." -ForegroundColor Cyan
        exit 0
    }

    Invoke-AutodiscoverCacheRefresh -EmailAddress $EmailAddress
    Write-Host "Try adding/re-adding the account in new Outlook now." -ForegroundColor Cyan
    Write-Host "(Cache refresh and Outlook reset are separate steps - use -Menu or run" -ForegroundColor DarkGray
    Write-Host " the script with no parameters if you also need to reset Outlook.)" -ForegroundColor DarkGray

    Show-Footer
}
