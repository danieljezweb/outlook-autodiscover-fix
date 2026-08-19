<#
.SYNOPSIS
    Fixes new Outlook autodiscover issues after migrating an email account to Axigen (ax.email).

.DESCRIPTION
    New Outlook for Windows does not query the mail server directly for autodiscover -
    it goes through a Microsoft cloud endpoint (prod-autodetect.outlookmobile.com) which
    caches domain configuration. After a domain is migrated to a new mail server (e.g.
    Rackspace -> Axigen), that cache can be stale, causing new Outlook to fail to add
    the account, loop on sign-in, or ask for an "app password".

    This script forces Microsoft's cache to refresh for the given email address, then
    optionally resets new Outlook so it picks up the corrected settings cleanly.

.PARAMETER EmailAddress
    The email address to refresh. If omitted, the script will prompt for it
    (except in -TestMode, where a domain is enough).

.PARAMETER TestMode
    Safe validation mode for use on a test VM or before a client rollout.
    - Checks that autodiscover DNS records resolve correctly for the domain
    - Sends the Microsoft cache-refresh request (this part is non-destructive
      even against a real domain, so it's safe to run against production domains too)
    - Never touches Outlook (no process kill, no reset) - just reports what it finds
    - Does not require a real mailbox to exist

.PARAMETER WhatIf
    Shows what the script WOULD do at the Outlook-reset step, without actually
    closing Outlook or resetting it. Use this to demo/validate the script flow
    safely. The DNS check and cache-refresh request still run for real, since
    those are non-destructive.

.EXAMPLE
    .\Fix-OutlookAutodiscover.ps1 -TestMode -EmailAddress test@yourtestdomain.com.au
    Validates DNS + cache refresh only, for use on a test VM. Nothing on the
    machine is changed.

.EXAMPLE
    .\Fix-OutlookAutodiscover.ps1 -WhatIf
    Runs the full interactive flow but only simulates the Outlook reset step.

.EXAMPLE
    .\Fix-OutlookAutodiscover.ps1 -EmailAddress name@client-domain.com.au
    Full run for a real client account (will still prompt before resetting Outlook).
#>

param(
    [string]$EmailAddress,
    [switch]$TestMode,
    [switch]$WhatIf
)

Write-Host "===================================================" -ForegroundColor Cyan
Write-Host " Outlook Autodiscover Cache Refresh (Jezweb)" -ForegroundColor Cyan
if ($TestMode) { Write-Host " *** TEST MODE - no changes will be made to Outlook ***" -ForegroundColor Magenta }
if ($WhatIf)   { Write-Host " *** -WhatIf active - Outlook reset step will be simulated only ***" -ForegroundColor Magenta }
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host ""

# --- Step 1: Get the email address / domain to work with ---
if ([string]::IsNullOrWhiteSpace($EmailAddress)) {
    if ($TestMode) {
        $EmailAddress = Read-Host "Enter a test email address, or just a domain (e.g. test@yourtestdomain.com.au)"
    }
    else {
        $EmailAddress = Read-Host "Enter the full email address that is failing to add (e.g. name@yourdomain.com.au)"
    }
}

# Allow TestMode to accept a bare domain (no @) and turn it into a throwaway address
if ($TestMode -and ($EmailAddress -notmatch "@")) {
    $EmailAddress = "autodiscover-test@$EmailAddress"
}

if ([string]::IsNullOrWhiteSpace($EmailAddress) -or ($EmailAddress -notmatch "^[^@\s]+@[^@\s]+\.[^@\s]+$")) {
    Write-Host "That doesn't look like a valid email address (or domain in -TestMode). Please re-run and try again." -ForegroundColor Red
    exit 1
}

$domain = $EmailAddress.Split("@")[1]

Write-Host ""
Write-Host "Working with:" -ForegroundColor Yellow
Write-Host "  Email/identifier : $EmailAddress" -ForegroundColor Yellow
Write-Host "  Domain           : $domain" -ForegroundColor Yellow
Write-Host ""

# --- Step 2 (TestMode only): Check autodiscover DNS records resolve ---
if ($TestMode) {
    Write-Host "--- DNS check: autodiscover.$domain ---" -ForegroundColor Cyan
    try {
        $cname = Resolve-DnsName -Name "autodiscover.$domain" -Type CNAME -ErrorAction Stop
        Write-Host "CNAME found:" -ForegroundColor Green
        $cname | Select-Object Name, NameHost | Format-Table -AutoSize | Out-String | Write-Host
    }
    catch {
        Write-Host "No CNAME record found for autodiscover.$domain (or it's an A record instead)." -ForegroundColor Yellow
        try {
            $arec = Resolve-DnsName -Name "autodiscover.$domain" -Type A -ErrorAction Stop
            Write-Host "A record found:" -ForegroundColor Green
            $arec | Select-Object Name, IPAddress | Format-Table -AutoSize | Out-String | Write-Host
        }
        catch {
            Write-Host "No autodiscover record resolves at all for $domain - this needs fixing in DNS before Outlook autodiscover can work." -ForegroundColor Red
        }
    }

    Write-Host "--- DNS check: SRV records ---" -ForegroundColor Cyan
    foreach ($srv in @("_imaps._tcp.$domain", "_submission._tcp.$domain", "_autodiscover._tcp.$domain")) {
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

# --- Step 3: Force the Microsoft cloud cache to refresh (safe / non-destructive) ---
Write-Host "Requesting Microsoft to refresh its cached autodiscover data for:" -ForegroundColor Yellow
Write-Host "  $EmailAddress" -ForegroundColor Yellow
Write-Host ""

try {
    $uri = "https://prod-autodetect.outlookmobile.com/detect?services=office365,outlook,google,icloud,yahoo&protocols=rest-cloud,rest-outlook,rest-office365,eas,imap,smtp"
    $response = Invoke-WebRequest -Uri $uri -Headers @{ "x-email" = $EmailAddress } -UseBasicParsing -ErrorAction Stop
    Write-Host "Request sent successfully (HTTP $($response.StatusCode))." -ForegroundColor Green
    Write-Host ""
    Write-Host "Response headers:" -ForegroundColor DarkGray
    $response.Headers | Format-Table -AutoSize | Out-String | Write-Host
}
catch {
    Write-Host "The refresh request failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "You can still continue and try adding the account again - this step is often silent even when it works." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Cache refresh request complete." -ForegroundColor Green
Write-Host ""

# --- Step 4: TestMode stops here - never touches Outlook ---
if ($TestMode) {
    Write-Host "===================================================" -ForegroundColor Cyan
    Write-Host " TEST MODE COMPLETE - Outlook was not touched." -ForegroundColor Cyan
    Write-Host " Review the DNS results and cache-refresh response above." -ForegroundColor Cyan
    Write-Host " If DNS looks correct, try adding a real test mailbox in new" -ForegroundColor Cyan
    Write-Host " Outlook on this VM now to confirm the fix actually works." -ForegroundColor Cyan
    Write-Host "===================================================" -ForegroundColor Cyan
    exit 0
}

# --- Step 5: Offer to reset new Outlook (real runs only) ---
$resetAnswer = Read-Host "Do you also want to reset new Outlook now (clears local accounts/cache, you'll re-add accounts after)? (y/N)"

if ($resetAnswer -match "^[Yy]") {
    Write-Host ""
    if ($WhatIf) {
        Write-Host "[WhatIf] Would stop any running 'olk' processes." -ForegroundColor Magenta
        Write-Host "[WhatIf] Would wait 2 seconds." -ForegroundColor Magenta
        Write-Host "[WhatIf] Would run: olk.exe --reset" -ForegroundColor Magenta
        Write-Host "No changes were actually made (WhatIf mode)." -ForegroundColor Green
    }
    else {
        Write-Host "Closing Outlook if it's running..." -ForegroundColor Yellow
        Get-Process -Name "olk" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 2

        Write-Host "Resetting new Outlook..." -ForegroundColor Yellow
        try {
            Start-Process "olk.exe" -ArgumentList "--reset"
            Write-Host "Reset command sent. A confirmation dialog should appear in Outlook - follow the prompts." -ForegroundColor Green
        }
        catch {
            Write-Host "Could not launch olk.exe --reset automatically ($($_.Exception.Message))." -ForegroundColor Red
            Write-Host "You can run this manually: Win+R -> olk.exe --reset" -ForegroundColor Yellow
        }
    }
}
else {
    Write-Host ""
    Write-Host "Skipping Outlook reset. Try adding/re-adding the account in new Outlook now." -ForegroundColor Cyan
}

Write-Host ""
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host " Done. If the account still fails to add:" -ForegroundColor Cyan
Write-Host "  1. Wait 5-10 minutes (cache refresh isn't always instant)" -ForegroundColor White
Write-Host "  2. Try again, choosing 'Advanced options' -> manual IMAP setup" -ForegroundColor White
Write-Host "  3. Contact Jezweb support with a screenshot of the error" -ForegroundColor White
Write-Host "===================================================" -ForegroundColor Cyan
