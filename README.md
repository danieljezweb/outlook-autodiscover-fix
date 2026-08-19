# Outlook Autodiscover Fix

PowerShell tool to fix new Outlook autodiscover/sign-in failures after Rackspace → Axigen email migration.

## The problem

New Outlook for Windows doesn't query the mail server directly for autodiscover — it goes through a Microsoft cloud endpoint (`prod-autodetect.outlookmobile.com`) which caches domain configuration. After a domain is migrated to a new mail server (e.g. Rackspace → Axigen), that cache can be stale, causing new Outlook to:

- Fail to add the account entirely
- Loop on sign-in
- Prompt for an "app password" (which Axigen doesn't support)

This script forces Microsoft's cache to refresh for the affected email address, and optionally resets new Outlook so it picks up the corrected settings cleanly.

## Quick run (no download needed)

Open PowerShell and run:

```powershell
irm https://raw.githubusercontent.com/danieljezweb/outlook-autodiscover-fix/main/Fix-OutlookAutodiscover.ps1 -OutFile "$env:TEMP\Fix-Outlook.ps1"; powershell -ExecutionPolicy Bypass -File "$env:TEMP\Fix-Outlook.ps1"
```

This downloads the script to a temp folder and runs it interactively — it will prompt for the email address that's failing, then ask before making any changes to Outlook.

The `-ExecutionPolicy Bypass` only applies to this one launch of the script; it does not change any system-wide PowerShell setting on the machine. This is included by default because most Windows PCs have a `Restricted` execution policy out of the box, which otherwise blocks the script from running with an `UnauthorizedAccess` error.

> Note: a plain `irm ... | iex` one-liner works too and avoids the execution-policy issue entirely (since the script runs in-memory rather than as a file), but won't let you pass parameters like `-TestMode`. Use it like this if you just need the default interactive run:
> ```powershell
> irm https://raw.githubusercontent.com/danieljezweb/outlook-autodiscover-fix/main/Fix-OutlookAutodiscover.ps1 | iex
> ```

## Usage

### Standard run (client-facing)

```powershell
.\Fix-OutlookAutodiscover.ps1
```

Prompts for the email address, refreshes Microsoft's autodiscover cache, then asks before optionally resetting new Outlook.

```powershell
.\Fix-OutlookAutodiscover.ps1 -EmailAddress name@client-domain.com.au
```

Same as above, but skips the email prompt.

If running the saved `.ps1` file directly (rather than the quick-run one-liner above) and you hit an `UnauthorizedAccess`/`cannot be loaded because running scripts is disabled` error, either:

- Run it via `powershell -ExecutionPolicy Bypass -File .\Fix-OutlookAutodiscover.ps1`, or
- Run `Unblock-File .\Fix-OutlookAutodiscover.ps1` once, then run it normally.

### Test mode (safe, for validation before rollout)

```powershell
.\Fix-OutlookAutodiscover.ps1 -TestMode -EmailAddress test@yourtestdomain.com.au
```

- Checks autodiscover DNS records (CNAME/A + SRV) for the domain
- Sends the Microsoft cache-refresh request (non-destructive, safe even on production domains)
- **Never touches Outlook** — no process kill, no reset
- Doesn't require a real mailbox to exist

You can also just pass a bare domain in test mode:

```powershell
.\Fix-OutlookAutodiscover.ps1 -TestMode -EmailAddress yourtestdomain.com.au
```

### WhatIf mode (dry run of the Outlook reset step)

```powershell
.\Fix-OutlookAutodiscover.ps1 -WhatIf
```

Runs the full flow, but only *simulates* the Outlook reset step (closing Outlook, running `olk.exe --reset`) instead of actually doing it. Everything else (DNS/cache refresh) runs for real since it's non-destructive.

## Recommended test process

1. Snapshot a clean test VM (e.g. Windows 11 with new Outlook installed).
2. Run `-TestMode` against a test domain — confirm DNS results look correct.
3. Add a real test mailbox in new Outlook on the VM, let it fail.
4. Run the full script (no switches) against that real test mailbox to confirm the fix works end-to-end.
5. Revert the VM snapshot and repeat for other scenarios as needed.

## If it still fails

1. Wait 5–10 minutes — the Microsoft cache refresh isn't always instant.
2. Try adding the account again, choosing **Advanced options → manual IMAP setup** instead of relying on autodiscover.
3. Contact Jezweb support with a screenshot of the error.
