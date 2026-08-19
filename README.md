# Outlook Autodiscover Fix

PowerShell tool to fix new Outlook autodiscover/sign-in failures after Rackspace → Axigen email migration.

## The problem

New Outlook for Windows doesn't query the mail server directly for autodiscover — it goes through a Microsoft cloud endpoint which caches domain configuration. After a domain is migrated to a new mail server (e.g. Rackspace → Axigen), that cache can be stale, causing new Outlook to:

- Fail to add the account entirely
- Loop on sign-in
- Prompt for an "app password" (which Axigen doesn't support)

This tool refreshes Microsoft's cache for the affected email address, and separately can reset new Outlook so it picks up the corrected settings cleanly. **It does not delete or clear any mailbox/account data anywhere** — refreshing the cache never touches Outlook at all, and the separate reset option only clears the local Outlook profile on that one PC.

> **Note on the Microsoft endpoint:** Microsoft has changed the hostname behind this cache-refresh endpoint before without documenting it (`prod-autodetect.outlookmobile.com` has previously started returning NXDOMAIN). The script tries several known candidate hostnames automatically and reports which one worked. If none of them work, use the manual tool at https://aka.ms/autodetect as a fallback.

## Quick run (no download needed)

Open PowerShell and run:

```powershell
irm https://raw.githubusercontent.com/danieljezweb/outlook-autodiscover-fix/main/Fix-OutlookAutodiscover.ps1 -OutFile "$env:TEMP\Fix-Outlook.ps1"; powershell -ExecutionPolicy Bypass -File "$env:TEMP\Fix-Outlook.ps1"
```

This downloads the script to a temp folder and launches the **interactive main menu**.

The `-ExecutionPolicy Bypass` only applies to this one launch; it does not change any system-wide PowerShell setting on the machine. This is included by default because most Windows PCs have a `Restricted` execution policy out of the box, which otherwise blocks the script from running with an `UnauthorizedAccess` error.

> Note: a plain `irm ... | iex` one-liner also works and avoids the execution-policy issue entirely, but skips straight to the menu the same way:
> ```powershell
> irm https://raw.githubusercontent.com/danieljezweb/outlook-autodiscover-fix/main/Fix-OutlookAutodiscover.ps1 | iex
> ```

## The main menu

Running the script with no parameters shows:

```
1. Refresh autodiscover cache (fix a failing account - no Outlook changes)
2. Reset new Outlook (local profile/cache only)
3. Backup local Outlook data (recommended before a reset)
4. Test Mode - check DNS + refresh cache only (safe, no Outlook changes)
5. WhatIf Mode - dry run of a full fix + reset flow (safe, for demos)
6. About this tool
7. Exit
```

Refreshing the cache and resetting Outlook are **two completely separate options** — running option 1 never touches Outlook, and option 2 never touches Microsoft's autodiscover cache. This keeps each action predictable and means nobody accidentally resets a profile while just trying to fix a cache issue.

- **Option 1** is the normal client-facing fix: prompts for the email address and refreshes Microsoft's autodiscover cache only. Tries multiple known Microsoft endpoint hostnames in case the primary one has changed.
- **Option 2** resets new Outlook's local profile/cache on its own, offering a backup first. Use this for cases where the account is already set up correctly but Outlook itself is stuck.
- **Option 3 (Backup)** copies new Outlook's local data folder and signatures to a timestamped folder on the Desktop, as a safety net before any reset. See below for what this actually covers.
- **Option 4 (Test Mode)** checks autodiscover DNS records (CNAME/A + SRV) for the domain and sends the cache-refresh request, but never touches Outlook and doesn't need a real mailbox. Safe to run against production domains.
- **Option 5 (WhatIf Mode)** runs the full fix + reset flow for demonstration purposes, with the Outlook reset step only simulated — useful for demos or training without risking a real profile.
- **Option 6** shows a short explanation of what the tool does and doesn't do.

## About the backup option

New Outlook doesn't use a PST/OST data file the way classic Outlook does — at time of writing, new Outlook for Windows does not support personal folders (PST) at all. All mail, contacts, and calendar data lives on the mail server and is synced live; a local reset doesn't delete any of that.

What the backup option actually copies, as an extra safety net:

- `%LocalAppData%\Microsoft\Olk` — new Outlook's local cache/drafts folder
- `%AppData%\Microsoft\Signatures` — classic signatures folder, if present (new Outlook normally stores signatures in the cloud, so this folder may not exist)

Backups are saved to a timestamped folder on the Desktop, e.g. `OutlookBackup_20260819_143000`. This is a local safety net for drafts/cache, not a full mailbox export — there is nothing to "restore" from it in the traditional PST-import sense, but it preserves anything that may not have fully synced yet.

## About the autodetect endpoint fallback

The cache-refresh request goes to a Microsoft-hosted endpoint that this tool doesn't control. Microsoft has been observed changing this endpoint's hostname without notice (`prod-autodetect.outlookmobile.com` has previously returned a DNS `NXDOMAIN` error). To stay resilient, the script:

1. Tries `prod-autodetect.outlookmobile.com` first (the historically documented endpoint)
2. Falls back to `prod.autodetect.outlook.cloud.microsoft` if that fails
3. Falls back to `autodetect.outlookmobile.com` as a last resort
4. Reports clearly which endpoint (if any) actually responded

If all three fail, try the manual tool at https://aka.ms/autodetect, or check whether Microsoft has published a newer endpoint.

## Direct-parameter mode (for scripting / RMM)

Parameters can be passed directly to skip the menu entirely, e.g. for pushing via an RMM tool:

```powershell
.\Fix-OutlookAutodiscover.ps1 -EmailAddress name@client-domain.com.au
```

This refreshes the autodiscover cache only — it will not touch Outlook. To reset Outlook via script/RMM as well, run the tool separately with `-Menu` and select option 2, or call it interactively.

```powershell
.\Fix-OutlookAutodiscover.ps1 -TestMode -EmailAddress test@yourtestdomain.com.au
```

```powershell
.\Fix-OutlookAutodiscover.ps1 -WhatIf
```

Pass `-Menu` alongside any of the above to force the menu to show anyway.

If running the saved `.ps1` file directly (rather than the quick-run one-liner above) and you hit an `UnauthorizedAccess`/`cannot be loaded because running scripts is disabled` error, either:

- Run it via `powershell -ExecutionPolicy Bypass -File .\Fix-OutlookAutodiscover.ps1`, or
- Run `Unblock-File .\Fix-OutlookAutodiscover.ps1` once, then run it normally.

## Recommended test process

1. Snapshot a clean test VM (e.g. Windows 11 with new Outlook installed).
2. Use **Test Mode** against a test domain — confirm DNS results look correct and check which autodetect endpoint responds.
3. Add a real test mailbox in new Outlook on the VM, let it fail.
4. Use **option 1 (Refresh autodiscover cache)** against that real test mailbox to confirm the cache-refresh step works, then try re-adding the account.
5. If the account still won't add, use **option 2 (Reset new Outlook)** separately.
6. Revert the VM snapshot and repeat for other scenarios as needed.

## If it still fails

1. Wait 5–10 minutes — the Microsoft cache refresh isn't always instant.
2. Try adding the account again, choosing **Advanced options → manual IMAP setup** instead of relying on autodiscover.
3. Contact Jezweb support with a screenshot of the error.
