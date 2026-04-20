#Requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
    Creates a Distribution Group or Mail-Enabled Security Group in Exchange Online.

.DESCRIPTION
    Prompts for group details including type, owner, ExtensionAttribute2,
    and optionally imports members from a TXT file.

.NOTES
    TXT file for member import must contain one UPN per line.
    Example TXT:
        user1@contoso.com
        user2@contoso.com
#>

#region -- Helper --------------------------------------------------------------

function Write-Header {
    param([string]$Text)
    Write-Host "`n$('-' * 60)" -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host "$('-' * 60)" -ForegroundColor Cyan
}

function Prompt-YesNo {
    param([string]$Question)
    do {
        $answer = (Read-Host "$Question [Y/N]").Trim().ToUpper()
    } while ($answer -notin @('Y', 'N'))
    return $answer -eq 'Y'
}

#endregion

#region -- Connect -------------------------------------------------------------

Write-Header "Exchange Online - Mail Group Creator"

$alreadyConnected = $false
try {
    $null = Get-OrganizationConfig -ErrorAction Stop
    $alreadyConnected = $true
    Write-Host "  Already connected to Exchange Online." -ForegroundColor Green
} catch {
    Write-Host "  Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowBanner:$false
}

#endregion

#region -- Group Type ----------------------------------------------------------

Write-Header "Step 1 - Group Type"
Write-Host "  1) Distribution Group"
Write-Host "  2) Mail-Enabled Security Group"
do {
    $typeChoice = (Read-Host "  Select group type [1/2]").Trim()
} while ($typeChoice -notin @('1', '2'))

$groupType = if ($typeChoice -eq '1') { 'Distribution' } else { 'Security' }
Write-Host "  -> $groupType selected." -ForegroundColor Green

#endregion

#region -- Group Details -------------------------------------------------------

Write-Header "Step 2 - Group Details"

do {
    $displayName = (Read-Host "  Display Name").Trim()
} while ([string]::IsNullOrWhiteSpace($displayName))

do {
    $alias = (Read-Host "  Alias (no spaces, no @domain)").Trim()
} while ($alias -notmatch '^[A-Za-z0-9._-]+$')

do {
    $primarySMTP = (Read-Host "  Primary SMTP address (e.g. group@contoso.com)").Trim()
} while ($primarySMTP -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$')

#endregion

#region -- Owner ---------------------------------------------------------------

Write-Header "Step 3 - Owner"
do {
    $ownerUPN = (Read-Host "  Owner UPN (e.g. admin@contoso.com)").Trim()
} while ($ownerUPN -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$')

# Validate the owner exists
try {
    $ownerObj = Get-Recipient -Identity $ownerUPN -ErrorAction Stop
    Write-Host "  Owner found: $($ownerObj.DisplayName)" -ForegroundColor Green
} catch {
    Write-Warning "  Owner '$ownerUPN' not found in Exchange. Proceeding anyway - verify manually."
}

#endregion

#region -- ExtensionAttribute2 ------------------------------------------------

Write-Header "Step 4 - ExtensionAttribute2"
$extensionAttribute2 = (Read-Host "  ExtensionAttribute2 value (leave blank to skip)").Trim()

if ([string]::IsNullOrWhiteSpace($extensionAttribute2)) {
    $extensionAttribute2 = $null
    Write-Host "  -> ExtensionAttribute2 will not be set." -ForegroundColor DarkGray
} else {
    Write-Host "  -> ExtensionAttribute2 = '$extensionAttribute2'" -ForegroundColor Green
}

#endregion

#region -- Create Group --------------------------------------------------------

Write-Header "Step 5 - Creating Group"

$groupParams = @{
    Name             = $displayName
    DisplayName      = $displayName
    Alias            = $alias
    PrimarySmtpAddress = $primarySMTP
    ManagedBy        = $ownerUPN
    Type             = $groupType
}

try {
    $newGroup = New-DistributionGroup @groupParams -ErrorAction Stop
    Write-Host "  [OK] Group created: $($newGroup.PrimarySmtpAddress)" -ForegroundColor Green
} catch {
    Write-Error "  Failed to create group: $_"
    exit 1
}

# Set ExtensionAttribute2 if provided
if ($extensionAttribute2) {
    try {
        Set-DistributionGroup -Identity $newGroup.Identity `
            -CustomAttribute2 $extensionAttribute2 -ErrorAction Stop
        Write-Host "  [OK] ExtensionAttribute2 set to '$extensionAttribute2'." -ForegroundColor Green
    } catch {
        Write-Warning "  Could not set ExtensionAttribute2: $_"
    }
}

#endregion

#region -- Import Members ------------------------------------------------------

Write-Header "Step 6 - Import Members"

$importMembers = Prompt-YesNo "  Do you want to import members from a TXT file?"

if ($importMembers) {
    do {
        $txtPath = (Read-Host "  Full path to TXT file (one UPN per line)").Trim().Trim('"')
        if (-not (Test-Path $txtPath)) {
            Write-Warning "  File not found. Please try again."
        }
    } while (-not (Test-Path $txtPath))

    try {
        $lines = Get-Content -Path $txtPath -ErrorAction Stop
    } catch {
        Write-Error "  Failed to read TXT file: $_"
        exit 1
    }

    # Filter out blank lines and comments (#)
    $members = $lines | Where-Object { $_.Trim() -ne '' -and $_ -notmatch '^\s*#' }

    if ($members.Count -eq 0) {
        Write-Warning "  TXT file is empty or contains no valid entries. Skipping member import."
    } else {
        $successCount = 0
        $failCount    = 0

        Write-Host "  Adding $($members.Count) member(s)..." -ForegroundColor Yellow

        foreach ($line in $members) {
            $upn = $line.Trim()

            if ($upn -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                Write-Warning "    [FAIL] Skipped (not a valid UPN): '$upn'"
                $failCount++
                continue
            }

            try {
                Add-DistributionGroupMember -Identity $newGroup.Identity `
                    -Member $upn -ErrorAction Stop
                Write-Host "    [OK] Added: $upn" -ForegroundColor Green
                $successCount++
            } catch {
                Write-Warning "    [FAIL] Failed to add '$upn': $_"
                $failCount++
            }
        }

        Write-Host "`n  Members added   : $successCount" -ForegroundColor Green
        if ($failCount -gt 0) {
            Write-Host "  Members failed  : $failCount" -ForegroundColor Red
        }
    }
} else {
    Write-Host "  -> Skipping member import." -ForegroundColor DarkGray
}

#endregion

#region -- Summary -------------------------------------------------------------

Write-Header "[OK] Done - Group Summary"
Write-Host "  Display Name       : $displayName"
Write-Host "  Alias              : $alias"
Write-Host "  Primary SMTP       : $primarySMTP"
Write-Host "  Type               : $groupType"
Write-Host "  Owner              : $ownerUPN"
Write-Host "  ExtensionAttribute2: $(if ($extensionAttribute2) { $extensionAttribute2 } else { '(not set)' })"
Write-Host ""

#endregion