# Variables
$DLName   = ""   # The alias or SMTP address of the DL
$TxtFile  = ""            # Path to TXT file with SMTP addresses (one per line)

# Read the addresses from file
$AddressesToRemove = Get-Content -Path $TxtFile

# Get DL members
$DLMembers = Get-DistributionGroupMember -Identity $DLName -ResultSize Unlimited

foreach ($Address in $AddressesToRemove) {
    # Trim whitespace
    $CleanAddress = $Address.Trim().ToLower()

    # Find matching member in the DL
    $Member = $DLMembers | Where-Object { $_.PrimarySmtpAddress -eq $CleanAddress }

    if ($Member) {
        Write-Host "Removing $CleanAddress from $DLName..." -ForegroundColor Yellow
        Remove-DistributionGroupMember -Identity $DLName -Member $Member.PrimarySmtpAddress -Confirm:$false
    }
    else {
        Write-Host "Address $CleanAddress not found in $DLName" -ForegroundColor Red
    }
}