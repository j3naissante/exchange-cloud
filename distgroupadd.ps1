# Define the distribution group
$DG = ""

# Import users from the txt file
$Users = Get-Content -Path ""

# Loop through each user and add to DG
foreach ($User in $Users) {
    try {
        Add-DistributionGroupMember -Identity $DG -Member $User -ErrorAction Stop
        Write-Host "Added $User to $DG" -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to add $User to $DG. Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}