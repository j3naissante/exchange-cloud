# Distribution list name
$DistributionList = "email"

# Path to text file containing email addresses (one per line)
$EmailFile = "path.txtfile"

# Read email addresses from file (ignore blanks)
$ContactEmails = Get-Content $EmailFile | Where-Object { $_.Trim() -ne "" }

foreach ($Email in $ContactEmails) {
    try {
        # Check if already a member before adding
        $isMember = Get-DistributionGroupMember -Identity $DistributionList -ResultSize Unlimited |
                    Where-Object { $_.PrimarySmtpAddress -eq $Email }

        if (-not $isMember) {
            Add-DistributionGroupMember -Identity $DistributionList -Member $Email -Confirm:$false
            Write-Host "Added $Email to $DistributionList"
        } else {
            Write-Host "$Email is already a member of $DistributionList"
        }

    } catch {
        Write-Warning "Error processing ${Email}: $_"
    }
}