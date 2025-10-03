# Distribution list (alias, display name, or email address)
$DistributionList = "email"

# List of email addresses to add
$ContactEmails = @(
 ""

)

foreach ($Email in $ContactEmails) {
    try {
        # In Exchange Online, external recipients are usually MailUsers, MailContacts, or just raw SMTP addresses.
        $Recipient = Get-Recipient -Identity $Email -ErrorAction SilentlyContinue

        if ($Recipient) {
            Add-DistributionGroupMember -Identity $DistributionList -Member $Recipient.Identity -Confirm:$false
            Write-Host "Added $Email to $DistributionList"
        } else {
            # If no matching recipient exists, add directly by SMTP (creates external member entry)
            try {
                Add-DistributionGroupMember -Identity $DistributionList -Member $Email -Confirm:$false
                Write-Host "Added external $Email to $DistributionList"
            } catch {
                Write-Warning "Could not add ${Email}: $_"

            }
        }

    } catch {
        Write-Warning "Error processing ${Email}: $_"
    }
}


