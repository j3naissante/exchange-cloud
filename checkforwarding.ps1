# Path to file containing user accounts (one per line)
$UserList = "path txtfile"
$Users = Get-Content $UserList

# Prepare results array
$Results = @()

foreach ($User in $Users) {
    Write-Host "Checking inbox rules for $User..." -ForegroundColor Yellow

    try {
        # Get inbox rules for the user
        $rules = Get-InboxRule -Mailbox $User -ErrorAction Stop

        # Filter rules that forward or redirect
        $forwardRules = $rules | Where-Object {
            $_.ForwardTo -ne $null -or
            $_.ForwardAsAttachmentTo -ne $null -or
            $_.RedirectTo -ne $null
        }

        if ($forwardRules) {
            foreach ($rule in $forwardRules) {
                $Results += [PSCustomObject]@{
                    User       = $User
                    RuleName   = $rule.Name
                    ForwardTo  = ($rule.ForwardTo | ForEach-Object { $_.Name }) -join ", "
                    ForwardAsAttachmentTo = ($rule.ForwardAsAttachmentTo | ForEach-Object { $_.Name }) -join ", "
                    RedirectTo = ($rule.RedirectTo | ForEach-Object { $_.Name }) -join ", "
                }
            }
        }
    }
    catch {
        Write-Host "Error processing $User : $_" -ForegroundColor Red
    }
}

# Display results
if ($Results.Count -gt 0) {
    $Results | Format-Table -AutoSize
    # Optional: export to CSV
    $Results | Export-Csv -Path "C:\Temp\ForwardingReport.csv" -NoTypeInformation
    Write-Host "Report exported to C:\Temp\ForwardingReport.csv" -ForegroundColor Cyan
} else {
    Write-Host "No forwarding rules found for any users." -ForegroundColor Green
}