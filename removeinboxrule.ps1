# Path to file containing user accounts (one per line, UPN or email)
$UserList = ""

# Read user accounts
$Users = Get-Content $UserList

f
foreach ($User in $Users) {
    Write-Host "Checking inbox rules for $User..." -ForegroundColor Yellow
    
    try {
        # Get inbox rules for the user
        $rules = Get-InboxRule -Mailbox $User -ErrorAction Stop

        # Find rules that forward or redirect
        $forwardRules = $rules | Where-Object {
            $_.ForwardTo -ne $null -or
            $_.ForwardAsAttachmentTo -ne $null -or
            $_.RedirectTo -ne $null
        }

        if ($forwardRules) {
            foreach ($rule in $forwardRules) {
                Write-Host "Removing rule '$($rule.Name)' from $User" -ForegroundColor Red
                
                # Correct removal using pipeline
                $rule | Remove-InboxRule -Confirm:$false
            }
        }
        else {
            Write-Host "No forwarding rules found for $User." -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Error processing $User : $_" -ForegroundColor Red
    }
}