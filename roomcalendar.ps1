
# 1. Variables
$RoomName = ""                       # Name
$RoomAlias = ""                      # Alias
$RoomEmail = ""      # Primary SMTP
$ManagerGroup = ""     # Group who manages the calendar 

# 2.  Check if the mailbox already exists
if (Get-Recipient -Identity $RoomAlias -ErrorAction SilentlyContinue) {
    Write-Host "Mailbox $RoomAlias exists. Quitting..." -ForegroundColor Red
    return
}

# 3. Create room
New-Mailbox -Room -Name $RoomName -Alias $RoomAlias -DisplayName $RoomName
Set-Mailbox -Identity $RoomAlias -PrimarySmtpAddress $RoomEmail

# 4. Assign permissions only to the management group
Add-MailboxPermission -Identity $RoomAlias -User $ManagerGroup -AccessRights FullAccess -InheritanceType All

# Remove Full Access permissions from everyone else
Get-MailboxPermission -Identity $RoomAlias | Where-Object {
    ($_.User -notlike "NT AUTHORITY\SELF") -and ($_.User -notlike $ManagerGroup)
} | ForEach-Object {
    Remove-MailboxPermission -Identity $RoomAlias -User $_.User -AccessRights FullAccess -Confirm:$false
}

# Grant SendAs permissions to the management group
Add-RecipientPermission -Identity $RoomAlias -Trustee $ManagerGroup -AccessRights SendAs

# Remove SendAs permissions from everyone else
Get-RecipientPermission -Identity $RoomAlias | Where-Object {
    ($_.Trustee -notlike $ManagerGroup)
} | ForEach-Object {
    Remove-RecipientPermission -Identity $RoomAlias -Trustee $_.Trustee -AccessRights SendAs -Confirm:$false
}

# 5.  Configure automatic booking for the management group only
Set-CalendarProcessing -Identity $RoomAlias `
    -AutomateProcessing AutoAccept `
    -AllBookInPolicy $false `
    -BookInPolicy $ManagerGroup `
    -RequestInPolicy $ManagerGroup `
    -AllRequestOutOfPolicy $false `
    -AllowConflicts $false `
    -BookingWindowInDay 180

# 6. Verification
Write-Host "Room calendar created successfully" -ForegroundColor Green
Get-Mailbox -Identity $RoomAlias | Format-List Name,Alias,PrimarySmtpAddress,RecipientTypeDetails,HiddenFromAddressListsEnabled
Get-MailboxPermission -Identity $RoomAlias
Get-RecipientPermission -Identity $RoomAlias
