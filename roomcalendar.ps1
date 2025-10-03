# -------------------------------
# Ruumikalendri Loomise Skript
# -------------------------------


# 1. Muutujad
$RoomName = ""                       # Kuvatav nimi
$RoomAlias = ""                      # Alias
$RoomEmail = ""      # Primaarsmtp
$ManagerGroup = ""     # Grupp, kes haldab ja näeb kalendrit

# 2. Kontrolli, kas postkast juba olemas
if (Get-Recipient -Identity $RoomAlias -ErrorAction SilentlyContinue) {
    Write-Host "Postkast $RoomAlias juba olemas. Skript lõpetatakse..." -ForegroundColor Red
    return
}

# 3. Loo ruumipostkast
New-Mailbox -Room -Name $RoomName -Alias $RoomAlias -DisplayName $RoomName
Set-Mailbox -Identity $RoomAlias -PrimarySmtpAddress $RoomEmail

# 4. Määra ainult juhtkonna grupile õigused
# Lisa Full Access juhtkonna grupile
Add-MailboxPermission -Identity $RoomAlias -User $ManagerGroup -AccessRights FullAccess -InheritanceType All

# Eemalda kõik teised Full Access õigused
Get-MailboxPermission -Identity $RoomAlias | Where-Object {
    ($_.User -notlike "NT AUTHORITY\SELF") -and ($_.User -notlike $ManagerGroup)
} | ForEach-Object {
    Remove-MailboxPermission -Identity $RoomAlias -User $_.User -AccessRights FullAccess -Confirm:$false
}

# Lisa SendAs õigused juhtkonna grupile
Add-RecipientPermission -Identity $RoomAlias -Trustee $ManagerGroup -AccessRights SendAs

# Eemalda SendAs õigused teistelt
Get-RecipientPermission -Identity $RoomAlias | Where-Object {
    ($_.Trustee -notlike $ManagerGroup)
} | ForEach-Object {
    Remove-RecipientPermission -Identity $RoomAlias -Trustee $_.Trustee -AccessRights SendAs -Confirm:$false
}

# 5. Konfigureeri automaatne broneerimine ainult juhtkonna grupile
Set-CalendarProcessing -Identity $RoomAlias `
    -AutomateProcessing AutoAccept `
    -AllBookInPolicy $false `
    -BookInPolicy $ManagerGroup `
    -RequestInPolicy $ManagerGroup `
    -AllRequestOutOfPolicy $false `
    -AllowConflicts $false `
    -BookingWindowInDay 180

# 6. Kontroll
Write-Host "Privaatne ruumikalender loodud edukalt!" -ForegroundColor Green
Get-Mailbox -Identity $RoomAlias | Format-List Name,Alias,PrimarySmtpAddress,RecipientTypeDetails,HiddenFromAddressListsEnabled
Get-MailboxPermission -Identity $RoomAlias
Get-RecipientPermission -Identity $RoomAlias
