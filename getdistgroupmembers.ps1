Get-DistributionGroupMember -Identity "your@domain.com" | Select Name,PrimarySmtpAddress| Export-Csv "$env:TEMP\DL_Members.csv" -NoTypeInformation -Encoding UTF8
