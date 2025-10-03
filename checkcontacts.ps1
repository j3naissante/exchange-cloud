# Define the path to the TXT file containing the list of email addresses
$txtFilePath = "path"

# Import the list of email addresses from the TXT file
# Assuming one email address per line
$emailList = Get-Content -Path $txtFilePath

# Initialize an array to store the results
$results = @()

# Connect to Exchange Online (if not already connected)
# You need the ExchangeOnlineManagement module installed
# Install-Module ExchangeOnlineManagement -Scope CurrentUser


# Loop through each email address in the list
foreach ($email in $emailList) {
    Write-Output "Checking email: $email"
    
    # Check if the mail contact exists for the current email address
    $contact = Get-MailContact -Identity $email -ErrorAction SilentlyContinue
    
    if ($contact) {
        Write-Output "Found: $email"
        $exists = "Exists"
    } else {
        Write-Output "Not found: $email"
        $exists = "Does not exist"
    }
    
    # Add the result to the results array as a custom object
    $results += [PSCustomObject]@{
        EmailAddress = $email
        Status       = $exists
    }
}


# Display the results in a grid view
$results | Out-GridView

# Optionally, export results to CSV
$results | Export-Csv -Path "C:\path\to\EmailCheckResults.csv" -NoTypeInformation
