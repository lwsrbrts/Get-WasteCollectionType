#requires -version 4

# Invoke-WebRequest changes significantly in PowerShell Core 6. ParsedHtml no longer exists!

$ScriptPath = Split-Path $Script:MyInvocation.MyCommand.Path
Import-Module -Name "$ScriptPath\Send-WasteCollectionInvitations.psm1"

#############
# VARIABLES #
#############

# Variables to use when sending the invitations emails.
$Sender = "Waste Collections <waste_collections@lewisroberts.com>"

# For Send-WasteCollectionInvitations, a hashtable (an associative array) is required.
# "Name" = display name, "Value" = email address
$To = @{
    "Lewis Roberts" = "lewis@lewisroberts.com";
    #"Joe Bloggs" = "joe.bloggs@lewisroberts.com";
}
# Can be accessed as $To.`Lewis Roberts` if you like.

# I always use Gmail for the solution so I'm not providing control
# over the use of EnableSSL, it is ALWAYS enabled.
# Without MFA being enabled ont eh account, no matter how many times
# I confirmed that "enable lesser security" was On for the account
# Gmail's SMTP would always fail to authorise. I had to enable MFA
# enable lesserver security and set up an application password. Grr.
$Subject = "Waste Collection"
$Server = "smtp.gmail.com"
$Port = "587"

# To create an encrypted credential used by this script, use this
# on the machine and logged on as the user that will run the script:
# Helps avoid having to leave plain text credentials in the script.
#Get-Credential | Export-Clixml -Path ($env:COMPUTERNAME+"bounce-lewisroberts.xml")
$EmailCredentials = Import-Clixml -Path ($env:COMPUTERNAME + "bounce-lewisroberts.xml")

# The house details to submit to the site (where the results should be based)
$Postcode = "CW48FE"
$HouseNo = "22"

##############
# PROCESSING #
##############

$search = Invoke-WebRequest -UseBasicParsing -Uri "http://online.cheshireeast.gov.uk/MyCollectionDay/SearchByAjax/Search?postcode=$($Postcode)&propertyname=$($HouseNo)" -Method Get

$searchResult = $search.Links | Where-Object {$_.class -match "get-job-details"}

$wasteCollections = Invoke-WebRequest -UseBasicParsing -Uri "http://online.cheshireeast.gov.uk/MyCollectionDay/SearchByAjax/GetBartecJobList?uprn=$($searchResult.'data-uprn')&onelineaddress=$([uri]::EscapeUriString(($searchResult.'data-onelineaddress')))"

$CollectionDatesTable = $wasteCollections.Content | Select-String -Pattern '<label for=(.|\n|\r)+?<\/label>' -AllMatches | ForEach-Object {$_.Matches} | ForEach-Object {$_.Value}

$CurTable = $CollectionDatesTable `
    -replace '<label for="(.)*">' `
    -replace '</label>', ',' `
    -join ''

$Results = $CurTable | Select-String -Pattern '(([^,]*,){3}\s*)' -AllMatches | ForEach-Object {$_.Matches} | ForEach-Object {$_.Value}

$collectionDates = ConvertFrom-Csv -Header Day, Date, Type -InputObject $Results

# Clean up the table to make sense - convert dates to datetime objects so we can do calculations
# on them and change verbose type of collection to "bin colour"
For ($i = 0; $i -le ($collectionDates.Count - 1); $i++) {
    $collectionDates[$i].Date = Get-Date $collectionDates[$i].Date
    Switch -Regex ($collectionDates[$i].Type) {
        "General" {$collectionDates[$i].Type = "Black Bin"}
        "Recycling" {$collectionDates[$i].Type = "Silver Bin"}
        "Garden" {$collectionDates[$i].Type = "Brown Bin"}
    }
}

# Select only the dates that occur in the next 6 days (ie. just this week's collections)
[Array]$Fragment = $collectionDates | Where-Object {$_.Date -le (Get-Date).AddDays(6)} | Select-Object Date, Type

# To set a reminder (invite) date on the evening before the collection, we need to know
# the date of the collection. This is sent to the Send-WasteCollectionInvitations function
# unedited. It is converted to "the night before" within that function.
[datetime]$EventDate = $collectionDates | Where-Object {$_.Date -le (Get-Date).AddDays(6)} | Group-Object -Property Date -NoElement | Select-Object -First 1 -Property Name -ExpandProperty Name | Get-Date

# Now change that date to a format we understand and put the collection bin colour in the subject line.
# I didn't do this earlier because I needed to select dates by calculation.
For ($i = 0; $i -le $Fragment.GetUpperBound(0); $i++) {
    $Fragment[$i].Date = Get-Date $Fragment[$i].Date -Format "dddd dd/MM/yyyy"
    If ($i % 2 -eq 1) {
        $Subject += ' & ' + $Fragment[$i].Type
    }
    Else {$Subject += ' - ' + $Fragment[$i].Type}
}

# Convert the table to HTML so we can insert it in to our email.
$Fragment = $Fragment | ConvertTo-Html -Fragment

# Get an HTML template for our email.
$HtmlTemplate = Get-Content "$ScriptPath\contact_template.htm" -Raw # -Raw parameter gets the file instead of loading each line in an Array.
    
# Change the template's placeholders with the tabular information.
$HtmlData = $HtmlTemplate.Replace('%%fragment%%', $Fragment)

# Send the email plus invites! using the template.
Send-WasteCollectionInvitations -To $To `
                                -Server $Server `
                                -Port $Port `
                                -From $Sender `
                                -Credentials $EmailCredentials `
                                -Subject $Subject `
                                -EventDate $EventDate `
                                -Body $HtmlData
