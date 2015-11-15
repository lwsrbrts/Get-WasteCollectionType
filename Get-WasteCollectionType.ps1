#requires -version 4

$ScriptPath = Split-Path $Script:MyInvocation.MyCommand.Path
Import-Module ".\Send-WasteCollectionInvitations.psm1"

#############
# VARIABLES #
#############

# Variables to use when sending the email.
$Sender = "Waste Collections <bounce@lewisroberts.com>"

# For Send-WasteCollectionInvitations, a hashtable (an associative array) is required.
# "Name" = display name, "Value" = email address
$To = @{
    "Lewis Roberts" = "lewis@lewisroberts.com";
    "Kym Jones" = "miss.kr.jones@gmail.com";
}
# Can be accessed as $To.`Lewis Roberts`

$Subject = "Waste Collection"
$Server = "smtp.gmail.com"
$Port = "587"

# To create an encrypted credential used by this script, use this
# on the machine and logged on as the user that will run the script:
#Get-Credential | Export-Clixml -Path ($env:COMPUTERNAME+"bounce-lewisroberts.xml")
$EmailCredentials = Import-Clixml -Path ($env:COMPUTERNAME+"bounce-lewisroberts.xml")

# The house details
$Postcode = "CW47BN"
$HouseNo = "9"

##############
# PROCESSING #
##############

# Set up a session with the page. ASP.
$r = Invoke-WebRequest -Uri 'http://online.cheshireeast.gov.uk/MyCollectionDay/' -SessionVariable RequestForm

# Set the form's required information
$r.Forms.fields['ctl00$ContentPlaceHolder1$txtPostcode'] = $Postcode
$r.Forms.fields['ctl00$ContentPlaceHolder1$txtHouseNameNum'] = $HouseNo
$r.Forms.fields['ctl00$ContentPlaceHolder1$btnSearch'] = 'Search'

# POST the form and get the results from the page
$response = Invoke-WebRequest -Uri 'http://online.cheshireeast.gov.uk/MyCollectionDay/' `
                              -WebSession $RequestForm `
                              -Method POST `
                              -Body $r.forms.Fields `
                              -ContentType 'application/x-www-form-urlencoded'

# Extract the collectionDates table from the response
$CollectionDatesTable = ($response.ParsedHtml.getElementsByTagName('table') | Where-Object {$_.className -eq 'collectionDates'}).outerHTML

# Compact the string array to a single string.
$CollectionDatesTable = $CollectionDatesTable.Replace("`r`n","")

# Clean up the table and convert to CSV
# All customised depending on the table structure
# Will break if they change the HTML.
$CurTable = $CollectionDatesTable `
                -replace "<THEAD>.*</THEAD>" `
                -replace "</TR>","`r`n" `
                -replace "<TR .*`">" `
                -replace "<TABLE .*<TH>" `
                -replace "<TD colSpan=5></TD>`r`n" `
                -replace "</?(TR|TH|B)>" `
                -replace "</TD><TD>","," `
                -replace "</?T(D|R)>" `
                -replace "<(TBODY|THEAD)>" `
                -replace "</(TBODY|THEAD)>" `
                -replace "</?(TABLE).*>" 
$collectionDates = $CurTable | ConvertFrom-Csv -Header Bin,Day,Date,Type | Select-Object Date,Type

# Clean up the table to make sense - convert dates to datetime and change verbose type of collection to "bin colour"
For ($i = 0; $i -le ($collectionDates.Count-1); $i++) {
    $collectionDates[$i].Date = Get-Date $collectionDates[$i].Date
    Switch -Regex ($collectionDates[$i].Type) {
        "Other" {$collectionDates[$i].Type = "Black Bin"}
        "Recycling" {$collectionDates[$i].Type = "Silver Bin"}
        "Garden" {$collectionDates[$i].Type = "Brown Bin"}
    }
}

# Select only the dates that occur in the next 6 days (ie. just this week's collections)
[Array]$Fragment = $collectionDates | Where {$_.Date -le (Get-Date).AddDays(6)} | Select-Object Date,Type

# Now change that date to a format we understand and put the collection bin colour in the subject line.
For ($i=0;$i -le $Fragment.GetUpperBound(0);$i++) {
    $Fragment[$i].Date = Get-Date $Fragment[$i].Date -Format "dddd dd/MM/yyyy"
    If ($i % 2 -eq 1) {
        $Subject += ' & '+$Fragment[$i].Type
    }
    Else {$Subject += ' - '+$Fragment[$i].Type}
}

# Convert the table to HTML so we can insert it in to our email.
$Fragment = $Fragment | ConvertTo-Html -Fragment

# Get an HTML template for our email.
$HtmlTemplate = Get-Content "$ScriptPath\contact_template.htm" -Raw # -Raw parameter gets the file instead of loading each line in an Array.
    
# Change the template's placeholders with useful information.
$HtmlData = $HtmlTemplate.Replace('%%fragment%%', $Fragment)

# Send the email using the template.
Send-WasteCollectionInvitations -To $To `
                                -Server $Server `
                                -Port $Port `
                                -From $Sender `
                                -Credentials $EmailCredentials `
                                -Subject $Subject `
                                -Body $HtmlData
