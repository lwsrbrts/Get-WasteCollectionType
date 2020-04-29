# POST method: $req


$requestBody = Get-Content $req -Raw | ConvertFrom-Json
$Postcode = [System.Web.HttpUtility]::HtmlEncode($requestBody.postcode)
$HouseNo = [System.Web.HttpUtility]::HtmlEncode($requestBody.houseno)
if ($requestBody.postcode -eq $null -or $requestBody.houseno -eq $null) {
    Out-File -Encoding Ascii -FilePath $res -inputObject '{"error": "Missing postcode or houseno. Both are required."}'
}

# Get all the variables and their names/values. Handy for debugging.
<#
$resp = @{}
get-variable | ? { $_.Value -is [string] } | % { $resp["$($_.Name)"] = $_.Value }
gci env:appsetting* | % { $resp["ENV:$($_.Name)"] = $_.Value }
$jsonResp = $resp | ConvertTo-Json -Compress
Out-File -Encoding Ascii -FilePath $res -inputObject $jsonResp
#>
Try {
    $search = Invoke-WebRequest -UseBasicParsing -Uri "http://online.cheshireeast.gov.uk/MyCollectionDay/SearchByAjax/Search?postcode=$($Postcode)&propertyname=$($HouseNo)" -Method Get -ErrorAction Stop
}
Catch {
    Out-File -Encoding Ascii -FilePath $res -inputObject '{"error": "The search using the details provided did not complete correctly. This may indicate that the address provided is incorrect or that the service/website is unavailable."}'
    exit
}

$searchResult = $search.Links | Where-Object { $_.class -match "get-job-details" }
If (-not($searchResult)) {
    Out-File -Encoding Ascii -FilePath $res -inputObject '{"error": "No results were returned or an error occurred. This may indicate that the service or website is unavailable."}'
    exit
}

Try {
    $wasteCollections = Invoke-WebRequest -UseBasicParsing -Uri "http://online.cheshireeast.gov.uk/MyCollectionDay/SearchByAjax/GetBartecJobList?uprn=$($searchResult.'data-uprn')&onelineaddress=$([uri]::EscapeUriString(($searchResult.'data-onelineaddress')))" -ErrorAction Stop
}
Catch {
    Out-File -Encoding Ascii -FilePath $res -inputObject '{"error": "The search for the collection dates for the property did not complete. This may indicate that the service or website is unavailable."}'
    exit
}

$CollectionDatesTable = $wasteCollections.Content | Select-String -Pattern '<label for=(.|\n|\r)+?<\/label>' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }

$CurTable = $CollectionDatesTable `
    -replace '<label for="(.)*">' `
    -replace '</label>', ',' `
    -join ''

$Results = $CurTable | Select-String -Pattern '(([^,]*,){3}\s*)' -AllMatches | ForEach-Object { $_.Matches } | ForEach-Object { $_.Value }

$collectionDates = ConvertFrom-Csv -Header Day, Date, Type -InputObject $Results

# Clean up the table to make sense - convert dates to datetime objects so we can do calculations
# on them and change verbose type of collection to "bin colour"
    
# If the "WEBSITE_TIME_ZONE" is set to a a region that implements daylight saving, dates change due to midnight being the default
# time when a date is converted.
# Still thinking about how this should be handled.

$cleanCollectionSchedule = foreach ($Entry in $collectionDates) {

    Try {
        $CollectionDate = [datetime]::ParseExact($Entry.Date, "dd/MM/yyyy", [cultureinfo]::InvariantCulture)
    }
    Catch {
        Continue # Skip that entry as we can't decipher the date.
    }

    Switch -Regex ($Entry.Type) {
        "General" { $Type = "Black Bin" }
        "Recycling" { $Type = "Silver Bin" }
        "Garden" { $Type = "Brown Bin" }
    }

    $Property = [ordered]@{
        Day  = $Entry.Day
        Date = $CollectionDate
        Type = $Type

    }

    New-Object -TypeName PSObject -Property $Property
}

Out-File -Encoding Ascii -FilePath $res -inputObject ($cleanCollectionSchedule | ConvertTo-Json -Depth 5)
