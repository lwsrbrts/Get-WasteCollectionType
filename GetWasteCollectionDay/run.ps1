using namespace System.Net

# POST method: $req
param($req)

#$requestBody = Get-Content $req -Raw | ConvertFrom-Json
$Postcode = [System.Web.HttpUtility]::HtmlEncode($req.Body.Postcode)
$HouseNo = [System.Web.HttpUtility]::HtmlEncode($req.Body.HouseNo)

if ($null -eq $req.Body.Postcode -or $null -eq $req.Body.HouseNo) {
    
    Push-OutputBinding -Name res -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = '{"error": "Missing postcode or houseno. Both are required."}'
    })

    exit
}

# Get all the variables and their names/values. Handy for debugging.
<#
$resp = @{}
get-variable | ? { $_.Value -is [string] } | % { $resp["$($_.Name)"] = $_.Value }
gci env:appsetting* | % { $resp["ENV:$($_.Name)"] = $_.Value }
$jsonResp = $resp | ConvertTo-Json -Compress
Push-OutputBinding -Name res -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $jsonResp
})
exit
#>

Try {
    $search = Invoke-WebRequest -UseBasicParsing -Uri "http://online.cheshireeast.gov.uk/MyCollectionDay/SearchByAjax/Search?postcode=$($Postcode)&propertyname=$($HouseNo)" -Method Get -ErrorAction Stop
}
Catch {
    Push-OutputBinding -Name res -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = '{"error": "The search using the details provided did not complete correctly. This may indicate that the address provided is incorrect or that the service/website is unavailable."}'
    })

    exit
}

$searchResult = $search.Links | Where-Object { $_.class -match "get-job-details" }
If (-not($searchResult)) {
    
    Push-OutputBinding -Name res -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = '{"error": "No results were returned or an error occurred. This may indicate that the service or website is unavailable."}'
    })
    
    exit
}

Try {
    $wasteCollections = Invoke-WebRequest -UseBasicParsing -Uri "http://online.cheshireeast.gov.uk/MyCollectionDay/SearchByAjax/GetBartecJobList?uprn=$($searchResult.'data-uprn')&onelineaddress=$([uri]::EscapeUriString(($searchResult.'data-onelineaddress')))" -ErrorAction Stop
}
Catch {
    Push-OutputBinding -Name res -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = '{"error": "The search for the collection dates for the property did not complete. This may indicate that the service or website is unavailable."}'
    })

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
<#
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
#>

$collectionSchedule = @{}

foreach ($Entry in $collectionDates) {
    Try {
        $CollectionDate = [datetime]::ParseExact($Entry.Date, "dd/MM/yyyy", [cultureinfo]::InvariantCulture)
    }
    Catch {
        Continue # Skip that entry as we can't decipher the date.
    }

    # Determine collection type
    $Type = switch -Regex ($Entry.Type) {
        "General" { "Black Bin" }
        "Recycling" { "Silver Bin" }
        "Garden" { "Brown Bin" }
        default { $_ } # Include type as is if not matched
    }

    # Check if the date already has an entry in the hashtable
    if ($collectionSchedule.ContainsKey($CollectionDate)) {
        # Append the type if it's not already listed for that date
        if ($collectionSchedule[$CollectionDate].Type -notcontains $Type) {
            $collectionSchedule[$CollectionDate].Type += "& $Type"
        }
    }
    else {
        # Create a new entry for the date with the type
        $collectionSchedule[$CollectionDate] = [PSCustomObject]@{
            Day  = $Entry.Day
            Date = $CollectionDate
            Type = $Type
        }
    }
}

$collectionSchedule.Values | Sort-Object Date| ConvertTo-Json -Depth 5


# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name res -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    #Body = ($cleanCollectionSchedule | ConvertTo-Json -Depth 5)
    Body = $collectionSchedule.Values | Sort-Object Date | ConvertTo-Json -Depth 5
})