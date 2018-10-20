# POST method: $req
$requestBody = Get-Content $req -Raw | ConvertFrom-Json
$Postcode = [System.Web.HttpUtility]::HtmlEncode($requestBody.postcode)
$HouseNo = [System.Web.HttpUtility]::HtmlEncode($requestBody.houseno)

# Get all the variables and their names/values. Handy for debugging.
<#
$resp = @{}
get-variable | ? { $_.Value -is [string] } | % { $resp["$($_.Name)"] = $_.Value }
gci env:appsetting* | % { $resp["ENV:$($_.Name)"] = $_.Value }
$jsonResp = $resp | ConvertTo-Json -Compress
Out-File -Encoding Ascii -FilePath $res -inputObject $jsonResp
#>

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

Out-File -Encoding Ascii -FilePath $res -inputObject ($collectionDates | ConvertTo-Json -Depth 5)
