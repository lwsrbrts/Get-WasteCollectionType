# GetWasteCollectionDay (Azure Function)
This is an Azure Function (using runtime version 1.0) and can be deployed from here directly to Azure. To do this for yourself, you'll need to clone the repository to your own account of course then connect Azure to that.

# Why is it written in PowerShell?
Because that's the language I know best. Yes, I'm sure it could be built quicker, run faster, do all sorts of weird and wonderful things if written in your favourite language but the point is that the end result is a set of collection dates presented in JSON for waste collections in Cheshire East.

# Can I use your Azure Function?
If you want to but know that it's liable to change or be removed at any moment. Obviously you'll have to live in Cheshire East but use something like this (converted in to the language you like the most)

## PowerShell

```powershell
Invoke-RestMethod -Method Post -Uri 'https://transishun.azurewebsites.net/api/GetWasteCollectionDay?code=RB4UpAld9G3xcZZyUPgmS7BQJ2gAPBtWsvGyCQ4T64YRs1AcyFqN/Q==' -Body '{"postcode": "CW11AA", "houseno": "48"}'
```

## Curl

```
curl --header "Content-Type: application/json" --request POST --data '{"postcode": "CW11AA", "houseno": "48"}' https://transishun.azurewebsites.net/api/GetWasteCollectionDay?code=RB4UpAld9G3xcZZyUPgmS7BQJ2gAPBtWsvGyCQ4T64YRs1AcyFqN/Q==
```