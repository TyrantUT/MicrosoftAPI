# Microsoft APIs

Collection of Microsoft APIs not generally available within Powershell modules

## API Reference

#### O365.OfficeInstalls
Returns a table of users and machines where Office 365 Products have been installed / activated.

```powershell
   Get-Office365Installs | Export-Csv -Path "file.csv" -NoTypeInformation
```
