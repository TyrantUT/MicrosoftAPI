# Microsoft APIs

Collection of Microsoft APIs not generally available within Powershell modules

## API Reference

#### O365.OfficeInstalls
Returns a table of users and machines where Office has been activated on.

```powershell
   Get-Office365Installs | Export-Csv -Path "file.csv" -NoTypeInformation
```
