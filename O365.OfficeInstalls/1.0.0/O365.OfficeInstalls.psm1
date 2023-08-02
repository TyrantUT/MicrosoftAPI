function Get-Office365Token {
    <#
    .SYNOPSIS
    Obtains the necessary Authentication tokens required to make API calls to the Microsoft API admin.microsoft.com/admin/api
 
    .DESCRIPTION
    Obtains the necessary Authentication tokens required to make API calls to the Microsoft API admin.microsoft.com/admin/api
 
    .EXAMPLE
    Get-Office365Token
 
    .NOTES
    Requires the Powershell module 'Az'
    Function is called from within Get-Office365Installs
    #>

    param(
        [int] $ExpiresIn = 3600,
        [int] $ExpiresTimeout = 30
    )
 
    if ($global:O365AuthorizationCache) {
        if ($global:O365AuthorizationCache.TokenExpiration -gt [datetime]::UtcNow) {
            Write-Host "[$([datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss tt'))] Connecting to Office 365 Admin API using Cached Token"
            return
        }
    }

    Write-Host "[$([datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss tt'))] Connecting to Office 365 Admin API"
    $null = Connect-AzAccount -WarningAction SilentlyContinue

    Write-Host "[$([datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss tt'))] Getting Microsoft API Token"
    $O365AuthorizationToken = (Get-AzAccessToken -ResourceUrl "https://admin.microsoft.com").Token

    if ($null -ne $O365AuthorizationToken) {
        $global:O365AuthorizationCache = [ordered] @{
            'TokenExpiration'           = ([datetime]::UtcNow).AddSeconds($ExpiresIn - $ExpiresTimeout)
            'Headers'                   = [ordered] @{
                "Content-Type"              = "application/json; charset=UTF-8"
                "Authorization"             = "Bearer $($O365AuthorizationToken)"
                'X-Requested-With'          = 'XMLHttpRequest'
                'x-ms-client-request-id'    = [guid]::NewGuid()
                'x-ms-correlation-id'       = [guid]::NewGuid()
            }
        }
    }    
}

function Get-Office365Users {
    <#
    .SYNOPSIS
    Returns all Office 365 Users with an Email address
 
    .DESCRIPTION
    Returns all Office 365 Users with an Email address
 
    .EXAMPLE
    Get-Office365Users
 
    .NOTES
    Function is called from Get-O365Installs
    #>

    Write-Host "[$([datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss tt'))] Getting list of Office 365 Users"
    $Users = Get-AzAdUser | Where-Object {$null -ne $_.Mail} | Select-Object DisplayName, Mail, Id

    $UserCount = $Users.Count
    Write-Host "[$([datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss tt'))] Got $($UserCount) Users"
    return $Users
}

function Get-UserOfficeInstalls {
    <#
    .SYNOPSIS
    Connects to the Microsoft Admin API using the officeInstalls endpoint not generrally available within Powershell modules
 
    .DESCRIPTION
    Connects to the Microsoft Admin API using the officeInstalls endpoint not generrally available within Powershell modules

    .PARAMETER User
    [System.Object] Containing Id, DisplayName, Mail
 
    .EXAMPLE
    Get-UserOfficeInstalls
 
    .NOTES
    Function is called from Get-Office365Installs
    #>

    param(
        [System.Object] $User
    )

    $UserId = $User.Id
    $UserDisplayName = $User.DisplayName
    $UserMail = $User.Mail

    if ($null -eq $UserId) {
        return
    }
    
    $Params = @{
        Method      = 'GET'
        ContentType = "application/json; charset=UTF-8"
        Headers     = $global:O365AuthorizationCache.Headers
        Uri         = "https://admin.microsoft.com/admin/api/users/$UserId/officeInstalls"
    }

    $QueryOutput = Invoke-RestMethod @Params
    $Results = $QueryOutput | Select-Object SoftwareMachineDetails -ExpandProperty SoftwareMachineDetails

    $CompleteRecord = @()
    foreach ($Result in $Results) {
        $MachineData = $Result.MachineDetails.Machines

        if ($null -ne $MachineData) {
            foreach ($MachineDetails in $MachineData) {
                foreach ($md in $MachineDetails) {
                    $MachineRecord = [PSCustomObject]@{
                        User = $UserDisplayName
                        UserMail = $UserMail
                        MachineName = $md.MachineName
                        MachineOs = $md.MachineOs
                        LastLicenseRequestedDate = [datetime]$md.LastLicenseRequestedDate
                    }
                    $CompleteRecord += $MachineRecord
                }
            }
        }
    }

    return $CompleteRecord
}

function Get-Office365Installs {
    <#
    .SYNOPSIS
    Wrapper for the Office 365 Office Installs API
 
    .DESCRIPTION
    Wrapper for the Office 365 Office Installs API
 
    .EXAMPLE
    Get-Office365Installs | Export-Csv -Path "file.csv" -NoTypeInformation
 
    .NOTES
    Notes
    #>

    $null = Get-Office365Token
    $Users = Get-Office365Users

    [int] $PercentComplete = 0
    $CurrentItem = 0
    $ResultsTable = @()

    Write-Host "[$([datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss tt'))] Getting Office Installs"
    foreach ($User in $Users) {
        Write-Progress -Activity "Office Installs for $UserDisplayName" -Status "$PercentComplete% Complete:" -PercentComplete $PercentComplete
        
        $UserDisplayName = $User.DisplayName
        $UserOfficeInstalls = Get-UserOfficeInstalls -User $User

        if ($UserOfficeInstalls) {
            $ResultsTable += $UserOfficeInstalls
        }

        $PercentComplete =  (($CurrentItem++ / $Users.Count) * 100)        
    }

    Write-Host "[$([datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss tt'))] Disconnecting from O365 Admin API"
    $null = Disconnect-AzAccount
     
    return $ResultsTable
}

Export-ModuleMember -Function Get-Office365Installs
