# PowerShell.Public
Public facing Sci-Net powershell modules

Connect-SmartAzAccount
  Synopsis
    Will attempt to login to Azure without manual user intervention
  Description
    Connect-SmartAzAccount will attempt to login the user without interaction if the current user is an Azure AD identity and the device is joined through a Hybrid Azure AD Join.
  Parameters
    TenantID : String : Preferred tenancy ID to connect in Azure
    SuppressWarnings : Switch : Include flag to hide warnings thrown when connecting to an Azure account with access to multiple subscriptions
    $DebugMode  : Switch : Include flag to run in debug mode and write access token generated to the host
    InstallPrerequisites : Switch : Include to automatically install pre-requsite modules
  Example
    Disconnect-AzAccount | Out-Null
    Start-sleep -Seconds 10
    Connect-SmartAzAccount 
    (Get-AzSubscription).Name
    
Connect-AzAccountManual
  Synopsis
    Will launch Connect-AzAccount with support for suppressing warnings
  Description
    Connect-AzAccountManual will attempt to login the user by prompting them for Azure AD credentials
  Parameters
    TenantID : String : Preferred tenancy ID to connect in Azure
    SuppressWarnings : Switch : Include flag to hide warnings thrown when connecting to an Azure account with access to multiple subscriptions
    InstallPrerequisites : Switch : Include to automatically install pre-requsite modules
    
Get-AzAccountToken
  Synopsis
    Will obtain a bearer token for the given user by their email address
  Description
    Obtains a code value from the Authorize OAuth API, captures the returned parameters through Internet Explorer and a listener on an empheral port, then uses the token OAuth API to obtain a bearer token and user ID from the obtained code
    Inspired by workflow described here: https://stackoverflow.com/questions/66328011/how-prevent-pick-an-account-interactive-workflow-when-using-connect-azaccount
  Parameters
    $EmailAddress : String : Email address of Azure user to obtain bearer token for
    $DebugMode  : Switch : Include flag to run in debug mode and write access token generated to the host
