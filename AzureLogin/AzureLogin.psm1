## Sci-Net.PowerShell.AzureLogin
## Allows automated login to Azure through Powershell, without the need to perform MFA if the account is used in a web session
## Adds following functons:
##   Connect-SmartAzAccount
##   Connect-AzAccountManual
##   Get-AzAccountToken

## Version 1.1.0

## Developed by: 
## Nathan Moore on behalf of Sci-Net Business Solutions
## 10-Mar-23

<#
.Synopsis
Will attempt to login to Azure without manual user intervention
.Description
Connect-SmartAzAccount will attempt to login the user without interaction if the current user is an Azure AD identity and the device is joined through a Hybrid Azure AD Join.
.Example
Disconnect-AzAccount | Out-Null
Start-sleep -Seconds 10
Connect-SmartAzAccount 
(Get-AzSubscription).Name
#>
Function Connect-SmartAzAccount {
param(
    #Preferred tenancy ID to connect in Azure
    [Parameter(Mandatory = $False)] [string] $TenantID,
    #Include flag to hide warnings thrown when connecting to an Azure account with access to multiple subscriptions
    [switch] $SuppressWarnings,
    #Include flag to run in debug mode and write access token generated to the host
    [switch] $DebugMode,
    #Include to automatically install pre-requsite modules
    [switch] $InstallPrerequisites,
    #Include to have prompt confirming account to attempt login with
    [switch] $ConfirmAccount
)

    Check-AzAccountsPreRequisite -InstallPrerequisites $InstallPrerequisites -DebugMode $DebugMode
       
    $Emailaddress = whoami.exe /upn
    try {
        if ($Emailaddress -match ".*@.*") {
            if ($ConfirmAccount) {
                $AttemptLogin = Switch([System.Windows.MessageBox]::Show(("Would you like to try logging into Azure as $EmailAddress ?"),"Azure Account Selection",[System.Windows.MessageBoxButton]::YesNo,[System.Windows.MessageBoxImage]::Question)) { 'Yes' {$true} 'No' {$False} }
            } else {
                $AttemptLogin = $true
            }
            if ($AttemptLogin) {
                if ($debugMode) {
                    if ($InstallPrerequisites) {
                        $Token = Get-AzAccountToken -EmailAddress $Emailaddress -DebugMode -InstallPrerequisites
                    } else {
                        $Token = Get-AzAccountToken -EmailAddress $Emailaddress -DebugMode
                    }
                    Write-Host ("Obtained Token:" + ($Token | Out-String))
                } else {
                    if ($InstallPrerequisites) {
                        $Token = Get-AzAccountToken -EmailAddress $Emailaddress -ErrorAction Stop -InstallPrerequisites
                    } else {
                        $Token = Get-AzAccountToken -EmailAddress $Emailaddress -ErrorAction Stop
                    }
                }
                
                if ($SuppressWarnings) {
                    if ($TenantID = $null) {
                        connect-azaccount -AccessToken $Token.access_token -AccountId $Token.ID_Token -WarningAction SilentlyContinue -TenantId $TenantID -ErrorAction Stop
                    } else {
                        connect-azaccount -AccessToken $Token.access_token -AccountId $Token.ID_Token -WarningAction SilentlyContinue -ErrorAction Stop
                    }
                } else {
                    if ($TenantID = $null) {
                        connect-azaccount -AccessToken $Token.access_token -AccountId $Token.ID_Token -TenantId $TenantID -ErrorAction Stop
                    } else {
                        connect-azaccount -AccessToken $Token.access_token -AccountId $Token.ID_Token -ErrorAction Stop
                    }
                }
            } else {
                Write-Host "User elected to login manually, please login:"
                if ($DebugMode) {
                    Connect-AzAccountManual -TenantID $TenantID -SuppressWarnings $SuppressWarnings -DebugMode
                } else {
                    Connect-AzAccountManual -TenantID $TenantID -SuppressWarnings $SuppressWarnings
                }
            }
        } else {
            Write-Host "No valid account found, please login:"
            if ($DebugMode) {
                Connect-AzAccountManual -TenantID $TenantID -SuppressWarnings $SuppressWarnings -DebugMode
            } else {
                Connect-AzAccountManual -TenantID $TenantID -SuppressWarnings $SuppressWarnings
            }
        }
    } catch {
        Write-Host "Unable to login automatically, please login:"
        if ($DebugMode) {
            Write-host $_ -ForegroundColor Red
            Connect-AzAccountManual -TenantID $TenantID -SuppressWarnings $SuppressWarnings -DebugMode
        } else {
            Connect-AzAccountManual -TenantID $TenantID -SuppressWarnings $SuppressWarnings
        }
    }
}

<#
.Synopsis
Will launch Connect-AzAccount with support for suppressing warnings
.Description
Connect-AzAccountManual will attempt to login the user by prompting them for Azure AD credentials
#>
function Connect-AzAccountManual
{
param(
    #Preferred tenancy ID to connect in Azure
    [string] $TenantID,
    #Include flag to hide warnings thrown when connecting to an Azure account with access to multiple subscriptions
    [Boolean] $SuppressWarnings = $False,
    #Include to automatically install pre-requsite modules
    [switch] $InstallPrerequisites,
    #Include to run in debug mode
    [Switch] $DebugMode
)
    
    Check-AzAccountsPreRequisite -InstallPrerequisites $InstallPrerequisites -DebugMode $DebugMode

    if ($SuppressWarnings) {
        if ($TenantID -eq $null) {
            connect-azaccount -WarningAction SilentlyContinue -TenantId $TenantID
        } else {
            connect-azaccount -WarningAction SilentlyContinue
        }
    } else {
        if ($TenantID -eq $null) {
            connect-azaccount -TenantId $TenantID
        } else {
            connect-azaccount
        }
    }
}


<#
.Synopsis
Will obtain a bearer token for the given user by their email address
.Description
Obtains a code value from the Authorize OAuth API, captures the returned parameters through Internet Explorer and a listener on an empheral port, then uses the token OAuth API to obtain a bearer token and user ID from the obtained code
Inspired by workflow described here: https://stackoverflow.com/questions/66328011/how-prevent-pick-an-account-interactive-workflow-when-using-connect-azaccount
#>
Function Get-AzAccountToken {
Param(
    #Email address of Azure user to obtain bearer token for
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $EmailAddress,
    #Include flag to run in debug mode and write access token generated to the host
    [switch] $DebugMode,
    #Include to automatically install pre-requsite modules
    [switch] $InstallPrerequisites
)
    <# Get-AzAccountToken Function Body #>

    Check-ThreadJobPreRequisite -InstallPrerequisites $InstallPrerequisites

    if ($debugMode) {Write-Host "Begin Code Obtain"}
    $ResponsePort = Get-InactiveTcpPort -Start 1024 -End 65535
    if ($debugMode) {Write-Host "ChosenPort: $ResponsePort"}
    
    $URL = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?scope=https://management.core.windows.net//.default openid profile offline_access&response_type=code&client_id=1950a258-227b-4e31-a9cf-717495945fc2&redirect_uri=http://localhost:$ResponsePort/&x-client-SKU=MSAL.Desktop&x-client-Ver=4.21.0.0&x-client-CPU=x64&x-client-OS=Microsoft Windows NT 10.0.19042.0&login_hint=$EmailAddress"
    $ie = New-Object -COM 'InternetExplorer.Application'
    if ($DebugMode) {$ie.Visible = $true}
    $ie.Navigate($URL)
    try {
        $RawURL = Get-HttpQueryParametersSentToLocalhostAsync -Verbose -Port $ResponsePort -ErrorAction Stop
    } catch [System.ArgumentException] { 
        if ($DebugMode) {
            Write-Host "Timeout invalid"
            Write-Error $PSItem.toString()
        }
        Close-IETab($URL)
        throw $_
    } catch [System.TimeoutException] {
        if ($DebugMode) {
            Write-Host "Request Timed Out"
            Write-Error $PSItem.toString()
        }
        Close-IETab($URL)
        throw $_
    } catch {
        if ($DebugMode) {
            Write-Host "Unknown Error Encountered"
            Write-Error $PSItem.toString()
        }
        Close-IETab($URL)
        throw $_
    }
    $parameters = $RawURL -split "[?&]" -like "*=*" | foreach -Begin {$h = @{}} -Process {$h[($_ -split "=",2 | select -index 0)] = ($_ -split "=",2 | select -index 1)} -End {$h}
    if ($debugMode) {Write-Host ("Code Returned:" + ($parameters | out-String))}
    Close-IETab($RawURL)

    $body = "client_id=1950a258-227b-4e31-a9cf-717495945fc2&client_info=1&scope=https://management.core.windows.net//.default offline_access openid profile&grant_type=authorization_code&redirect_uri=http://localhost:$ResponsePort/&code=" + $parameters.code
    
    return ConvertFrom-Json (Invoke-WebRequest "https://login.microsoftonline.com/organizations/oauth2/v2.0/token" -Method POST -body $body).Content
}


<# Helper Functions #>

<#
.Synopsis
Ensure Az.Accounts module is installed
.Description
Will install Az.Accounts module if paramater is true, otherwise will throw error prompting user to install the module manually
#>
Function Check-AzAccountsPreRequisite
{
param (
    [Boolean] $InstallPrerequisites,
    [Boolean] $DebugMode
)
    Import-Module Az.Accounts -ErrorAction SilentlyContinue
    
    if ("Az.Accounts" -notin (Get-Module).Name) {
        if ($InstallPrerequisites) {
            Write-Host "Installing Az.Accounts"
            if ($DebugMode) {
                start-process -FilePath "powershell" -ArgumentList "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted `nInstall-Module az.accounts" -WindowStyle Normal -Verb RunAs -Wait
            } else {
                start-process -FilePath "powershell" -ArgumentList "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted `nInstall-Module az.accounts" -WindowStyle Hidden -Verb RunAs -Wait
            }
            import-module Az.Accounts
        } else {
            Throw "Missing Pre-Requsite Modules : Ensure Az.Accounts module is installed and available before trying again. Run:`nInstall-Module Az.Accounts"
        }
    }
}

<#
.Synopsis
Ensure ThreadJob module is installed
.Description
Will install ThreadJob module if paramater is true, otherwise will throw error prompting user to install the module manually
#>
Function Check-ThreadJobPreRequisite
{
param (
    [Boolean] $InstallPrerequisites,
    [Boolean] $DebugMode
)
    Import-Module ThreadJob -ErrorAction SilentlyContinue
    
    if ("ThreadJob" -notin (Get-Module).Name) {
        if ($InstallPrerequisites) {
            Write-Host "Installing ThreadJob"
            if ($DebugMode) {
                start-process -FilePath "powershell" -ArgumentList "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted `nInstall-Module ThreadJob" -WindowStyle Normal -Verb RunAs -Wait
            } else {
                start-process -FilePath "powershell" -ArgumentList "Set-PSRepository -Name PSGallery -InstallationPolicy Trusted `nInstall-Module ThreadJob" -WindowStyle Hidden -Verb RunAs -Wait
            }
            import-module ThreadJob
        } else {
            Throw "Missing Pre-Requsite Modules : Ensure ThreadJob module is installed and available before trying again. Run:`nInstall-Module ThreadJob"
        }
    }
}

<#
.Synopsis
Will find an inactive TCP port
Source: https://codeandkeep.com/PowerShell-Tcp-Port-Finder/
.Output
A randomly selected port that is currently not in use by any other process locally
#>
Function Get-InactiveTcpPort
{
    [CmdletBinding()]
    Param(
    #Smallest port number that can be used
    [Parameter(Position=0)]
    [uint16]$Start = 1024,

    #Largest port number that can be used
    [Parameter(Position=1)]
    [uint16]$End = 5000
    )
    $attempts = 100
    $counter = 0

    $activePorts = Get-ActiveTcpPort

    while ($counter -lt $attempts)
    {
    $counter++
    $port = Get-Random -Minimum ($Start -as [int]) -Maximum ($End -as [int])

    if ($port -notin $activePorts)
    {
        return $port
    }
    }
    $emsg = [string]::Format(
    'Unable to find available TCP Port. Range: {0}, Attempts: [{1}]',
    "[$Start - $End]",
    $attempts
    )
    throw $emsg
}


<#
.Synopsis
Will list active TCP ports
Source: https://codeandkeep.com/PowerShell-Tcp-Port-Finder/
.Output
A hashset of active ports
#>
Function Get-ActiveTcpPort
{
    # Use a hash set to avoid duplicates
    $portList = New-Object -TypeName Collections.Generic.HashSet[uint16]

    $properties = [Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()

    $listener = $properties.GetActiveTcpListeners()
    $active = $properties.GetActiveTcpConnections()

    foreach ($serverPort in $listener)
    {
    [void]$portList.Add($serverPort.Port)
    }
    foreach ($clientPort in $active)
    {
    [void]$portList.Add($clientPort.LocalEndPoint.Port)
    }

    return $portList
}

<#
.Synopsis
Launches HTTP listener on given port, captures parameters sent and returns them in a raw format
.Description
Launches HTTP listener on given port, captures parameters sent and returns them in a raw format
Source: https://goodworkaround.com/2019/12/20/quick-powershell-cmdlet-to-get-query-parameters-sent-to-localhost/
#>
Function Get-HttpQueryParametersSentToLocalhost
{
    [CmdletBinding()]
    Param
    (
        #Port to listen on
        [Parameter(Mandatory=$false,
                    ValueFromPipelineByPropertyName=$true,
                    Position=0)]
        [int] $Port = 8080,
 
        #HTML content to display once parameters have been captured
        [Parameter(Mandatory=$false,
                    ValueFromPipelineByPropertyName=$true,
                    Position=1)]
        [string] $Response = "Done"
    )
 
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$Port/")
    # Write-verbose "Waiting for request at http://localhost:$Port/"
    $listener.Start()
    $context = $listener.GetContext()
    $Content = [System.Text.Encoding]::UTF8.GetBytes($Response)
    $Context.Response.OutputStream.Write($Content, 0, $Content.Length)
    $Context.Response.Close()
    $listener.Dispose()
    return $context.Request.RawUrl
     
}

<#
.Synopsis
Launches HTTP listener on given port as an async job, captures parameters sent and returns them in a raw format
.Description
Launches HTTP listener on given port as an async job, captures parameters sent and returns them in a raw format, 
Based on: https://goodworkaround.com/2019/12/20/quick-powershell-cmdlet-to-get-query-parameters-sent-to-localhost/
#>
Function Get-HttpQueryParametersSentToLocalhostAsync
{
    [CmdletBinding()]
    Param
    (
        #Port to listen on
        [Parameter(Mandatory=$false,
                    ValueFromPipelineByPropertyName=$true,
                    Position=0)]
        [int] $Port = 8080,
 
        #HTML content to display once parameters have been captured
        [Parameter(Mandatory=$false,
                    ValueFromPipelineByPropertyName=$true,
                    Position=1)]
        [string] $Response = "Done",

        #Response Timeout - Will throw timeout error if exceeded
        [Parameter(Mandatory = $False,
                    ValueFromPipelineByPropertyName=$True,
                    Position=2)]
        [int] $Timeout = 10
    )
 
    if ($timeout -lt 1) { Throw [system.ArgumentException]::New("Provided value for Timeout is invalid. Timeout must be greater than 0") }
    
    
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$port/")
    # Write-verbose "Waiting for request at http://localhost:$Port/"
    $listener.Start()
    $Jobs = Start-ThreadJob -Argumentlist @($Port, $Response, $Listener) -ScriptBlock { 
        Param ( $port , $response , $Listener )
        
        $Context = $listener.GetContext() 
        $Content = [System.Text.Encoding]::UTF8.GetBytes($Response)
        $Context.Response.OutputStream.Write($Content, 0, $Content.Length)
        $Context.Response.Close()
        $listener.Dispose()
        $Listener.Close()
        return $context.Request.RawUrl
    }


    foreach ( $i in @(0..$Timeout) )
    {
        if ( $Jobs.State -eq "Completed" ) {
            Return Receive-Job $Jobs
        } else {
            Start-sleep -Seconds 1
        }
    }

    Write-Host "Automated Login Timed out"
    $Listener.Dispose()
    $Listener.Close()
    throw [System.TimeoutException]::New("Provided timeout for response exceeded")
}


<#
.Synopsis
Will scan for Internet Explorer tab with URL ending in parameter value, and close the tab
.Description
Will scan for Internet Explorer tab with URL ending in parameter value, and close the tab
Source: https://stackoverflow.com/questions/49575365/how-to-close-a-tab-in-existing-ie-window-using-powershell
#>
Function Close-IETab {
param(
    $url
)

    $oWindows = (New-Object -ComObject Shell.Application).Windows
    foreach ($oWindow in $oWindows.Invoke()) {
        if ($oWindow.Fullname -match "IEXPLORE.EXE" -and (($oWindow.LocationURL -match ".*$url") -or ($oWindow.LocationURL -eq "$url"))) {
            Write-verbose "Closing tab $($oWindow.LocationURL)"
            $oWindow.Quit()
        }
    }
}

<# Only allow access to following functions #>
Export-ModuleMember -Function Connect-SmartAzAccount
Export-ModuleMember -Function Connect-AzAccountManual
Export-ModuleMember -Function Get-AzAccountToken




# SIG # Begin signature block
# MIIaEwYJKoZIhvcNAQcCoIIaBDCCGgACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCe4UT7yHtVvq6i
# XNyxVTDiyhWt4Mip8euYwexrcOFoDaCCFk8wggQyMIIDGqADAgECAgEBMA0GCSqG
# SIb3DQEBBQUAMHsxCzAJBgNVBAYTAkdCMRswGQYDVQQIDBJHcmVhdGVyIE1hbmNo
# ZXN0ZXIxEDAOBgNVBAcMB1NhbGZvcmQxGjAYBgNVBAoMEUNvbW9kbyBDQSBMaW1p
# dGVkMSEwHwYDVQQDDBhBQUEgQ2VydGlmaWNhdGUgU2VydmljZXMwHhcNMDQwMTAx
# MDAwMDAwWhcNMjgxMjMxMjM1OTU5WjB7MQswCQYDVQQGEwJHQjEbMBkGA1UECAwS
# R3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHDAdTYWxmb3JkMRowGAYDVQQKDBFD
# b21vZG8gQ0EgTGltaXRlZDEhMB8GA1UEAwwYQUFBIENlcnRpZmljYXRlIFNlcnZp
# Y2VzMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAvkCd9G7h6naHHE1F
# RI6+RsiDBp3BKv4YH47kAvrzq11QihYxC5oG0MVwIs1JLVRjzLZuaEYLU+rLTCTA
# vHJO6vEVrvRUmhIKw3qyM2Di2olV8yJY897cz++DhqKMlE+faPKYkEaEJ8d2v+PM
# NSyLXgdkZYLASLCokflhn3YgUKiRx2a163hiA1bwihoT6jGjHqCZ/Tj29icyWG8H
# 9Wu4+xQrr7eqzNZjX3OM2gWZqDioyxd4NlGs6Z70eDqNzw/ZQuKYDKsvnw4B3u+f
# mUnxLd+sdE0bmLVHxeUp0fmQGMdinL6DxyZ7Poolx8DdneY1aBAgnY/Y3tLDhJwN
# XugvyQIDAQABo4HAMIG9MB0GA1UdDgQWBBSgEQojPpbxB+zirynvgqV/0DCktDAO
# BgNVHQ8BAf8EBAMCAQYwDwYDVR0TAQH/BAUwAwEB/zB7BgNVHR8EdDByMDigNqA0
# hjJodHRwOi8vY3JsLmNvbW9kb2NhLmNvbS9BQUFDZXJ0aWZpY2F0ZVNlcnZpY2Vz
# LmNybDA2oDSgMoYwaHR0cDovL2NybC5jb21vZG8ubmV0L0FBQUNlcnRpZmljYXRl
# U2VydmljZXMuY3JsMA0GCSqGSIb3DQEBBQUAA4IBAQAIVvwC8Jvo/6T61nvGRIDO
# T8TF9gBYzKa2vBRJaAR26ObuXewCD2DWjVAYTyZOAePmsKXuv7x0VEG//fwSuMdP
# WvSJYAV/YLcFSvP28cK/xLl0hrYtfWvM0vNG3S/G4GrDwzQDLH2W3VrCDqcKmcEF
# i6sML/NcOs9sN1UJh95TQGxY7/y2q2VuBPYb3DzgWhXGntnxWUgwIWUDbOzpIXPs
# mwOh4DetoBUYj/q6As6nLKkQEyzU5QgmqyKXYPiQXnTUoppTvfKpaOCibsLXbLGj
# D56/62jnVvKu8uMrODoJgbVrhde+Le0/GreyY+L1YiyC1GoAQVDxOYOflek2lphu
# MIIFbzCCBFegAwIBAgIQSPyTtGBVlI02p8mKidaUFjANBgkqhkiG9w0BAQwFADB7
# MQswCQYDVQQGEwJHQjEbMBkGA1UECAwSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYD
# VQQHDAdTYWxmb3JkMRowGAYDVQQKDBFDb21vZG8gQ0EgTGltaXRlZDEhMB8GA1UE
# AwwYQUFBIENlcnRpZmljYXRlIFNlcnZpY2VzMB4XDTIxMDUyNTAwMDAwMFoXDTI4
# MTIzMTIzNTk1OVowVjELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEtMCsGA1UEAxMkU2VjdGlnbyBQdWJsaWMgQ29kZSBTaWduaW5nIFJvb3Qg
# UjQ2MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAjeeUEiIEJHQu/xYj
# ApKKtq42haxH1CORKz7cfeIxoFFvrISR41KKteKW3tCHYySJiv/vEpM7fbu2ir29
# BX8nm2tl06UMabG8STma8W1uquSggyfamg0rUOlLW7O4ZDakfko9qXGrYbNzszwL
# DO/bM1flvjQ345cbXf0fEj2CA3bm+z9m0pQxafptszSswXp43JJQ8mTHqi0Eq8Nq
# 6uAvp6fcbtfo/9ohq0C/ue4NnsbZnpnvxt4fqQx2sycgoda6/YDnAdLv64IplXCN
# /7sVz/7RDzaiLk8ykHRGa0c1E3cFM09jLrgt4b9lpwRrGNhx+swI8m2JmRCxrds+
# LOSqGLDGBwF1Z95t6WNjHjZ/aYm+qkU+blpfj6Fby50whjDoA7NAxg0POM1nqFOI
# +rgwZfpvx+cdsYN0aT6sxGg7seZnM5q2COCABUhA7vaCZEao9XOwBpXybGWfv1Vb
# HJxXGsd4RnxwqpQbghesh+m2yQ6BHEDWFhcp/FycGCvqRfXvvdVnTyheBe6QTHrn
# xvTQ/PrNPjJGEyA2igTqt6oHRpwNkzoJZplYXCmjuQymMDg80EY2NXycuu7D1fkK
# dvp+BRtAypI16dV60bV/AK6pkKrFfwGcELEW/MxuGNxvYv6mUKe4e7idFT/+IAx1
# yCJaE5UZkADpGtXChvHjjuxf9OUCAwEAAaOCARIwggEOMB8GA1UdIwQYMBaAFKAR
# CiM+lvEH7OKvKe+CpX/QMKS0MB0GA1UdDgQWBBQy65Ka/zWWSC8oQEJwIDaRXBeF
# 5jAOBgNVHQ8BAf8EBAMCAYYwDwYDVR0TAQH/BAUwAwEB/zATBgNVHSUEDDAKBggr
# BgEFBQcDAzAbBgNVHSAEFDASMAYGBFUdIAAwCAYGZ4EMAQQBMEMGA1UdHwQ8MDow
# OKA2oDSGMmh0dHA6Ly9jcmwuY29tb2RvY2EuY29tL0FBQUNlcnRpZmljYXRlU2Vy
# dmljZXMuY3JsMDQGCCsGAQUFBwEBBCgwJjAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuY29tb2RvY2EuY29tMA0GCSqGSIb3DQEBDAUAA4IBAQASv6Hvi3SamES4aUa1
# qyQKDKSKZ7g6gb9Fin1SB6iNH04hhTmja14tIIa/ELiueTtTzbT72ES+BtlcY2fU
# QBaHRIZyKtYyFfUSg8L54V0RQGf2QidyxSPiAjgaTCDi2wH3zUZPJqJ8ZsBRNraJ
# AlTH/Fj7bADu/pimLpWhDFMpH2/YGaZPnvesCepdgsaLr4CnvYFIUoQx2jLsFeSm
# TD1sOXPUC4U5IOCFGmjhp0g4qdE2JXfBjRkWxYhMZn0vY86Y6GnfrDyoXZ3JHFuu
# 2PMvdM+4fvbXg50RlmKarkUT2n/cR/vfw1Kf5gZV6Z2M8jpiUbzsJA8p1FiAhORF
# e1rYMIIGGjCCBAKgAwIBAgIQYh1tDFIBnjuQeRUgiSEcCjANBgkqhkiG9w0BAQwF
# ADBWMQswCQYDVQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMS0wKwYD
# VQQDEyRTZWN0aWdvIFB1YmxpYyBDb2RlIFNpZ25pbmcgUm9vdCBSNDYwHhcNMjEw
# MzIyMDAwMDAwWhcNMzYwMzIxMjM1OTU5WjBUMQswCQYDVQQGEwJHQjEYMBYGA1UE
# ChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1YmxpYyBDb2Rl
# IFNpZ25pbmcgQ0EgUjM2MIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIBigKCAYEA
# myudU/o1P45gBkNqwM/1f/bIU1MYyM7TbH78WAeVF3llMwsRHgBGRmxDeEDIArCS
# 2VCoVk4Y/8j6stIkmYV5Gej4NgNjVQ4BYoDjGMwdjioXan1hlaGFt4Wk9vT0k2oW
# JMJjL9G//N523hAm4jF4UjrW2pvv9+hdPX8tbbAfI3v0VdJiJPFy/7XwiunD7mBx
# NtecM6ytIdUlh08T2z7mJEXZD9OWcJkZk5wDuf2q52PN43jc4T9OkoXZ0arWZVef
# fvMr/iiIROSCzKoDmWABDRzV/UiQ5vqsaeFaqQdzFf4ed8peNWh1OaZXnYvZQgWx
# /SXiJDRSAolRzZEZquE6cbcH747FHncs/Kzcn0Ccv2jrOW+LPmnOyB+tAfiWu01T
# PhCr9VrkxsHC5qFNxaThTG5j4/Kc+ODD2dX/fmBECELcvzUHf9shoFvrn35XGf2R
# PaNTO2uSZ6n9otv7jElspkfK9qEATHZcodp+R4q2OIypxR//YEb3fkDn3UayWW9b
# AgMBAAGjggFkMIIBYDAfBgNVHSMEGDAWgBQy65Ka/zWWSC8oQEJwIDaRXBeF5jAd
# BgNVHQ4EFgQUDyrLIIcouOxvSK4rVKYpqhekzQwwDgYDVR0PAQH/BAQDAgGGMBIG
# A1UdEwEB/wQIMAYBAf8CAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwGwYDVR0gBBQw
# EjAGBgRVHSAAMAgGBmeBDAEEATBLBgNVHR8ERDBCMECgPqA8hjpodHRwOi8vY3Js
# LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ1Jvb3RSNDYuY3Js
# MHsGCCsGAQUFBwEBBG8wbTBGBggrBgEFBQcwAoY6aHR0cDovL2NydC5zZWN0aWdv
# LmNvbS9TZWN0aWdvUHVibGljQ29kZVNpZ25pbmdSb290UjQ2LnA3YzAjBggrBgEF
# BQcwAYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AAb/guF3YzZue6EVIJsT/wT+mHVEYcNWlXHRkT+FoetAQLHI1uBy/YXKZDk8+Y1L
# oNqHrp22AKMGxQtgCivnDHFyAQ9GXTmlk7MjcgQbDCx6mn7yIawsppWkvfPkKaAQ
# siqaT9DnMWBHVNIabGqgQSGTrQWo43MOfsPynhbz2Hyxf5XWKZpRvr3dMapandPf
# YgoZ8iDL2OR3sYztgJrbG6VZ9DoTXFm1g0Rf97Aaen1l4c+w3DC+IkwFkvjFV3jS
# 49ZSc4lShKK6BrPTJYs4NG1DGzmpToTnwoqZ8fAmi2XlZnuchC4NPSZaPATHvNIz
# t+z1PHo35D/f7j2pO1S8BCysQDHCbM5Mnomnq5aYcKCsdbh0czchOm8bkinLrYrK
# pii+Tk7pwL7TjRKLXkomm5D1Umds++pip8wH2cQpf93at3VDcOK4N7EwoIJB0kak
# 6pSzEu4I64U6gZs7tS/dGNSljf2OSSnRr7KWzq03zl8l75jy+hOds9TWSenLbjBQ
# UGR96cFr6lEUfAIEHVC1L68Y1GGxx4/eRI82ut83axHMViw1+sVpbPxg51Tbnio1
# lB93079WPFnYaOvfGAA0e0zcfF/M9gXr+korwQTh2Prqooq2bYNMvUoUKD85gnJ+
# t0smrWrb8dee2CvYZXD5laGtaAxOfy/VKNmwuWuAh9kcMIIGhDCCBOygAwIBAgIQ
# Wzj4+IGGhK+UBDvj6IumZTANBgkqhkiG9w0BAQwFADBUMQswCQYDVQQGEwJHQjEY
# MBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSswKQYDVQQDEyJTZWN0aWdvIFB1Ymxp
# YyBDb2RlIFNpZ25pbmcgQ0EgUjM2MB4XDTIxMDcwMTAwMDAwMFoXDTI0MDYzMDIz
# NTk1OVowaTELMAkGA1UEBhMCR0IxFDASBgNVBAgMC094Zm9yZHNoaXJlMRgwFgYD
# VQQHDA9DaGlwcGluZyBOb3J0b24xFDASBgNVBAoMC1NjaS1OZXQgTHRkMRQwEgYD
# VQQDDAtTY2ktTmV0IEx0ZDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
# ANY5ObCXMpA7agdo7V0JR/UBWqPVAuJg1zFjfvol/BZEQi9TfrieysjKz0wx3rYW
# W0ENF20f94n2z7DKolncm0fT2ho0A63lHzOvGLv/vF+vQIXWNo9QlOKzxovkv0U8
# xvwUS2cBd5U1KL+aqrtLpFJArvcPbwNE8emvwwRKEG2IzjKsyIAxbU3xJSp5qa5X
# R4XIKx5GJR+aOepZgXmd/Qk4WexApo3rzMcBx8Ho+SC6kJqHeJoyn91og02n0T1P
# oops5ajyT3wcE/T9yYfwMWQGhJ1Puf3vhJsAdMImmDpPMs24wffPdalaC98Cuwt6
# CPgUVN+y9hB4BuedWgT5n1Bw60URj01Dpb6gEbVXrmSVDU7b3bB1l/+R+IeHFuZ3
# ee6LEUQ4d1voMyAXyxFw1ysy8GXIGcZjLpan7bwy8JndM8Hei9E5o7zRU38FzXB9
# voT6HFg/eL4wRGxNtMhu6d1wYm329FJ0SIokFdlJUCGX6yEtJAHBawEYEiuPIXXb
# wsaRy9wJmrOpOeBIR9c0HIrUYwTelTBWnvjSlCD6wibifvQQ+oYWYKa8dVZmqJum
# 9HIJ/v+vT7E9nWndn4T9dytRneKjVKW8rRuD3OzQqNXwhU6JCz5bJ8F3BDviJ6pf
# 5Wk9PqagKib4X4ajEdTUrQBSQXIa0uDQLjX5kSgXNiVtAgMBAAGjggG7MIIBtzAf
# BgNVHSMEGDAWgBQPKssghyi47G9IritUpimqF6TNDDAdBgNVHQ4EFgQUrFpYmz2K
# CuhaJOtPKPaNPDr1iQcwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwEQYJYIZIAYb4QgEBBAQDAgQQMEoGA1UdIARDMEEw
# NQYMKwYBBAGyMQECAQMCMCUwIwYIKwYBBQUHAgEWF2h0dHBzOi8vc2VjdGlnby5j
# b20vQ1BTMAgGBmeBDAEEATBJBgNVHR8EQjBAMD6gPKA6hjhodHRwOi8vY3JsLnNl
# Y3RpZ28uY29tL1NlY3RpZ29QdWJsaWNDb2RlU2lnbmluZ0NBUjM2LmNybDB5Bggr
# BgEFBQcBAQRtMGswRAYIKwYBBQUHMAKGOGh0dHA6Ly9jcnQuc2VjdGlnby5jb20v
# U2VjdGlnb1B1YmxpY0NvZGVTaWduaW5nQ0FSMzYuY3J0MCMGCCsGAQUFBzABhhdo
# dHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAdBgNVHREEFjAUgRJtYXR0QHNjaS1uZXQu
# Y28udWswDQYJKoZIhvcNAQEMBQADggGBAIyvq4BpTaNDYlRRmlp/6y8WCbcbOYLN
# qCLqTuEPcgIEGSFIGdzvTkSJhmMhbn6tMxz2TXp4dcOBMksTpI2VPP+xdjzUT3yR
# dp6tq0DcCG/WJbN/OwGQiTVvoY2Y4pwGzRcbh77ol6Tynp2Pw8BYa8WALL7WKx1T
# D3uOR84Zb8brHtnkVmDM+peok7XitE+rGaWvW5sQuGFGPYb9Oj9kfpzVGhgOR+E7
# 1qqeFUd8V8JeIBZt04ix0AbGkBbpA5F/Mo0Hljt7tTTo74usmz89Qw7VIFVoiezT
# g88uaQztwQFQJwN6Zp8pxY+f63SWtX18sYcY6WqBxjnXwNb/CsZEZRrQOYVQ7Y6V
# BaIzJ8Vb9ChVgtR7f1Ooua/M4seFWNQh6SnBPZ0Bm1gfRd9gO4KmsyzcUMiAxXaw
# RIvzflOSwsTKFFLSKeXSXHxsSB8PByKwcRM1CXjbAknoRwLk6nAD21nVY7PZe7gO
# jHPVO869CABQ3OA2lUxLd78rDG1NZK93KjGCAxowggMWAgEBMGgwVDELMAkGA1UE
# BhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDErMCkGA1UEAxMiU2VjdGln
# byBQdWJsaWMgQ29kZSBTaWduaW5nIENBIFIzNgIQWzj4+IGGhK+UBDvj6IumZTAN
# BglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqG
# SIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3
# AgEVMC8GCSqGSIb3DQEJBDEiBCDz8TiA5UkbduWwXsp3c2woCgkgPAd7Gqhb1fFf
# yOTxODANBgkqhkiG9w0BAQEFAASCAgBRw7XnyceZHPWVOOZMYP5sQ66WrZrtih9j
# EjncTFiJnhVFQxmFHDYzX00JwFZqRptmbr9N6zS/YJyZsz0s4JrWUMTCoASqBYJW
# beN+TMc7oTCxl769nNLBu3Tpz3EbUnfd5hkK+UHYUMktZZ8QwjwQa17r7cz+/FUP
# DxhTLD1K+Cjk+KFdCGCDyeoLLWOXfJiDC/WkKaFFrZEHgfS2uzD05qRoByF2w/zt
# TexpsbUl1RO0/kT7UA1SlCtriCvwIhFh+sdkNZR42PAbel/qn0lS7dLXo+z2hmod
# qmrp6FzRl92GKOAU/xgR3wOwu1eXfyFKLcb/X7Z0o22WXImU5xA0sbs0Z6y1mdv4
# DAKO+9CQLM8pwK9ncJ5SHxGitJx8Cc7TwGT9bvtJpRarccxiNjcMCQX5KB69a5l7
# gElX6gs4AH4apEhgq1PdKISaJLtN9Uu36fStT272cpxr483lLpZkynKg/TjqSXjb
# dNcdsjzkvL93n7JIyI3a1Vne5b+qAXh6v3xL+/Ga7DczFZzttj+3n/nGQRw1bxgC
# kVjE7MpaOZSK0ZRysmlkbI+NVdPQJBJDsNgatbOkWAEQYtRGSvLLpi+uxxx0qzur
# QKphwScRENP4t018YagJLP1m38Scnlm2R972Y6Z9zsgRtF3isZDS1Q9+qcsrsZhR
# FxjbCcTVYg==
# SIG # End signature block
