## Sci-Net.PowerShell.AzureLogin
## Allows automated login to Azure through Powershell, without the need to perform MFA if the account is used in a web session
## Adds following functons:
##   Connect-SmartAzAccount
##   Connect-AzAccountManual
##   Get-AzAccountToken

## Version 1.0.0

## Developed by: 
## Nathan Moore on behalf of Sci-Net Business Solutions
## 03-Mar-23

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
    [switch] $InstallPrerequisites
)
    Check-AzAccountsPreRequisite -InstallPrerequisites $InstallPrerequisites
       
    $Emailaddress = whoami.exe /upn
    try {
        if ($Emailaddress -match ".*@.*") {
            $AttemptLogin = Switch([System.Windows.MessageBox]::Show(("Would you like to try logging into Azure as $EmailAddress ?"),"Azure Account Selection",[System.Windows.MessageBoxButton]::YesNo,[System.Windows.MessageBoxImage]::Question)) { 'Yes' {$true} 'No' {$False} }
            if ($AttemptLogin) {
                if ($debugMode) {
                    $Token = Get-AzAccountToken -EmailAddress $Emailaddress -DebugMode
                    Write-Host ("Obtained Token:" + ($Token | Out-String))
                } else {
                    $Token = Get-AzAccountToken -EmailAddress $Emailaddress
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
                Connect-AzAccountManual -TenantID $TenantID -SuppressWarnings $SuppressWarnings
            }
        } else {
            Write-Host "No valid account found, please login:"
            Connect-AzAccountManual -TenantID $TenantID -SuppressWarnings $SuppressWarnings
        }
    } catch {
        Write-Host "Unable to login automatically, please login:"
        if ($DebugMode) {Write-host $_ -ForegroundColor Red}
        Connect-AzAccountManual -TenantID $TenantID -SuppressWarnings $SuppressWarnings
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
    [Boolean] $SuppressWarnings,
    #Include to automatically install pre-requsite modules
    [switch] $InstallPrerequisites
)
    
    Check-AzAccountsPreRequisite -InstallPrerequisites $InstallPrerequisites
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
    [switch] $DebugMode
)
    <# Get-AzAccountToken Function Body #>

    if ($debugMode) {Write-Host "Begin Code Obtain"}
    $ResponsePort = Get-InactiveTcpPort -Start 1024 -End 65535
    if ($debugMode) {Write-Host "ChosenPort: $ResponsePort"}
    
    $URL = "https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?scope=https://management.core.windows.net//.default openid profile offline_access&response_type=code&client_id=1950a258-227b-4e31-a9cf-717495945fc2&redirect_uri=http://localhost:$ResponsePort/&x-client-SKU=MSAL.Desktop&x-client-Ver=4.21.0.0&x-client-CPU=x64&x-client-OS=Microsoft Windows NT 10.0.19042.0&login_hint=$EmailAddress"
    $ie = New-Object -COM 'InternetExplorer.Application'
    #$ie.Visible = $true
    $ie.Navigate($URL)
    $RawURL = Get-HttpQueryParametersSentToLocalhost -Verbose -Port $ResponsePort
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
    [Boolean] $InstallPrerequisites
)
    if ("Az.Accounts" -notin (Get-Module).Name) {
        if ($InstallPrerequisites) {
            start-process -FilePath "powershell" -ArgumentList "Install-Module az.accounts" -WindowStyle Hidden
        } else {
            Throw "Missing Pre-Requsite Modules : Ensure Az.Accounts module is installed and available before trying again. Run:`nInstall-Module Az.Accounts"
        }
    }
    import-module Az.Accounts
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
        if ($oWindow.Fullname -match "IEXPLORE.EXE" -and $oWindow.LocationURL -match ".*$url") {
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCxhsT8ducSjPAA
# /aoaeJRcYFeI4Z5TjdRCnIkQqo3jHqCCFk8wggQyMIIDGqADAgECAgEBMA0GCSqG
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
# AgEVMC8GCSqGSIb3DQEJBDEiBCCHJpShS5yu/fiwSdg1W76W4Z3OGFjaI7oacrMm
# ko87XzANBgkqhkiG9w0BAQEFAASCAgAR5ZuzCrIHMcvwupbK5qk6x1hsX0hfW5we
# xvyXWJiwfcZ83By+9pUoe/C8yDUClE4DZXPL7yFcweOISO3Msqltnt0le0J8Vf5t
# 8OrxkthdbrCpIXCRa+BrPunqD7ADsVIIHC1rx812gNf84Xk+cnHSTgg+oreQklAl
# XJT3AvK14FpTZ1+GOwWfWXL9VBn8qGEg0rxP5wEBC7MWqgHzwaMW/9CM8RQgnk5y
# U9Dhh2Hz95PveDGUNi4KrsFTO0ix7Wylitys0m5QNHNk2PMB74NWmazkCqP+Buvj
# wYyBKZSJNAaLQoNmTpfjtIkiVsXNu3dRNWF22S451Vie0PvvmqjVivFfwfQ0SObl
# XoInrQpl9WuUNxXnZtjv7bNFvljzB7AhJ/Jhu+QE1jQuAKz6FSSAkRd+WB6V8jWq
# mYefbOqhFfQLp4y8OuCiyBU2Z854BUQogm3idhqoW5enWXMqAx2iA8RcLH4d6wG7
# vanXhERovUaN17mwSlkxXRUJo/uVaGhA0P6kIwsl7jJ8pwSUxsVRWdeu63m1dmC6
# GUChCUPKuB+KqxYo5hl1RQ1j06M8vvi1mahRWw6SV7UZXAVNIZu7YZhumCKB5Ge9
# zgHP6Q4coAjHszxd1LOrMP+XCu1TRHiNTzPheOyjhkoRupTYv88OOgzVBhvvpF0B
# RBCSl4qczQ==
# SIG # End signature block
