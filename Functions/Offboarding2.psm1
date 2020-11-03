function Connect-OnlineExchange {
    [CmdletBinding()]
    param (
        [Parameter()]
        [PSCredential]
        $AAD_Credential
    )

    $OpenPSSessions = Get-PSSession

    # If there is an open session to Office 365, we do not re-connect.
    if ($OpenPSSessions.ComputerName -contains 'outlook.office365.com' -and $OpenPSSessions.Availability -eq 'Available') {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm);		Write-Host "[$timer] - Exchange Online already available."
    }

    # If there is no open session, then we do connect
    else {

        # import exchange online module
        If (!(Get-Module -Name ExchangeOnlineManagement)) {
            $timer = (Get-Date -Format yyy-MM-dd-HH:mm);		Write-Host "[$timer] - Importing Exchange Online modules"
            [void] (Import-Module ExchangeOnlineManagement -Verbose:$false)
        }

        # connect
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm);		Write-Host "[$timer] - Connecting to Exchange Online." 
        [void] (Connect-ExchangeOnline -Credential $AAD_Credential -ShowProgress $false -ShowBanner:$false)

    }
}

function Get-InactiveUsers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [PSCredential]
        $AAD_Credential,
        [string]
        $server,
        $since,
        $OU,
        [Parameter(Mandatory = $false)]
        [switch]
        $First50
    )
    
    #1. Collect inactive users
    <#
    These users are:
    - within the lieaver OU
    - have never logged in 
    - or their last login date was outside of the defined - currently 365 - days)
    - are not service accounts
    - ordered by their creation date
    #>

    if ($First50.IsPresent) {
        $number = 50
        Write-Host #lazy line break
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm);		Write-Host "[$timer] - Only the first 50 inactive users are returned"
        Write-Host #lazy line break
        $InactiveLeavers = Get-ADUser -Credential $AAD_Credential -Server $server -Filter * -SearchBase $OU -Properties * | Where-Object { ($null -EQ $_.LastLogonDate) -or ($_.LastLogonDate -lt $since) -and ($_.Name -notlike 'svc.*') } | Select-Object Name, EmailAddress, LastLogonDate, whenCreated, UserPrincipalName, SAMAccountName, DistinguishedName | Sort-Object whenCreated | Select-Object -First $number
    }
    else {
        Write-Host #lazy line break
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm);		Write-Host "[$timer] - All inactive users are returned"
        Write-Host #lazy line break
        $InactiveLeavers = Get-ADUser -Credential $AAD_Credential -Server $server -Filter * -SearchBase $OU -Properties * | Where-Object { ($null -EQ $_.LastLogonDate) -or ($_.LastLogonDate -lt $since) -and ($_.Name -notlike 'svc.*') } | Select-Object Name, EmailAddress, LastLogonDate, whenCreated, UserPrincipalName, SAMAccountName, DistinguishedName | Sort-Object whenCreated
    }

    return $InactiveLeavers
}

function Get-UsersWithMailbox {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $users
    )

    $x = 0
    $UsersWithMailbox = $null
    $UsersWithMailbox = @()

    Write-Host #lazy line break
    foreach ($u in $users) {

        $x++
        $Obj = New-Object -TypeName PSObject

        #Actual mailbox name
        $MailboxAddress = $u.UserPrincipalName
        $ADAccount = $u.SAMAccountName
        $lastLogonDate = $u.LastLogonDate
        $whenCreated = $u.whenCreated

        #Users with mailbox
        if (Get-Mailbox $MailboxAddress -ErrorAction SilentlyContinue) {
            #Report on screen
            Write-Host "Mailbox [$MailboxAddress] found [$x / $($users.count)]" -ForegroundColor Green
            #Report in array
            $details = Get-Mailbox $MailboxAddress | Select-Object Name, *type*
            $Obj | Add-Member -MemberType NoteProperty -Name AccountName -Value $ADAccount
            $obj | Add-Member -MemberType NoteProperty -Name AccountCreationDate -Value $whenCreated        
            $obj | Add-Member -MemberType NoteProperty -Name LastLogonDate -Value $lastLogonDate
            $Obj | Add-Member -MemberType NoteProperty -Name MailboxName -Value $($details.Name)
            $Obj | Add-Member -MemberType NoteProperty -Name Type -Value $($details.RecipientTypeDetails)
            $Obj | Add-Member -MemberType NoteProperty -Name EmailAddress -Value $MailboxAddress
            $UsersWithMailbox += $obj
        
        }
    
    }
    Write-Host #lazy line break

    return $UsersWithMailbox
}

function Get-UsersWithoutMailbox {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $users
    )

    $x = 0
    $UsersWithoutMailbox = $null
    $UsersWithoutMailbox = @()

    Write-Host #lazy line break
    foreach ($u in $users) {

        $x++
        $Obj = New-Object -TypeName PSObject

        #Actual mailbox name
        $MailboxAddress = $u.UserPrincipalName
        $ADAccount = $u.SAMAccountName
        $lastLogonDate = $u.LastLogonDate
        $whenCreated = $u.whenCreated

        #Users with mailbox
        if (!(Get-Mailbox $MailboxAddress -ErrorAction SilentlyContinue)) {
            #Report on screen
            Write-Host "Mailbox [$MailboxAddress] not found [$x / $($Users.count)]" -ForegroundColor Red
            #Write-Host "Mailbox [$MailboxAddress] found" -ForegroundColor Green
            #Report in array
            $Obj | Add-Member -MemberType NoteProperty -Name AccountName -Value $ADAccount
            $obj | Add-Member -MemberType NoteProperty -Name AccountCreationDate -Value $whenCreated        
            $obj | Add-Member -MemberType NoteProperty -Name LastLogonDate -Value $lastLogonDate        
            $Obj | Add-Member -MemberType NoteProperty -Name MailboxName -Value 'n/a'
            $Obj | Add-Member -MemberType NoteProperty -Name Type -Value 'n/a'
            $Obj | Add-Member -MemberType NoteProperty -Name EmailAddress -Value 'n/a'
            $UsersWithoutMailbox += $obj
        }
    
    }
    Write-Host #lazy line break

    return $UsersWithoutMailbox
}

function Set-LitigationHold {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $user,
        [int]
        $litigationHoldTime,
        [Parameter(Mandatory = $false)]
        [switch]
        $Test,
        [switch]
        $Westcoast

    )

    if ($Test.IsPresent) {
        Write-Verbose "Setting user mailbox $($user.EmailAddress) to litigation for $litigationHoldTime days" -Verbose
    }
    else {

        # set the mailbox to litigation
        Set-Mailbox -Identity $($user.EmailAddress) -LitigationHoldEnabled $true -LitigationHoldDuration $litigationHoldTime -ErrorAction SilentlyContinue #-WhatIf
        
        # Wait for litigation to take effect - but only for WC, as XMA probably wont be able to configure this.
        if ($Westcoast.IsPresent) {
            do {
                $litigationHoldStatus = $null
                $litigationHoldStatus = (Get-mailbox -Identity $($user.EmailAddress)).LitigationHoldEnabled
                if ($litigationHoldStatus -ne 'True') {
                    Write-Output "Litigation hold is not yet enabled on [$($user.EmailAddress)]. Retry in 10 seconds"
                    Start-Sleep 10
                }
                else {
                    Write-Output "Mailbox [$($user.EmailAddress)] is in litigation hold." 
                }
            } until ($litigationHoldStatus -eq 'True')
        }

    }
}


function Send-EmailReports {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [PSCustomObject]
        $recipients,
        [Parameter(Mandatory = $true)] [string]
        $SmtpServer,
        $emailSender,
        $attachments,
        $dom,
        $days

    )

    $TextEncoding = [System.Text.Encoding]::UTF8
    $EmailSubject = "Historical user removal - $dom (inactive more than $days days )"

    #Construct Password Email Body
    $EmailBody =
    "
            <font face= ""Century Gothic"">
            Hello,
            <p> Please find attached the latest batch of old user accounts that has been purged.<br>

            <p> Please note the following: <br>

            <ul style=""list-style-type:disc"">
            <li> <p> These AD accounts have now been deleted.</li>
            <li> <p> The mailbox of these users has also been removed.</li>
            <li> <p> For Westcoast users restoration is possible using normal backup tools, and mailboxes are also has been put to <span style=`"color:blue`">" + ' litigation hold ' + '</span>. </li>
            <li> <p> For XMA users restoration is possible using normal backup tools. </li>
            
            </ul>

            <p> Thank you. <br>
            <p>Regards, <br>
            Westcoast Group IT
            </P>
            </font>
        '

    #Send email
    foreach ($r in $recipients) {

        $EmailRecipient = $r.recipient

        Write-Host "Sending email to $EmailRecipient !"

        Send-MailMessage -SmtpServer $SmtpServer -From $emailSender -To $EmailRecipient -Subject $EmailSubject -Body $EmailBody -BodyAsHtml -Priority High -Encoding $TextEncoding -Attachments $attachments # -ErrorAction SilentlyContinue

    }

}

# SIG # Begin signature block
# MIIOWAYJKoZIhvcNAQcCoIIOSTCCDkUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU76q+4ggqeKRNK6UbyAL4U7m1
# dwKgggueMIIEnjCCA4agAwIBAgITTwAAAAb2JFytK6ojaAABAAAABjANBgkqhkiG
# 9w0BAQsFADBiMQswCQYDVQQGEwJHQjEQMA4GA1UEBxMHUmVhZGluZzElMCMGA1UE
# ChMcV2VzdGNvYXN0IChIb2xkaW5ncykgTGltaXRlZDEaMBgGA1UEAxMRV2VzdGNv
# YXN0IFJvb3QgQ0EwHhcNMTgxMjA0MTIxNzAwWhcNMzgxMjA0MTE0NzA2WjBrMRIw
# EAYKCZImiZPyLGQBGRYCdWsxEjAQBgoJkiaJk/IsZAEZFgJjbzEZMBcGCgmSJomT
# 8ixkARkWCXdlc3Rjb2FzdDEmMCQGA1UEAxMdV2VzdGNvYXN0IEludHJhbmV0IElz
# c3VpbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC7nBk9j3wR
# GgkxrPuXjIXlptisoOhKZp7KCB+BhxaxlTGW5lxhEaNirirM4jaM04kXojFZxhHV
# lTl2W3TPOfeIEXxcZYigPgh9d6wgTTb2cSRq1872YjMytxSps14LAbY8CEu+fQmC
# AbL6V8EgtnAmzMBBqOOi6x7bMHoGkJPwDOSUM01LHPoT8cg9KVIFioJHpex/Xeko
# FiRwgW7uS+dh57iCGRWVCZaDrFIXWKj4dOHJigsEPkbmJUPSYILF8SYglFiJpM7b
# xl3RPuy2GvJRq5Ikyn0SvnpAG72Ge664PV5sFdtzdNkIE7RsE6zUEqK1v2pt7CcC
# qh4en3v54ouZAgMBAAGjggFCMIIBPjASBgkrBgEEAYI3FQEEBQIDAQABMCMGCSsG
# AQQBgjcVAgQWBBSBYkDZbTpVK0nuvapWivWUf0tBKDAdBgNVHQ4EFgQUU3PVQuhx
# ickSLEsfPyKpNozqrT8wGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0P
# BAQDAgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAUuxfhV4noKzmJ
# eDD6ejIRp0cSBu8wPQYDVR0fBDYwNDAyoDCgLoYsaHR0cDovL3BraS53ZXN0Y29h
# c3QuY28udWsvcGtpL3Jvb3RjYSgxKS5jcmwwSAYIKwYBBQUHAQEEPDA6MDgGCCsG
# AQUFBzAChixodHRwOi8vcGtpLndlc3Rjb2FzdC5jby51ay9wa2kvcm9vdGNhKDEp
# LmNydDANBgkqhkiG9w0BAQsFAAOCAQEAaYMr/xfHuo3qezz8rtbzGkfUwqNFjd0s
# 7d02B07aO5q0i7LMtZTMxph9DbeJRvm+d8Sr4DSiWgtJdb0eYsx4xj5lDrsXDuO2
# 2Mb4hKjtqzDVW5PEJzC72BPOSfkgfW6PZmscMPtJnn0TPM24DzkYmjhnsA97Ltjv
# 1wuvUi2G0nPIbzfBZWnnuCx5PhSovssQU5E3ZlVLew6a8WME0lPOmR9c38TARqWh
# tvS/wqmUaCEUF6rmUDY0MgY/Wrg2TIbtlYFWe9PksI4jmTE4Ndy5BW8smx+8YOoF
# fCOldshHHgFJVG7Bat6vrT8AaUSs6crPBRMpbeouD0iujXts+LdV2TCCBvgwggXg
# oAMCAQICEzQAA+ZyHBAttK7qIqcAAQAD5nIwDQYJKoZIhvcNAQELBQAwazESMBAG
# CgmSJomT8ixkARkWAnVrMRIwEAYKCZImiZPyLGQBGRYCY28xGTAXBgoJkiaJk/Is
# ZAEZFgl3ZXN0Y29hc3QxJjAkBgNVBAMTHVdlc3Rjb2FzdCBJbnRyYW5ldCBJc3N1
# aW5nIENBMB4XDTIwMDUxODA4MTk1MloXDTI2MDUxODA4Mjk1MlowgacxEjAQBgoJ
# kiaJk/IsZAEZFgJ1azESMBAGCgmSJomT8ixkARkWAmNvMRkwFwYKCZImiZPyLGQB
# GRYJd2VzdGNvYXN0MRIwEAYDVQQLEwlXRVNUQ09BU1QxDTALBgNVBAsTBExJVkUx
# DjAMBgNVBAsTBVVTRVJTMQ8wDQYDVQQLEwZBZG1pbnMxHjAcBgNVBAMTFUZhYnJp
# Y2UgU2VtdGkgKEFETUlOKTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEB
# APVwqF2TGtzPlxftCjtb23neDu2cWyovIpo1TgU0ptNYrJM8tAY6W8Yt5Vw+8xzU
# 45sxmbMzU2JpJaqEPFe3+gXWJtL99/ZusyXCDbubzYmNu06WE6XqMqG/KRfZ3BpN
# Gw5s3KlxWVj/H12i7JPbMvfyAl8lgz/YBO0XVdoozcAglEck7c8DBaRTb4J7vX/O
# IS7dYu+gmkZJCv2+O6vTNTlK7bIHAQPWzSPibzU9dRPlHiPOTcHoYB+YNpmbgNxn
# fdaFMB+xY1GcYoKwVRl6UEF/od8TKehzUp/hHFlXiH+miz692ptXhi3dOp6R4Stn
# Ku0IoBfBi/CQcgl5Uko6kckCAwEAAaOCA1YwggNSMD4GCSsGAQQBgjcVBwQxMC8G
# JysGAQQBgjcVCIb24huEi+UUg4mdM4f4p0GE8aVDgSaGkPwogZ23PAIBZAIBAjAT
# BgNVHSUEDDAKBggrBgEFBQcDAzALBgNVHQ8EBAMCB4AwGwYJKwYBBAGCNxUKBA4w
# DDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQU7eheFlEriypJznAoYQVEx7IAmBkwHwYD
# VR0jBBgwFoAUU3PVQuhxickSLEsfPyKpNozqrT8wggEuBgNVHR8EggElMIIBITCC
# AR2gggEZoIIBFYY6aHR0cDovL3BraS53ZXN0Y29hc3QuY28udWsvcGtpLzAxX2lu
# dHJhbmV0aXNzdWluZ2NhKDEpLmNybIaB1mxkYXA6Ly8vQ049V2VzdGNvYXN0JTIw
# SW50cmFuZXQlMjBJc3N1aW5nJTIwQ0EoMSksQ049Qk5XQURDUzAxLENOPUNEUCxD
# Tj1QdWJsaWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNlcyxDTj1Db25maWd1
# cmF0aW9uLERDPXdlc3Rjb2FzdCxEQz1jbyxEQz11az9jZXJ0aWZpY2F0ZVJldm9j
# YXRpb25MaXN0P2Jhc2U/b2JqZWN0Q2xhc3M9Y1JMRGlzdHJpYnV0aW9uUG9pbnQw
# ggEmBggrBgEFBQcBAQSCARgwggEUMEYGCCsGAQUFBzAChjpodHRwOi8vcGtpLndl
# c3Rjb2FzdC5jby51ay9wa2kvMDFfaW50cmFuZXRpc3N1aW5nY2EoMSkuY3J0MIHJ
# BggrBgEFBQcwAoaBvGxkYXA6Ly8vQ049V2VzdGNvYXN0JTIwSW50cmFuZXQlMjBJ
# c3N1aW5nJTIwQ0EsQ049QUlBLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2VzLENO
# PVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9d2VzdGNvYXN0LERDPWNvLERD
# PXVrP2NBQ2VydGlmaWNhdGU/YmFzZT9vYmplY3RDbGFzcz1jZXJ0aWZpY2F0aW9u
# QXV0aG9yaXR5MDUGA1UdEQQuMCygKgYKKwYBBAGCNxQCA6AcDBp3Y2FkbWluLmZz
# QHdlc3Rjb2FzdC5jby51azANBgkqhkiG9w0BAQsFAAOCAQEAeM0HkiWDX+fmhIsv
# WxZb+D/tLDztccfYND16zFAoReu0VmTUz570CEMhLyHGh1jk3y/pb26UmjqHFeVh
# /EVu/EQNCuT5gQPKh64FQsBVinugNHWMhDySywykKwkdnqEpY++UNxQyyj6xpTM0
# tg+h8Wd1IlDN98SwLBy4x16SwgGTdwKvU9CyBuMRQjPlSJKjCL+14T0C8d2SBGW3
# 9uLCqjyMd288Q3QgrbDoHSg/x+vsnrDzOHMThM/2aMPbcO0wqafK9G5qdoIc0dqe
# So/vU6rsNLwQ1sniJQxerKZnWJjEfl8M5OcUxws5n7D3fqpHZ2VxLCIYp6yuPkHY
# R5daezGCAiQwggIgAgEBMIGCMGsxEjAQBgoJkiaJk/IsZAEZFgJ1azESMBAGCgmS
# JomT8ixkARkWAmNvMRkwFwYKCZImiZPyLGQBGRYJd2VzdGNvYXN0MSYwJAYDVQQD
# Ex1XZXN0Y29hc3QgSW50cmFuZXQgSXNzdWluZyBDQQITNAAD5nIcEC20ruoipwAB
# AAPmcjAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkq
# hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGC
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUgDfV5L+VP+aLIuQHmjaoM4+tYOowDQYJKoZI
# hvcNAQEBBQAEggEAV+EXbXh6eP3V4O2Dtxd3n/ebeOfjj1n8sVT3zEHBzXYE96M1
# F0NZrVpwgH/go9eAuhZDQrRoXM/8U5AGQPr0JvuxLnCBC1ob8sckLo2wH2cAZunW
# UG87sh0UcjkJ8SuXGsPiyqznF6zRIB1ozQks4jkn/lMmKzB65Z1FveMQxCrIQ7GC
# Z7GwlA7ur45esGRiHoizJSh8SN30foqNmGimXPaGxWvVQ2cv8AgHZVJAs5hbVtsy
# wEIW0VZfHK5+Z82JsNQ7RFBe6WRinNB4qCSgmKA8orV7DaOnJpThlQRgH+wHBj8v
# jwoJa5dOdtuHCG2e2pR25nInIWwZnO242dL9pQ==
# SIG # End signature block
