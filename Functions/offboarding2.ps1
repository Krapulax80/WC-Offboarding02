
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)] [switch]$Live,
    [Parameter(Mandatory = $false)] [switch]$Test ,
    [Parameter(Mandatory = $False)]
    [string]
    $login_file_wc_AD = 'c:\Scripts\AD\Offboarding02\Credentials\wc_scriptcreds_AD.xml',
    $login_file_xma_AD = 'c:\Scripts\AD\Offboarding02\Credentials\xma_scriptcreds_AD.xml',  
    $login_file_wc_AAD = 'c:\Scripts\AD\Offboarding02\Credentials\wc_scriptcreds_AAD.xml',
    $login_file_xma_AAD = 'c:\Scripts\AD\Offboarding02\Credentials\xma_scriptcreds_AAD.xml',            
    $logFolder = ('c:\Scripts\AD\Offboarding02\logs\' + (Get-Date -Format yyy-MM-dd)),
    [Parameter(Mandatory = $False)]
    [switch]
    $wc,
    [Parameter(Mandatory = $False)]
    [switch]
    $xma
)
    
BEGIN {

    $ErrorActionPreference = 'Stop'

    # declare work folders
    # $FunctionFolder = 'Functions'
    # $InputFolder = 'Input'
    # $OutputFolder = 'Output'
    # $ConfigFolder = 'Config'
    # $loginFolder = 'Credentials'
    # $ReportFolder = 'Reports'
    # $Logs = 'Logs'
        
    # # save the current path
    # $CurrentPath = $null
    # $CurrentPath = Split-Path -Parent $PSCommandPath
    # Set-Location $CurrentPath   
    
    # import configurations
    if ($WC.Ispresent) {
        $config = Import-Csv ('c:\Scripts\AD\Offboarding02\config\' + 'westcoast.csv')
    }
    elseif ($xma.Ispresent) {
        $config = Import-Csv ('c:\Scripts\AD\Offboarding02\config\' + 'xma.csv')
    }

    # configuration parameters
    $searchBase = $config.searchbase
    $systemDomain = $config.systemdomain
    $DomainNetBIOS = $config.domainnetbios
    $infraserver = $config.$infraserver#"BNWINFRATS01.westcoast.co.uk"

    $SmtpServer = $config.SMTPServer
    $emailSender = 'offboarding-2@westcoast.co.uk'  
    
    $daysInactive = 365
    [int]$litigationHoldTime = 2555
    #$date = Get-date -Format dd_MM_yyyy
    $inactiveDate = (Get-Date).Adddays( - ($daysInactive))

    # import work module
    Import-Module 'c:\Scripts\AD\Offboarding02\Functions\Offboarding2.psm1' -Force

    # import recipients lists
    if ($Test.IsPresent) {
        $recipientCSV = 'c:\Scripts\AD\Offboarding02\config\' + 'test_recipients.csv'
    }
    else {
        $recipientCSV = 'c:\Scripts\AD\Offboarding02\config\' + 'recipients.csv'
    }
    $recipients = Import-Csv $recipientCSV

    # create credentials for the work and start transcripting
    if ($WC.Ispresent) {
        if (-not (Test-Path $login_file_wc_AD)) {
            Get-Credential -Message "Please supply AD login credentials for WestCoast [$($config.AD_Admin)]" | Export-Clixml -Path $login_file_wc_AD -Force
        }
        if (-not (Test-Path $login_file_wc_AAD)) {
            Get-Credential -Message "Please supply AAD login credentials for WestCoast [$($config.AAD_Admin)]" | Export-Clixml -Path $login_file_wc_AAD -Force
        }
    }
    elseif ($XMA.IsPresent) {
        if (-not (Test-Path $login_file_xma_AD)) {
            Get-Credential -Message "Please supply AD login credentials for XMA [$($config.AD_Admin)]" | Export-Clixml -Path $login_file_xma_AD -Force
        } 
        if (-not (Test-Path $login_file_xma_AAD)) {
            Get-Credential -Message "Please supply AAD login credentials for XMA [$($config.AAD_Admin)]" | Export-Clixml -Path $login_file_xma_AAD -Force
        }  
    }

    if (-not (Test-Path $logFolder )) {
        [void] (New-Item -Path $logFolder -ItemType Directory -Force -ErrorAction SilentlyContinue)
    }
    if ($wc.IsPresent) {
    
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)
        Write-Host # lazy line break
        Write-Host -BackgroundColor Black "[$timer] - Working on WestCoast domain"

        $AD_Credential = Import-Clixml -Path $login_file_wc_AD
        $AAD_Credential = Import-Clixml -Path $login_file_wc_AAD
        #Connect-MsolService -Credential $AD_Credential
        $transcriptFile = ($logFolder + '\transcript_WC.log')
        $errorFile = ($logFolder + '\errors_WC.log')
        Start-Transcript -Path $transcriptFile -Force
    }
    elseif ($xma.IsPresent) {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)
        Write-Host # lazy line break
        Write-Host -BackgroundColor Black "[$timer] - Working on XMA domain"

        $AD_Credential = Import-Clixml -Path $login_file_xma_AD
        $AAD_Credential = Import-Clixml -Path $login_file_xma_AAD
        #Connect-MsolService -Credential $AD_Credential
        $transcriptFile = ($logFolder + '\transcript_XMA.log')
        $errorFile = ($logFolder + '\errors_XMA.log')
        Start-Transcript -Path $transcriptFile -Force
    }
    else {
        $timer = (Get-Date -Format yyy-MM-dd-HH:mm)
        Write-Host -ForegroundColor Red "[$timer] - Domain unrecoqnized - terminating"
        Stop-Transcript -ErrorAction Ignore
        Break
    }

    # store work server (DC)
    $DC = (Get-ADForest -Identity $systemDomain -Credential $AD_Credential |	Select-Object -ExpandProperty RootDomain |	Get-ADDomain |	Select-Object -Property PDCEmulator).PDCEmulator

    # connect to exchange online
    Get-PSSession | Remove-PSSession
    Connect-OnlineExchange -AAD_Credential $AAD_Credential

}
    
PROCESS {
        
    Write-Host
    Write-Host -BackgroundColor Black 'Starting the offboarding process'

    # collect inactive leavers
    $InactiveLeavers = Get-InactiveUsers -AAD_Credential $AD_Credential -server $DC -since $inactiveDate -OU $searchBase -First50
    
    # collect leavers with mailbox
    $leaversWithMailbox = Get-UsersWithMailbox -users $InactiveLeavers
    #$leaversWithMailbox | Format-Table -AutoSize -Wrap

    # collect leavers without mailbox
    $leavresWithoutMailbox = Get-UsersWithoutMailbox -users $InactiveLeavers
    #$leavresWithoutMailbox | Format-Table -AutoSize -Wrap

    # save leavers with mailbox to a CSV
    $today = Get-Date -Format ddMM
    $withMailboxCSV = 'c:\Scripts\AD\Offboarding02\output\' + $today + '_' + $DomainNetBIOS + '_leavers_WITH_mailbox.csv'
    $leaversWithMailbox | Export-Csv -Path $withMailboxCSV -Force

    # save leavers without a mailbox to a CSV
    $today = Get-Date -Format ddMM
    $withoutMailboxCSV = 'c:\Scripts\AD\Offboarding02\output\' + $today + '_' + $DomainNetBIOS + '_leavers_WITHOUT_mailbox.csv' 
    $leavresWithoutMailbox | Export-Csv -Path $withoutMailboxCSV -Force

    # set the leavers with mailboxes to litigation
    foreach ($user in $leaversWithMailbox) {

        if ($Test.IsPresent) {
            # test mode
            if ($WC.IsPresent) {
                Set-LitigationHold -user $user -litigationHoldTime $litigationHoldTime -WestCoast -Test
            }
            else {
                Set-LitigationHold -user $user -litigationHoldTime $litigationHoldTime -Test
            }
            
        }
        else {
            # live mode
            if ($wc.IsPresent) {
                Set-LitigationHold -user $user -litigationHoldTime $litigationHoldTime -Westcoast
            }
            else {
                Set-LitigationHold -user $user -litigationHoldTime $litigationHoldTime
            }
            
        }
        
    }

    #region  delete the AD objects
    # $today = Get-Date -Format ddMM
    # if ($wc.IsPresent) {
    #     $deletionErrorFile = ($logFolder + '\deletion_WC.log')
    # }
    # elseif ($xma.IsPresent) {
    #     $deletionErrorFile = ($logFolder + '\deletion_XMA.log')
    # }

    Write-Host #lazy line break
    foreach ($user in $InactiveLeavers) {
    
        if ($Test.IsPresent) {
            # test mode
            Write-Verbose "Deleting user object $($user.UserPrincipalName)" -Verbose
        }
        else {
            # live mode
            try {
                Remove-ADObject -Identity $user.DistinguishedName -Server $DC -Credential $AD_Credential -Recursive -Verbose -Confirm:$false #-ErrorAction Continue #-WhatIf
            }
            catch {
                $user.DistinguishedName | Out-File $errorFile -Append -Force
                $_.Exception.Message | Out-File $errorFile -Append -Force
                Continue
            }
            
        }
    }
    Write-Host #lazy line break
    #endregion

    # define attachments
    $attachments = @()
    # if (Test-Path -Path $deletionErrorFile -ErrorAction Ignore) {
    #     $attachments += $errorFile
    # }
    if (Test-Path -Path $withMailboxCSV -ErrorAction Ignore) {
        $attachments += $withMailboxCSV 
    }
    if (Test-Path -Path $withoutMailboxCSV -ErrorAction Ignore) {
        $attachments += $withoutMailboxCSV
    }
    
    # define domain
    if ($XMA.IsPresent) {
        $dom = 'XMA.co.uk'
    }
    elseif ($Westcoast.IsPresent) {
        $dom = 'WestCoast.co.uk'
    }

    # send report
    Send-EmailReports -recipients $recipients -emailSender $emailSender -SmtpServer $SmtpServer -attachments $attachments -dom $dom -days $daysInactive


}
    
END {
    Stop-Transcript
}

# SIG # Begin signature block
# MIIOWAYJKoZIhvcNAQcCoIIOSTCCDkUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUR8anf2p8Pa8qI0zHmD+EXxNs
# AKigggueMIIEnjCCA4agAwIBAgITTwAAAAb2JFytK6ojaAABAAAABjANBgkqhkiG
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUYazk1Z/tc62f/TH2z4bHcWDHgLEwDQYJKoZI
# hvcNAQEBBQAEggEA1HY091ajYE/auD0JtfVRxTet4JJG4+6NeTjdvCDpstMhEKmy
# r8r7rlLhKo5kJYmTmnnsF+k6PofoSt9k0E8w4jpmii/YLpSqXZISbnk03JlPkI7G
# UrlQ6i0GPZ3av1Y+2NZIw5FKiXfFweS3tDuNq25FJjRRs8tz/FGqN05EsNJ4Dobf
# /6YoSqsj74Cr37EfLoy8Xuo6fWNKwmIuW6NVBrsm9vxYqij1/oeib0WZ4uE+wWGx
# cvMauEdMbIXQtzt275UgpFTP30SQ+hiCMKBnIFH1hcUrL242giLsddo/RUUg0xkS
# ATCs7nkJ0C5EEB1CGZT+MEwQK55Ra4KJup+VHg==
# SIG # End signature block
