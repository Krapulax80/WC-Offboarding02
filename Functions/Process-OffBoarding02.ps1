function Process-OffBoarding02 {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true , ParameterSetName = "WestCoast")] [switch]$Westcoast,
    [Parameter(Mandatory = $true , ParameterSetName = "XMA")] [switch]$XMA,
    [Parameter(Mandatory = $false)] [switch]$Summary,
    [Parameter(Mandatory = $false)] [switch]$Live,
    [Parameter(Mandatory = $false)] [switch]$Dryrun,
    [Parameter(Mandatory = $false)] [switch]$First10,
    [Parameter(Mandatory = $false)] [string]$OutputFolder,
    [Parameter(Mandatory = $false)] [PSCustomObject] $config,
    [Parameter(Mandatory = $false)] [string] $credfolder
  )
  
  begin {


    #Configuration definition
    $searchBase = $config.searchbase
    $systemDomain = $config.systemdomain
    $DomainNetBIOS = $config.domainnetbios
    $infraserver = $config.$infraserver#"BNWINFRATS01.westcoast.co.uk"

    if ($Westcoast.IsPresent) {
      # Credentials for WC
      $AD_Credential = Create-Credential -WestCoast -AD -CredFolder $credfolder -AD_Admin $($config.AD_admin) #"\\$infraserver\c$\Scripts\AD\OffBoarding\Credentials\"
      $AAD_Credential = Create-Credential -WestCoast -AAD -CredFolder $credfolder -AAD_Admin $($config.AAD_admin) #"\\$infraserver\c$\Scripts\AD\OffBoarding\Credentials\"
      # Config file
      #$config = Import-Csv "\\$infraserver\c$\Scripts\AD\OffBoarding\config\westcoast.csv"
    }
    elseif ($XMA.IsPresent) {
      # Credential for XMA
      $AD_Credential = Create-Credential -XMA -AD -CredFolder $credfolder -AD_Admin $($config.AD_admin) #"\\$infraserver\c$\Scripts\AD\OffBoarding\Credentials\"
      $AAD_Credential = Create-Credential -XMA -AAD -CredFolder $credfolder -AAD_Admin $($config.AAD_admin) #"\\$infraserver\c$\Scripts\AD\OffBoarding\Credentials\"
      # Config file
      #$config = Import-Csv "\\$infraserver\c$\Scripts\AD\OffBoarding\config\xma.csv"
    }

    $DC = (Get-ADForest -Identity $systemDomain -Credential $AD_Credential |	Select-Object -ExpandProperty RootDomain |	Get-ADDomain |	Select-Object -Property PDCEmulator).PDCEmulator

    $daysInactive = 90
    $litigationHoldTime = 2555
    $date = Get-date -Format dd_MM_yyyy
    $inactiveDate = (Get-Date).Adddays( - ($daysInactive))

    $InactiveLeavers = $UsersWithMailbox = $UsersWithoutMailbox = $null
    $UsersWithMailbox = @()
    $UsersWithoutMailbox = @()

    Get-PSSession | Remove-PSSession
    Connect-OnlineExchange -AAD_Credential $AAD_Credential
  }
  
  process {
    #1. Collect all inactive users (not active for 90 days or never logged in
    if ($First10.IsPresent) {
      $number = 10
      Write-Host # lazy line break
      Write-Host "WARNING (!) You are in 'First10' mode. Only the first $number results are returned. For real work ommit the '-First10' switch !" -ForegroundColor Yellow
      $InactiveLeavers = Get-ADUser -Credential $AD_Credential -Server $DC -Filter * -SearchBase $searchBase -Properties *  | Where-Object { ($null -EQ $_.LastLogonDate) -or ($_.LastLogonDate -lt $inactiveDate) -and ($_.Name -notlike 'svc.*') } | Select-Object Name, EmailAddress, LastLogonDate, whenCreated, UserPrincipalName, SAMAccountName | Sort-Object whenCreated | Select-Object -First $number
    }
    else {
      $InactiveLeavers = Get-ADUser -Credential $AD_Credential -Server $DC -Filter * -SearchBase $searchBase -Properties *  | Where-Object { ($null -EQ $_.LastLogonDate) -or ($_.LastLogonDate -lt $inactiveDate) -and ($_.Name -notlike 'svc.*') } | Select-Object Name, EmailAddress, LastLogonDate, whenCreated, UserPrincipalName, SAMAccountName | Sort-Object whenCreated
    }

    #2. Separate leavers WITH and WITHOUT mailbox to work with
    foreach ($Leaver in $InactiveLeavers) {
      $obj = $null
      $Obj = $null ; $Obj = New-Object -TypeName PSObject
  
      #Actual mailbox name
      $MailboxAddress = $null
      $MailboxAddress = $Leaver.UserPrincipalName
      $ADAccount = $Leaver.SAMAccountName
  
  
      #Users with mailbox
      if (Get-Mailbox $MailboxAddress -ErrorAction SilentlyContinue) {
        #Report on screen
        Write-Host "Mailbox [$MailboxAddress] found" -ForegroundColor Green
        #Report in array
        $details = Get-Mailbox $MailboxAddress | Select-Object Name, *type*
        $Obj | Add-Member -MemberType NoteProperty -Name AccountName -Value $ADAccount
        $Obj | Add-Member -MemberType NoteProperty -Name MailboxName -Value $($details.Name)
        $Obj | Add-Member -MemberType NoteProperty -Name Type -Value $($details.RecipientTypeDetails)
        $Obj | Add-Member -MemberType NoteProperty -Name EmailAddress -Value $MailboxAddress
        $UsersWithMailbox += $obj
  
        #Users without mailbox
      }
      else {
        #Report on screen
        Write-Host "Mailbox [$MailboxAddress] not found" -ForegroundColor Red
        Write-Host "Mailbox [$MailboxAddress] found" -ForegroundColor Green
        #Report in array
        $Obj | Add-Member -MemberType NoteProperty -Name AccountName -Value $ADAccount
        $Obj | Add-Member -MemberType NoteProperty -Name MailboxName -Value "n/a"
        $Obj | Add-Member -MemberType NoteProperty -Name Type -Value "n/a"
        $Obj | Add-Member -MemberType NoteProperty -Name EmailAddress -Value "n/a"
        $UsersWithoutMailbox += $obj
      }
    }

    #3. Do something

    ## LITIGATION HOLD
    # ACTION - DRYRUN
    if ($Dryrun.IsPresent) {
      Write-Host # separator line
      Write-Host "Dry run started" -ForegroundColor Yellow
      # Put the mailbox into litigation hold
      foreach ($mbx in $UsersWithMailbox) {
        Set-Mailbox -Identity $($mbx.EmailAddress) -LitigationHoldEnabled $true -LitigationHoldDuration $litigationHoldTime -WhatIf
      }
    }
    # ACTION - LIVERUN
    elseif ($Live.IsPresent) {
      Write-Host # separator line
      Write-Host "Live run started" -ForegroundColor Yellow
      foreach ($mbx in $UsersWithMailbox) {
        # Set mailbox to litigation
        Set-Mailbox -Identity $($mbx.EmailAddress) -LitigationHoldEnabled $true -LitigationHoldDuration $litigationHoldTime #-WhatIf
        # Wait for litigation to take effect
        do {
          $litigationHoldStatus = $null
          $litigationHoldStatus = (Get-mailbox -Identity $($mbx.EmailAddress)).LitigationHoldEnabled
          if ($litigationHoldStatus -ne "True") {
            Write-Output "Litigation hold is not yet enabled on [$($mbx.EmailAddress)]. Retry in 10 seconds"
            Start-Sleep 10
          }
          else { Write-Output "Mailbox [$($mbx.EmailAddress)] is in litigation hold." }
        } until ($litigationHoldStatus -eq "True")

      }

      Write-Host # separator line
      foreach ($mbx in $UsersWithMailbox) {
        # Finally delete the account
        Remove-ADUser -Identity $($mbx.AccountName) -Server $DC -Credential $AD_Credential -WhatIf #-Confirm:$false
      }

    }

    # ACTION - SUMMARY
    if ($Summary.IsPresent) {
      # First10 MODE - display calculation for check
      $today = Get-date -Format ddMM
      Write-Host # lazy line break
      Write-Host "The number of leavers with mailboxes: $($UsersWithMailbox.count)" -ForegroundColor Magenta ; $UsersWithMailbox | Export-Csv -Path ($OutputFolder + "\" + $today + "_" + $DomainNetBIOS + "_leavers_WITH_mailbox.csv" ) -Force
      Write-Host "The number of leavers without mailboxes: $($UsersWithoutMailbox.count)" -ForegroundColor Magenta -NoNewLine ; $UsersWithoutMailbox | Export-Csv -Path ($OutputFolder + "\" + $today + "_" + $DomainNetBIOS + "_leavers_WITHOUT_mailbox.csv" ) -Force 
      Write-Host " (This should be equal to $($InactiveLeavers.Count) )"
    }

  }
  
  end {
    
  }
}

# SIG # Begin signature block
# MIIOWAYJKoZIhvcNAQcCoIIOSTCCDkUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUoyATosKiWY/lQaa2+HNyG9w/
# 6+SgggueMIIEnjCCA4agAwIBAgITTwAAAAb2JFytK6ojaAABAAAABjANBgkqhkiG
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUoy6nf2VTs1sL8TVj4e40EFf9cJQwDQYJKoZI
# hvcNAQEBBQAEggEAiEd6pXFsfLCyfnukVJI6pUqvfZoOQRjWMATrlk3xq6YmVsUd
# auDGX8OT/VQ51z8y/D6u5Y4zdq91jLCKn9d0t9RBgZr2huCB6vtqmrslmDPVv0T0
# q8VNHZ7063UaKSOldmIk09WUAj/tuy6oV8A7gaptSqAtM3yCcoYsuS4BA6JWMBDE
# EAiXObg5nurnXH4kfZ7mDjnTb6F/msfCJBF0MJiAWcUi7IdVatx7WsRQrDqzyyUB
# Lea5iySyURt3I6kvxMGOXzBHFwLSrHeOkm8ZraL3eVCsii23owyMHLQ1M1M3O3Fx
# dYNHqSckICnTXM2t0Q6p096Kk//7mBAskKKvkA==
# SIG # End signature block
