function Process-OffBoarding02 {
    [CmdletBinding()]
	param(## Domain selector
    [Parameter(Mandatory=$true , ParameterSetName="WestCoast")] [switch]$Westcoast,
    [Parameter(Mandatory=$true , ParameterSetName="XMA")] [switch]$XMA
  )

  #Domain-specific variables
  if ($Westcoast.IsPresent){
    # Credentials for WC
    Create-Credential -WestCoast -AD
    Create-Credential -WestCoast -AAD
    $searchBase = "OU=Leavers Pending Export,OU=Active Employees,OU=USERS,OU=WC2014,DC=westcoast,DC=co,DC=uk"
    $UserDomain = "westcoast.co.uk"
    $DomainNetBIOS = "WESTCOASTLTD"
  }
  elseif ($XMA.IsPresent){
    # Credentials for XMA
    Create-Credential -XMA -AD
    Create-Credential -XMA -AAD
    $searchBase = "OU=90 day notice user accounts,DC=xma,DC=co,DC=uk"
    $UserDomain = "xma.co.uk"
    $DomainNetBIOS = "XMA"
  }

  #Shared variables
  $DC = (Get-ADForest -Identity $UserDomain -Credential $AD_Credential |	Select-Object -ExpandProperty RootDomain |	Get-ADDomain |	Select-Object -Property PDCEmulator).PDCEmulator
  $daysInactive = 90
  $litigationHoldTime = 2555
  $date = Get-date -Format dd_MM_yyyy
  $Phase2CSVContents = @() # Create the empty array that will eventually be the CSV file

  #1 -  Connect exchange online
  Connect-OnlineExchange

  #2 - Find users older then the inactivity date set in the parameters

  $inactiveDate = (Get-Date).Adddays(-($daysInactive))

  # Get AD Users that haven't logged on in xx days and are not Service Accounts
  foreach ($sb in $searchBase){

    #Processed leavers - leaverts in the leaver OU
    $ProcessedLeavers = Get-ADUser -SearchBase $searchBase -Filter { SamAccountName -notlike "*svc*" } -Properties LastLogonDate -Server $DC -Credential $AD_Credential | Select-Object @{ Name="Username"; Expression={$_.SamAccountName} }, Name, LastLogonDate, DistinguishedName

    #Truly inactive leavers - leavers in the leaver OU whom not logged in for 90 days
    $InactiveLeavers = Get-ADUser -SearchBase $searchBase -Filter { LastLogonDate -lt $inactiveDate -and SamAccountName -notlike "*svc*" } -Properties LastLogonDate -Server $DC -Credential $AD_Credential | Select-Object @{ Name="Username"; Expression={$_.SamAccountName} }, Name, LastLogonDate, DistinguishedName

    foreach ($iu in $InactiveLeavers)
    {
    #REPORT
    $row = New-Object System.Object # Create an object to append to the array
    $row | Add-Member -Force -MemberType NoteProperty -Name "Domain" -Value $DomainNetBIOS
    $row | Add-Member -Force -MemberType NoteProperty -Name "InactiveUser" -Value $($iu.Username) # create a property called InactiveUser. This will be the User column

      #3 - Find the mailbox of each of these users, if it exists
      $eapCurrent = $ErrorActionPreference
      $ErrorActionPreference = "silentlycontinue" # turning off warning messages for manual run
        $mbx = Get-Mailbox ( (Get-ADUser -Identity $iu.Username -Server $DC -Credential $AD_Credential).UserPrincipalName )
      $ErrorActionPreference = $eapCurrent # turning on warning messages

      #4 - If it exits, puts the mailbox into litigation hold
      if ($mbx){
        $mbx | Set-Mailbox -LitigationHoldEnabled $true -LitigationHoldDuration $litigationHoldTime -WhatIf
        # REPORT
        $row | Add-Member -Force -MemberType NoteProperty -Name "SetToLitigation" -Value "Yes"
      }
      #TODO: Add check for litigation enabled

      #5 - If there is no mailbox, only note this to the report
      else {
      # REPORT
      $row | Add-Member -Force -MemberType NoteProperty -Name "SetToLitigation" -Value "n/a"
      }

      #6 - Finally remove the AD object
      Remove-ADUser -Identity $iu.DistinguishedName -Confirm:$false -Server $DC -Credential $AD_Credential -WhatIf
      # REPORT
      $row | Add-Member -Force -MemberType NoteProperty -Name "ADAccountDeleted" -Value "Yes"
      Write-Output "$($iu.Username) - Deleted"

      # REPORT
      # For Each user
      $Phase2CSVContents += $row # append the new data to the array
    }

    # Populate CSV and export it
    $CSVExport02 = ".\" + $OutputFolder + "\" +$DomainNetBIOS + "_" +  "InactiveAccounts_" + (Get-Date -Format d_M_yyyy) + ".csv"
    $Phase2CSVContents | Export-csv -Path $CSVExport02 -NoTypeInformation -Force # first add in the original import
      }

    # Disconnect from o365
    Get-Pssession | Remove-PSSession # get rid of the no-longer-needed session

}
# SIG # Begin signature block
# MIIOWAYJKoZIhvcNAQcCoIIOSTCCDkUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU9wSbNZn+DM0XAinUTUj4IjjM
# 3CSgggueMIIEnjCCA4agAwIBAgITTwAAAAb2JFytK6ojaAABAAAABjANBgkqhkiG
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUUAnaCWDCLuUJkZmSm2p5/HNIIBAwDQYJKoZI
# hvcNAQEBBQAEggEAqE93OzyJAXWm6K/eEzm6MgxUyV4Cbu3twkAjN/LurGsykG3P
# AKIuOvRdSm8JZJnbdraWL3FTounXHhqOJw8fKxf9oiFtk90bml0EshUDZ9wI6LZz
# WabdcJvBUkrJJdAl47mfGOhfXf5mKwbAx6vMV0brlg/tRxXsdFCOLBKLpoplrKh2
# arwrRM8iB06FekpGDT/SRbU82xVd8wYOu4pNsYlZVoS9VmXPM56sugJYWnHLFjXO
# kHiFUt9nNn7kayaBtBdu2s1/pfYC0VaDJ3oxME0Dku/FnIQE5UxVT04lpcwHCdF0
# 5u+HY6IEzqaOx/9LfQmIiFtgUqPZ+C227TzddA==
# SIG # End signature block
