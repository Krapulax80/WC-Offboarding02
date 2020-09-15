<#
  Usage examples:

  #Update Westcoat Active directory password:
  Create-Credential -WestCoast -AD -PasswordUpdate

  Use XMA AAD password:
  #>
function Create-Credential {
  [CmdletBinding()]
  Param (
    [Parameter(ParameterSetName = 'WestCoast ParamSet 1')] # Parameter set for WestCoast
    [Switch]
    $WestCoast,
    [Parameter(ParameterSetName = 'XMA ParamSet 2')] # Parameter set for XMA
    [Switch]
    $XMA,
    [Parameter(Mandatory = $false)]
    [Switch]
    $PasswordUpdate, # if this switch is used, the function is in "Update" mode (this is to save / update passwords)
    [switch]
    $AD, # AD credentials
    [switch]
    $AAD, # O365 / AAD credentials
    [switch]
    $Exchange, # Exchange credentials
    [string]
    $CredFolder,
    $AD_Admin,
    $AAD_Admin,
    $Exchange_Admin

  )

  # Westcoast
  if ($WestCoast.IsPresent) {
    $domain = "WC"
    $AD_CredentialFile = $CredFolder + "\" + $domain + "_AD_credential.txt"
    $AAD_CredentialFile = $CredFolder + "\" + $domain + "_AAD_credential.txt"
    $Exchange_CredentialFile = $CredFolder + "\" + $domain + "_Exchange_credential.txt"
    if ($AD.IsPresent) {
      # If specified, update AD password for Westcoast
      if ($PasswordUpdate.IsPresent) {
        read-host "Please enter password for [$AD_Admin]" -assecurestring | convertfrom-securestring | out-file $AD_CredentialFile
      }
      # Use AD password for Westcoast
      $AD_Password = Get-Content $AD_CredentialFile | ConvertTo-SecureString
      # Create the AD credential for Westcoast
      $AD_Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $AD_Admin, $AD_Password
      return $AD_Credential
    }
    elseif ($AAD.IsPresent) {
      # If specified, update the AAD password for Westcoast
      if ($PasswordUpdate.IsPresent) {
        read-host "Please enter password for [$AAD_Admin]" -assecurestring | convertfrom-securestring | out-file $AAD_CredentialFile
      }
      # Use AAD password for Westcoast
      $AAD_Password = Get-Content $AAD_CredentialFile | ConvertTo-SecureString
      # Create the AAD credential for Westcoast
      $AAD_Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $AAD_Admin, $AAD_Password
      return $AAD_Credential
    }
    elseif ($Exchange.IsPresent) {
      # If specified, update the AAD password for Westcoast
      if ($PasswordUpdate.IsPresent) {
        read-host "Please enter password for [$Exchange_Admin]" -assecurestring | convertfrom-securestring | out-file $Exchange_CredentialFile
      }
      # Use AAD password for Westcoast
      $Exchange_Password = Get-Content $Exchange_CredentialFile | ConvertTo-SecureString
      # Create the AAD credential for Westcoast
      $Exchange_Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $Exchange_Admin, $Exchange_Password
      return $Exchange_Credential
    }
  }
  # XMA
  elseif ($XMA.IsPresent) {
    $domain = "XMA"
    $AD_CredentialFile = $CredFolder + "\" + $domain + "_AD_credential.txt"
    $AAD_CredentialFile = $CredFolder + "\" + $domain + "_AAD_credential.txt"
    $Exchange_CredentialFile = $CredFolder + "\" + $domain + "_Exchange_credential.txt"
    if ($AD.IsPresent) {
      # If specified, update AD password for XMA
      if ($PasswordUpdate.IsPresent) {
        read-host "Please enter password for [$AD_Admin]" -assecurestring | convertfrom-securestring | out-file $AD_CredentialFile
      }
      # Use AD password for XMA
      $AD_Password = Get-Content $AD_CredentialFile | ConvertTo-SecureString
      # Create AD credential for XMA
      $AD_Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $AD_Admin, $AD_Password
      return $AD_Credential
    }
    elseif ($AAD.IsPresent) {
      # If specified, update AAD password for XMA
      if ($PasswordUpdate.IsPresent) {
        read-host "Please enter password for [$AAD_Admin]" -assecurestring | convertfrom-securestring | out-file $AAD_CredentialFile
      }
      # Use AAD password for XMA
      $AAD_Password = Get-Content $AAD_CredentialFile | ConvertTo-SecureString
      # Create AAD credential for XMA
      $AAD_Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $AAD_Admin, $AAD_Password # this is the AAD credential
      return $AAD_Credential
    }
    elseif ($Exchange.IsPresent) {
      # If specified, update the AAD password for Westcoast
      if ($PasswordUpdate.IsPresent) {
        read-host "Please enter password for [$Exchange_Admin]" -assecurestring | convertfrom-securestring | out-file $Exchange_CredentialFile
      }
      # Use AAD password for Westcoast
      $Exchange_Password = Get-Content $Exchange_CredentialFile | ConvertTo-SecureString
      # Create the AAD credential for Westcoast
      $Exchange_Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $Exchange_Admin, $Exchange_Password
      return $Exchange_Credential
    }
  }
  # Incorrect domain selection
  else {
    Write-Host "Correct domain was not selected, exiting";
    Break
  }
}

# SIG # Begin signature block
# MIIOWAYJKoZIhvcNAQcCoIIOSTCCDkUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUlrrHA/g2wL0V2qZERwoRjIle
# iWmgggueMIIEnjCCA4agAwIBAgITTwAAAAb2JFytK6ojaAABAAAABjANBgkqhkiG
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
# NwIBFTAjBgkqhkiG9w0BCQQxFgQUb0KfJXcelpUGSjUVbRgeI/CYTM4wDQYJKoZI
# hvcNAQEBBQAEggEAN+0oPaTky5Hh/69k9kK/ivEP6FblSoNxdn3tYcJZZ9fOH+6r
# cDyuaEPfpBzKoUEEDA6q3dm5+wjg8kKD6lPFERRIpC3e4p2AR5Ms/FuQPnYSLIpw
# qa2PrWldQblFuG4i1ifB2sDliD5oBBG1IgDEm2/p7ttC1BsLparDanqyEfjMIdNZ
# s22MF/rsRmYFBJBS2O0O0dnH2bC9dPMMHk/vxQflgMBbRl+tpEkBbvYcPC8SfWxl
# O5Vrw+ka4rHYY32Cr5TFSDx4xbG/1nkG8DGo9IcNqOthR34BfwGpwKXNbki9Sa9P
# 3RfKXJO0OWqGoCDMCck5MUzSC753kzBuJR5RtQ==
# SIG # End signature block
