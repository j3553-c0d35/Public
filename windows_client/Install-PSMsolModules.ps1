# Set PSGallery to trusted
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted


# Microsoft Online Services Module
# This is the original tenant directory management tool
# Prefix: MSOL
# To update module, run Update-Module -Name MSOnline
Install-Module -Name MSOnline
# Connect with:
# Connect-MsolService


# Exchange Online Module with 2FA support
# https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps


# AzureAD
# This is the newer tenant directory management tool 
# Prefix: AzureAD
# To update module, run Update-Module -Name azuread
Install-Module -Name azuread
# Connect with:
# Connect-AzureAD


# AzureADPreview
# This is the latest tenant directory management tool in beta
# Prefix: AzureAD
# To update module, run Update-Module -Name azureadpreview
Install-Module -Name AzureADPreview
# Connect with:
# Connect-AzureAD


# SharePoint Online
# Manage SharePoint sites and related services
# Prefix: SPO
# To update module, run Update-Module -Name Microsoft.Online.SharePoint.PowerShell
Install-Module -Name Microsoft.Online.SharePoint.PowerShell
# Connect with:
# Connect-SPOService


# SharePoint Online pnp tool. Provides library of PowerShell commands that allows you to 
# perform complex provisioning and artifact management actions towards SharePoint. 
# The commands use a combination of CSOM and REST behind the scenes, and can work against 
# both SharePoint Online as SharePoint On-Premises.
# https://github.com/SharePoint/PnP-PowerShell
Install-Module SharePointPnPPowerShellOnline


# Teams
# Microsoft Teams Management Preview (0.9.6) ​
# (uses Graph beta/preview APIs)
# Prefix: Team
Install-Module -Name MicrosoftTeams -RequiredVersion 0.9.6
# Connect with:
# Connect-MicrosoftTeams


# Microsoft Teams Management (1.0.2)
# (uses 1.0 Graph API) 
# Prefix: Team 
# To update module, run Update-Module -Name MicrosoftTeams
Install-Module -Name MicrosoftTeams
# Connect with:
# Connect-MicrosoftTeams


# Skype for Business​
# (No PowerShell module install from Gallery)
# Below download path may change, if you need the latest build: https://www.microsoft.com/en-us/download/confirmation.aspx?id=39366
# $skype_save_path = ("$env:USERPROFILE" + "\Downloads\SkypeOnlinePowerShell.exe")
# Invoke-WebRequest -Uri https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowerShell.exe -OutFile $skype_save_path
# & $skype_save_path


# Standard PowerShell Exchange Online
# (No local install)
# $UserCredential = Get-Credential
# $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
# Import-PSSession $Session -DisableNameChecking


# Flow and PowerApps
Install-Module -Name Microsoft.PowerApps.PowerShell # user cmdlets
Install-Module -Name Microsoft.PowerApps.Administration.PowerShell –AllowClobber # sysadmin cmdlets


# PowerShell credential manager - provides access to Windows Cred manager from PowerShell
Install-Module -Name CredentialManager -RequiredVersion 1.0
