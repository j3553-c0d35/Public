# Changed AD User Surname
# Based on .local tld with firstname.surname user naming convention
# Allows to choose a domain in multi-domain environments
# Sets SIP address for lync environments
# Turns on exchange email address policy so that email address is updated

Import-Module -Name activedirectory

$Domain = (($Domain = Read-Host "Enter Domain that user is a member of, eg. <dom1, dom2, dom3>") + '.local')

$GivenName = (($GivenName = Read-Host "Enter Customers GivenName..........").Substring(0,1).toupper()+$GivenName.Substring(1).tolower())
$OldSurName = (($OldSurName = Read-Host "Enter Customers 'OLD' SurName......").Substring(0,1).toupper()+$OldSurName.Substring(1).tolower())
$NewSurName = (($NewSurName = Read-Host "Enter Customers 'NEW' SurName......").Substring(0,1).toupper()+$NewSurName.Substring(1).tolower())
$OldSamAccountName = (("$GivenName" + "." + "$OldSurName").ToLower())
$NewSamAccountName = (("$GivenName" + "." + "$NewSurName").ToLower())
$NewDisplayName = ("$GivenName" + " " + "$NewSurName")

if ($Domain -eq '<dom1>') {$UserPrincipalName = ("$NewSamAccountName" + '@dom1.local')}
if ($Domain -eq '<dom2>') {$UserPrincipalName = ("$NewSamAccountName" + '@dom2.local')}
if ($Domain -eq '<dom3>') {$UserPrincipalName = ("$NewSamAccountName" + '@dom3.local')}

$SIP = ("sip:" + "$UserPrincipalName")


Write-Host "Proceed with name change ?"
pause

$User_Name_To_Change = (Get-ADUser -Server $Domain -Identity $OldSamAccountName -Properties DisplayName,SamAccountName,Surname,UserPrincipalName)
$User_Name_To_Change | Set-ADUser -Server $Domain -DisplayName $NewDisplayName -SamAccountName $NewSamAccountName -Surname $NewSurName -UserPrincipalName $UserPrincipalName -EmailAddress $UserPrincipalName -Replace @{mail="$UserPrincipalName";mailNickname="$NewSamAccountName";"msRTCSIP-PrimaryUserAddress"="$SIP"}
$User_Name_To_Change | Rename-ADObject -Server $Domain -NewName $NewDisplayName


Get-mailbox -Identity $NewSamAccountName -Properties -EmailAddressPolicyEnabled