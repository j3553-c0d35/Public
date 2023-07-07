Install-WindowsFeature -Name ad-domain-services,RSAT-AD-Tools,RSAT-AD-PowerShell,RSAT-ADDS, DNS, RSAT-DNS-Server

$smp = ConvertTo-SecureString -String "<PASSWORD>" -AsPlainText -Force


Import-Module ADDSDeployment
Install-ADDSForest `
-CreateDnsDelegation:$false `
-DatabasePath "C:\Windows\NTDS" `
-DomainMode "WinThreshold" `
-DomainName "lab.local" `
-DomainNetbiosName "lab" `
-ForestMode "WinThreshold" `
-InstallDns:$true `
-LogPath "C:\Windows\NTDS" `
-NoRebootOnCompletion:$false `
-SysvolPath "C:\Windows\SYSVOL" `
-SafeModeAdministratorPassword $smp `
-Force:$true

#DNS
Import-Module DnsServer

#Set dns forwarded
Set-DnsServerForwarder -IPAddress 1.1.1.1 -ComputerName "<NEW DC HOSTNAME>"

#Set dns listen interfaces
dnscmd "<NEW DC HOSTNAME>" /ResetListenAddresses "<NEW DC IP>"

#Create reverse zone
Add-DnsServerPrimaryZone -NetworkID "<0.0.0.0/0>" -ReplicationScope "Forest"

#Add PTR for server
Add-DnsServerResourceRecordPtr -Name 10 -PtrDomainName "<NEW DC FQDN>" -ZoneName "<0.0.0>.in-addr.arpa" -AgeRecord -AllowUpdateAny