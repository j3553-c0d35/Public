function Create-Prompt
    {
     [CmdletBinding()]
     Param()

     $title = "Prompt Menu"
     $message = "Please choose an action..."
     $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", `
     "Performs action."
     $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", `
     "Does not perform action."
     $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
     $global:result = $host.ui.PromptForChoice($title, $message, $options, 0) 
     switch ($global:result)
        {
            0 {"You selected Yes... proceeding"}
            1 {"You selected No... exiting"}
        }
}

function Rename-Host 
    {
     [CmdletBinding()]
     Param()

     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Do you want to rename this computer?                                                                " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {     
                 $oldname = (hostname) 
                 $newname = Read-Host -Prompt "Please enter new HostName"
                 Rename-Computer -ComputerName $oldname -NewName $newname -Force -Verbose
                }     
}

function Update-PSHelp
    {
     [CmdletBinding()]
     Param()

     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Do you want to update powershell help?                                                              " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {     
                 Update-Help -Force
                }     
}

function Install-TelnetClient
    {
     [CmdletBinding()]
     Param()

     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Do you want to install Telnet Client?                                                               " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {     
                 Install-WindowsFeature -Name Telnet-Client -Verbose
                }     
}

function Enable-RDP
    {
     [CmdletBinding()]
     Param()

     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Do you want to allow RDP connections to this computer?                                              " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {     
                 set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server'-name "fDenyTSConnections" -Value 0 -Verbose
                 set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -name "UserAuthentication" -Value 0 -Verbose
                 Set-NetFirewallRule -DisplayName "Remote Desktop - User Mode (TCP-In)" -Enabled True -Profile Domain,Private -Direction Inbound -Action Allow -Verbose
                 Set-NetFirewallRule -DisplayName "Remote Desktop - User Mode (UDP-In)" -Enabled True -Profile Domain,Private -Direction Inbound -Action Allow -Verbose
                }     
}

function Enable-Ping
    {
     [CmdletBinding()]
     Param()

     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Do you want to allow inbound ipv4 ping requests to this computer?                                   " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {
                 Set-NetFirewallRule -DisplayName "File and Printer Sharing (Echo Request - ICMPv4-In)" -Enabled True -Profile Domain,Private -Direction Inbound -Action Allow -Verbose
                 Set-NetFirewallRule -DisplayName "File and Printer Sharing (Echo Request - ICMPv4-Out)" -Enabled True -Profile Domain,Private -Direction Inbound -Action Allow -Verbose
                }     
}
 

function Set-StaticIP
    {
     [CmdletBinding()]
     Param()

     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Do you want to set a static IP address on this computer?                                            " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {
                    Write-Host "Please copy and paste an adapter name from this list..." -ForegroundColor White -BackgroundColor Black
                    gwmi Win32_NetworkAdapter | Select-Object -Property Name
                 
                    $networkinterface = read-host "Enter the name of the NIC (ie Intel Local Area Connection)"
                    $ip = read-host "Enter an IP Address (ie 10.1.1.10)"
                    $mask = read-host "Enter the subnet mask (ie 255.255.255.0)"
                    $gateway = read-host "Enter the IP Address of the gateway (ie 10.1.1.1)"
                    $dns1 = read-host "Enter the first DNS Server (ie 8.8.8.8)"
                    $dns2 = ""
                    $dns2 = read-host "Enter the second DNS Server, or just press enter (ie 8.8.8.8)"
                    $registerDns = "TRUE"
              
                    $dns = $dns1
                        if($dns2)
                            {
                            $dns ="$dns1","$dns2"
                            }
                $index = (gwmi Win32_NetworkAdapter | Where-Object {$_.name -eq $networkinterface}).InterfaceIndex
                $NetInterface = Get-WmiObject Win32_NetworkAdapterConfiguration | Where-Object {$_.InterfaceIndex -eq $index}
                $NetInterface.EnableStatic($ip,$mask)
                $NetInterface.SetGateways($gateway)
                $NetInterface.SetDNSServerSearchOrder($dns)
                $NetInterface.SetDynamicDNSRegistration($registerDns)
                }
            
}

function Set-TimeZone
    {
     [CmdletBinding()]
     Param()

     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Do you want to set the timezone to AUS Central Standard Time?                                       " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {
                 tzutil.exe /s "AUS Central Standard Time"
                 if ($LASTEXITCODE -eq '0')
                        {Write-Verbose -Message "Time zone set successfully"}
                }     
}

function Set-NTP
    {
     [CmdletBinding()]
     Param()

     # Sets NTP servers to use, creates scheduled task for resync every day at 1pm
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Do you want to setup NTP?                                                                           " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {
                 Start-Service -Name W32Time -Verbose
                 w32tm /configure /manualpeerlist:"0.au.pool.ntp.org,0x1 1.au.pool.ntp.org,0x1 2.au.pool.ntp.org,0x1 3.au.pool.ntp.org,0x1" /syncfromflags:manual /update
                 w32tm /resync
                }     
}


function Enable-DefaultIcons
    {
     [CmdletBinding()]
     Param()

     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Do you want to enable default desktop icons?                                                        " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {
                 $computer = '{20D04FE0-3AEA-1069-A2D8-08002B30309D}'
                 $userfiles = '{59031a47-3f72-44a7-89c5-5595fe6b30ee}'
                 $controlpanel = '{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}'
                 $recyclebin = '{645FF040-5081-101B-9F08-00AA002F954E}'
                 $network = '{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}'
                 
                 Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel -Name $computer -Value 0 -Verbose
                 Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel -Name $userfiles -Value 0 -Verbose
                 Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel -Name $controlpanel -Value 0 -Verbose
                 Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel -Name $recyclebin -Value 0 -Verbose
                 Set-ItemProperty -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel -Name $network -Value 0 -Verbose
                }     
}

function Set-IEHomePage
    {
     [CmdletBinding()]
     Param()

     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Do you want to set default IE home page to https://www.google.com.au/ ?                             " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {
                 $homepage = 'Start Page'
                 Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Internet Explorer\Main' -Name $homepage -Value 'https://www.google.com.au/' -Verbose
                }     
}


function Disable-IEEnhancedSecurity
    {
     [CmdletBinding()]
     Param()

     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Do you want to disable IE enhanced security ?                                                       " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {
                 Set-ItemProperty -Path “HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}” -Name isinstalled -Value 0
                 Set-ItemProperty -Path “HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}” -Name isinstalled -Value 0
                 Rundll32 iesetup.dll, IEHardenLMSettings,1,True
                 Rundll32 iesetup.dll, IEHardenUser,1,True
                 Rundll32 iesetup.dll, IEHardenAdmin,1,True
                    If (Test-Path “HKCU:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}”)
                            {
                             Remove-Item -Path “HKCU:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}”
                            }
                    If (Test-Path “HKCU:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}”)
                            {
                             Remove-Item -Path “HKCU:\SOFTWARE\Microsoft\Active Setup\Installed Components\{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}”
                            }
                 Remove-ItemProperty -Path “HKCU:\SOFTWARE\Microsoft\Internet Explorer\Main” -Name “First Home Page” -ErrorAction SilentlyContinue
                }     
}

function Disable-MouseShadow
    {
     [CmdletBinding()]
     Param()

     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Do you want to disable mouse shadow ?                                                               " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {
                 Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "UserPreferencesMask" -Value ([byte[]](0x90,0x12,0x03,0x80,0x10,0x00,0x00,0x00))
                 Write-Verbose -Message "Reboot is required... you will be prompted to reboot at the end of this session."
                } 
}

function Restart-Host
    {
     [CmdletBinding()]
     Param()

     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Write-Host "Some changes require a reboot, do you want to reboot now?                                           " -ForegroundColor Black -BackgroundColor White
     Write-Host "::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::" -ForegroundColor Magenta -BackgroundColor Black
     Create-Prompt
        if ($global:result -eq '0')
                {
                 Restart-Computer -Force
                } 
}

Rename-Host -Verbose
Set-StaticIP -Verbose
Enable-Ping -Verbose
Enable-RDP -Verbose
Set-TimeZone -Verbose
Set-NTP -Verbose
Install-TelnetClient -Verbose
Enable-DefaultIcons -Verbose
Disable-IEEnhancedSecurity -Verbose
Set-IEHomePage -Verbose
Disable-MouseShadow -Verbose
Update-PSHelp -Verbose
Restart-Host -Verbose
