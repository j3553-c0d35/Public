#Set PowerShell execution policy
Set-ExecutionPolicy RemoteSigned


function New-WSLConfiguration
{
<#
.Synopsis
   Confiugre a workstation with WSL
.DESCRIPTION
   Downloads, installs and configures Windows Subsystem for Linux and Ubuntu 16.04. Optionally specify a download path to save the Ubuntu installer to, by default it will be the users download folder.
.EXAMPLE
   New-WSLConfiguration -DownloadPath C:\users\<user name>\Downloads
#>
    [CmdletBinding()]
    Param
    (
        # Path to save downloaded files to, eg. C:\users\<user name>\Downloads
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$DownloadPath = ("$env:USERPROFILE" + "\Downloads\ubuntu1604.appx")
    )

    Begin
    {
        # Check if WSL is already installed, if so skip
        if ((Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux).state -eq "Enabled")
                {
                    $WSLEnabled = $true
                }
    }
    Process
    {
        try
            {
                if ($WSLEnabled)
                {
                    Write-Host "WSL Already installed, skipping..."
                }
                else { # Install WSL
                    Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux -NoRestart -ErrorAction Stop -Verbose
                }
            }
        catch
            {
                Write-Host "Failed to install WSL..."
            }
        
        try {
                # Download Ubuntu for use with WSL
                Invoke-WebRequest -Uri https://aka.ms/wsl-ubuntu-1604 -OutFile  $DownloadPath -UseBasicParsing -Verbose
            }
        catch 
            {
                Write-Host "Failed to download Ubuntu Appx package..."
            }
        
        try {
                # Install Ubuntu for use with WSL
                Add-AppxPackage -Path $DownloadPath -ErrorAction Stop -Verbose
            }
        catch 
            {
                Write-Host "Failed to install Ubuntu Appx package"
            }


New-WSLConfiguration

<#
Tested working on Windows 10 Pro 1903

TO DO
- After WSL feature is installed, reboot is required
- After 1604 is intalled, it needs to be launched to configure it, it prompts for:
 Enter new UNIX username: 
 Enter new UNIX password:
- Ubuntu first run, setup user, SSH key, symlink to ssh config
- Install standard linux tools that are needed (including virt-manager)
- Install and configure vcxsrv (via choco below)
- Integrate with conemu
- Disable tightvnc server
- Install vscode WSL extension
- Configure default vscode settings


Conemu:
- create startup tasks file
- point conemu to the tasks file
- copy settings xml from pre-setup instance

#>

}


# Turn on Telnet Client
Enable-WindowsOptionalFeature -Online -FeatureName TelnetClient



# Install Choco
$ChocoPath = "C:\ProgramData\chocolatey\choco.exe"
    if (Test-Path -Path $ChocoPath) {
        Write-Host "Choco already exists, continuing..."
    }
    else {
        Write-Host "Installing Choco"
        Set-ExecutionPolicy Bypass -Scope Process -Force; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    }
# Tested on Win 10 Pro 1903


# Create lists of apps (required, optional)
$apps_required= @(
    "7zip",
    "azcopy", # Azcopy version 8.1.0, needed for PST uploads using SAS URLs
    "azcopy10", # Azcopy version 10.x, needed for newer on-prem to Azure copy operations
    "azure-cli", # Azure cli tool used for scripting, querying data, supporting long-running operations etc.
    "conemu", # Windows console emulator, drop down, multiple consoles, themes, lots of development
    "notepadplusplus",
    "git",
    "vscode", # Primary IDE that most devs are using these days
    "vscode-icons", # Enhanced vscode icon sets
    "vscode-powershell", # Native PowerShell support for vscode
    "pwsh", # PowerShell core 6.x
    "mremoteng", # Windows RDP client
    "adobereader",
    "dia", # Diagram creation tool (similar to Visio)
    "filezilla", # FTP/S Client
    "teamviewer", # Remote access client
    "openvpn", # VPN client
    "vmware-powercli-psmodule", # VMWare PowerShell cmdlets
    "vmwarevsphereclient", # Latest VMWare vsphere client (6.x)
    "tightvnc", # VNC Viewer
    "putty", # Console emulation tool
    "vcxsrv", # Windows X Server for forwarding X windows from WSL
    "windows-adk", # Windows Asessment and Deployment Toolkit (Build WinPE ISO's etc.)
    "mdt", # Microsoft Deployment Toolkit, used for creating task sequences for depling SOE's etc.
    "firefox", # web browser
    "nagstamon", # nagios monitoring tool
    "msoidcli" # Sign-in assistant. Neeeded for MSOL and Azure AD modules
    "nagstamon" # Nagios monitoring tool.
)

# confirmed all working on Win10 Pro 1903

$apps_optional= @(
    "atom", # Open source text editor
    "vlc", # media player
    "chromium", # additional browser
    "spotify", # online steaming music player
    "gimp" # GNU Image manipulation program
    "pdfill" # Ghost script + gui front end for joining, splitting, rotating, formatting PDF files 
)

# confirmed all working on Win10 Pro 1903

# To upgrade choco itself, run:
# choco upgrade chocolatey

# To upgrade packages, run:
# choco upgrade <package_name>

# To upgrade all packages, run:
# choco upgrade all

# Install required apps
foreach ($app in $apps_required) {
    choco install $app -y
}

# Install option if install_optional flag used
if ($InstallOptionalApps.IsPresent) {
    foreach ($app in $apps_optional) {
        choco install $app -y
    }
}
