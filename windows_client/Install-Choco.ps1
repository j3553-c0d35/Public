$ChocoPath = "C:\ProgramData\chocolatey\choco.exe"

    if (Test-Path -Path $ChocoPath) {
        Write-Host "Choco already exists, continuing..."
    }
    else {
        Write-Host "Installing Choco"
        Set-ExecutionPolicy Bypass -Scope Process -Force; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    }