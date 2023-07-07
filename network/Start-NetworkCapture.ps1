# Start capture
# Start-NetworkCapture.ps1
function startCapture {
    param ()
netsh trace start capture=yes report=disabled
}
function stopCapture {
    param ()
netsh trace stop
}

function convertCapture {
    param ()
# Setup function variables
$ETLToolZipPath = "$env:USERPROFILE\Downloads\etl2pcapng.zip"
$CapturePath = "$env:LOCALAPPDATA\Temp\NetTraces\NetTrace.etl"
$PCAPNGPath = "$env:LOCALAPPDATA\Temp\NetTraces\$(Get-Date -Format dd-MM-yyy-hhmm)_NetTrace.pcapng"
$ETLToolPath = "$env:USERPROFILE\Downloads\etl2pcapng\x64\"    

if ($ETLToolPath) {
    Write-Host "Tool already exists, skipping..."
}
else {
# Get Windows ETL to PCAPNG tool
Invoke-WebRequest -Uri https://github.com/microsoft/etl2pcapng/releases/download/v1.4.0/etl2pcapng.zip -OutFile $ETLToolZipPath -UseBasicParsing 

# Unzip the archive
Expand-Archive -Path $ETLToolZipPath -DestinationPath $env:USERPROFILE\Downloads\ -Force
}

# Run conversion
Set-Location -Path $ETLToolPath
.\etl2pcapng.exe $CapturePath $PCAPNGPath

Write-Output -InputObject "Capture has been converted and can be found here: $PCAPNGPath"
}

