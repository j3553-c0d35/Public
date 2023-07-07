Import-Module -Name ActiveDirectory

$all_computers = Get-ADComputer -Filter * -SearchBase 'OU=computers,DC=somedomain,DC=com' | Select-Object -Property Name

$ExportPath = "$env:TEMP\$(Get-date -Format 'yyyyMMddhhmmss')_workstation_os_build_report.csv"

foreach ($c in $all_computers.name)

{

    if (Test-Connection -ComputerName $c -count 1 -Quiet ) {

    Write-Host "Processing $c" -ForegroundColor Cyan

    $CurrentBuild = ""
    $UBR = ""
    $OSVersion = ""
    $ComputerSystem = ""
    $props = ""
    $obj = ""

        $CurrentBuild = Invoke-Command -ComputerName $c -ScriptBlock { (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' –Name CurrentBuild).CurrentBuild } -ErrorAction SilentlyContinue
        $UBR = Invoke-Command -ComputerName $c -ScriptBlock { (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' –Name UBR).UBR } -ErrorAction SilentlyContinue
        $OSVersion = $CurrentBuild + "." + $UBR
        $ComputerSystem = Get-WmiObject -ComputerName $c -Class Win32_ComputerSystem -ErrorAction SilentlyContinue

       $props = [ordered]@{ 
        'HostName' = $ComputerSystem.Name;
        'OSVerion' = $OSVersion
        }
        $obj = New-Object -TypeName PSObject -Property $props

    Write-Output $obj | Export-Csv -Path $ExportPath -NoTypeInformation -Append -NoClobber -Force
    }
    else {
        Write-Host "$c is offline..." -ForegroundColor Green
    }

}
Write-Host "Output csv file is located here: `n `n $ExportPath `n" -ForegroundColor Yellow