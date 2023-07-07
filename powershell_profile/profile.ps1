function Start-PSEXOnline {
    $pwsh_exch_online_module = "C:\Users\<username>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Corporation\Microsoft Exchange Online Powershell Module.appref-ms"
    explorer.exe $pwsh_exch_online_module
    }
    
    
    function Add-CustomStaticRoute {
    # Grab OpenVPN tunnel interface index
    $ifIndex = (Get-NetAdapter | Where-Object {$_.InterfaceDescription -eq "TAP-Windows Adapter V9"}).ifIndex
    # Prompt for dest and hop
    $Dest = Read-Host -Prompt "What is the destination? eg. <x.x.x.x>/32"
    $Hop = Read-Host -Prompt "What is the gateway you want to route via? eg. <x.x.x.x>"
    
    New-NetRoute -DestinationPrefix $Dest -InterfaceIndex $ifIndex -NextHop $Hop
    
    }

    function Search-Internet {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory=$True,Position=0)]
            [String]$SearchFor,
            
            [Parameter(Mandatory=$True,Position=1)]
            [String]$Use
        )
        
        $ErrorActionPreference = "SilentlyContinue"
        If ($Error) {$Error.Clear()}
        $SearchFor = $SearchFor.Trim()
        If (!($SearchFor)) {
            Write-Host
            Write-Host "Text That You Wish To Search For Has Not Been Entered." -ForeGroundColor "Yellow"
            Write-Host "Execution of the Script Has been Ternimated." -ForeGroundColor "Yellow"
            Write-Host
            Exit
        }
        $Use = $Use.Trim()
        If (!($Use)) {
            Write-Host
            Write-Host "Search Engine To Use Has Not Been Specified." -ForeGroundColor "Yellow"
            Write-Host "Execution of the Script Has been Ternimated." -ForeGroundColor "Yellow"
            Write-Host
            Exit
        }
        $SearchFor = $SearchFor -Replace "\s+", " "
        $SearchFor = $SearchFor -Replace " ", "+"
        
        Switch ($Use) {
            "Google" {
                # -- "Use Google To Search"
                $Query = "https://www.google.com.au/search?q=$SearchFor"
            }
            "Bing" {
                # -- "Use Bing Search Engine To Search"
                $Query = "http://www.bing.com/search?q=$SearchFor"
            }
            Default {$Query = "No Search Engine Specified"}
        }
        If ($Query -NE "No Search Engine Specified") {
            ## -- Detect the Default Web Browser
            Start-Process $Query
        }
        Else {
            Write-Host
            Write-Host $Query -ForeGroundColor "Yellow"
            Write-Host "Execution of the Script Has been Ternimated." -ForeGroundColor "Yellow"
            Write-Host
        }
        }

        
function Get-RandomPassword {
                Invoke-WebRequest -Uri https://www.dinopass.com/password/strong | Select-Object -ExpandProperty content
    
}