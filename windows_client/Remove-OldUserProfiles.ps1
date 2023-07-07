$ProfileKeys = (Get-ChildItem -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList').name

$FullProfileKeys = @()
foreach ($key in $ProfileKeys) {$FullProfileKeys += $($key -replace "HKEY_LOCAL_MACHINE","HKLM:")}

$KeysToDelete = @()
foreach ($subkey in $FullProfileKeys) 
            {
            $KeysToDelete += $(Get-ItemProperty -Path $subkey | Where-Object {(($_.profileimagepath -notmatch 'systemprofile'))
                                                                          -or (($_.profileimagepath -notmatch 'LocalService'))
                                                                          -or ($_.profileimagepath -notmatch 'NetworkService') 
                                                                          -or ($_.profileimagepath -notmatch '<keep this profile>') 
                                                                          -or ($_.profileimagepath -notmatch 'administrator')})
            }

foreach ($DelKey in $KeysToDelete) {Remove-Item -Recurse -Path $DelKey.PSPath }

New-Item -Type Directory "C:\Users\Backups"
foreach ($Path in $KeysToDelete) {Move-Item $Path.ProfileImagePath "C:\Users\Backups" -Force}

Write-Output $env:COMPUTERNAME >> C:\profiles_done.txt
