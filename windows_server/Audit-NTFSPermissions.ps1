$FolderPath = Get-ChildItem -Directory -Path "\\<server>\<share>" -Force
$Report = @()
Foreach ($Folder in $FolderPath) {
    $Acl = Get-Acl -Path $Folder.FullName
    foreach ($Access in $acl.Access)
        {
            $Properties = [ordered]@{
                'FolderName' = $Folder.FullName;
                'ADGroup or User' = $Access.IdentityReference;
                'Permissions' = $Access.FileSystemRights;
                'Inherited' = $Access.IsInherited
                 }
            $Report += New-Object -TypeName PSObject -Property $Properties
        }
}
$Report | Export-Csv -path "C:\Temp\FolderPermissions.csv" -NoTypeInformation -NoClobber