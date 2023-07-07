$Name = '<vm name>'
$SwitchName = '<vSwitch name>'
$HardDiskSize = <HDD size in GB>
$HDPath = '<path\to\vhdx.vhdx>'
$RAM = '<RAM in GB>'
$Generation = '<1|2>'
$ISO_Path = '<path\to\iso.iso>'

New-VM -Name $Name -SwitchName $SwitchName `
-NewVHDSizeBytes $HardDiskSize `
-NewVHDPath $HDPath -Generation $Generation -MemoryStartupBytes $RAM

Add-VMDvdDrive -VMName $Name -Path $ISO_Path