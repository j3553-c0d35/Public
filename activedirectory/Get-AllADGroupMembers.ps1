# Get AD Group members for all groups and export to CSV

$groups = Get-ADGroup -Filter *

foreach ($group in $groups) {

$members = Get-ADGroupMember -Identity $group

    foreach ($member in $members) {

            [PSCustomObject]@{
            GroupName = $group.Name
            GroupDN = $group.DistinguishedName
            GroupSID = $group.SID
            MemberName = $member.name
            MemberDN = $member.DistinguishedName
            MemberSID = $member.SID
            MemberObjectClass = $member.ObjectClass
            } | Export-Csv -Path C:\temp\all_adgroupmembers_20220323_1.csv -NoClobber -NoTypeInformation -Append 
        }

}