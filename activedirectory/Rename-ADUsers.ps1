Import-Module -Name activedirectory

$old_new = Import-Csv -Path C:\manage\old_name_new_name.csv

$old_new |  ForEach-Object {

$old_upn =          $_.old_upn
$new_upn =          $_.new_upn
$old_sam_acc =      $_.old_sam_acc
$new_sam_acc =      $_.new_sam_acc
$new_display_name = $_.new_display_name
$firstname =        $_.firstname
$surname =          $_.surname

$ObjectGUID = ((Get-ADUser -Identity $old_sam_acc).objectguid | Select-Object -ExpandProperty guid)

    $UserNameToChange = (Get-ADUser `
                        -Identity $ObjectGUID `
                        -Properties *
                        )

    $UserNameToChange | Set-ADUser `
                        -DisplayName $new_display_name `
                        -SamAccountName $new_sam_acc `
                        -GivenName $firstname `
                        -Surname $surname `
                        -UserPrincipalName $new_upn `
                        -EmailAddress $new_upn `
                        -Add @{proxyAddresses = ("SMTP:" + $new_upn)}

    $UserNameToChange | Rename-ADObject `
                        -NewName $new_display_name

    Set-ADAccountPassword -Identity $ObjectGUID -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "PasswordHere" -Force)
}