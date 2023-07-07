#Example of how to bulk create AD users from a CSV import. Assumes headers first_name and last_name.

Import-Module activedirectory

$users = Import-Csv -Path C:\Manage\MOCK_DATA_USERS.csv

$users | ForEach-Object -Process {
            $SamAcctName = ($_.first_name + "." + $_.last_name).ToLower()
            $DisplayName = $_.first_name + " " + $_.last_name
            $Password = ConvertTo-SecureString -String "P@ssw0rd!" -AsPlainText -Force
            
            New-ADUser -GivenName $_.first_name `
                       -Surname $_.last_name `
                       -Name $DisplayName `
                       -DisplayName $DisplayName `
                       -SamAccountName $SamAcctName `
                       -UserPrincipalName ($SamAcctName + "@lab.local") `
                       -AccountPassword $Password `
                       -Enabled $true `
                       -PasswordNeverExpires $true
            }