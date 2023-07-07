function New-MyCompanyUser
{
<#
.Synopsis
   Create a New My Company user
.DESCRIPTION
   Creates a new user with optional email aliases and optionally adds to RDS users group, then syncs them to Azure AD
.EXAMPLE
   New-function New-MyCompanyUser -FirstName john -SurName SMITH -TempPassword PassWord1234 -Verbose
.EXAMPLE
   New-function New-MyCompanyUser -FirstName john -SurName SMITH -Aliases j.smithy@mycompany.org.au -TempPassword PassWord1234 -RDSAccess -Verbose
#>
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # Users first name
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [string]$FirstName,

        # Users surname 
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [string]$SurName,

        # Optional email aliases
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [string[]]$Aliases,

        # A temporary password you are assigning to user
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [string]$TempPassword,

        # Optionallly add user to RDS group
        [Parameter(Mandatory=$false,
                   ValueFromPipelineByPropertyName=$true,
                   Position=4)]
        [switch]$RDSAccess
    )

    Begin
    {
    #Server and Destination OU
    $DomainController = "MyCompany-DC01.ad.mycompany.org.au"
    $DestOU = "OU=Staff,OU=Users,OU=ad.mycompany.org.au,DC=ad,DC=mycompany,DC=org,DC=au"
    }
    Process
    {
    #Set variables for use in New-ADUser
    $SamAccountName = (("$FirstName" + "." + "$SurName").ToLower())
    $FirstName = (($FirstName).Substring(0,1).toupper()+$FirstName.Substring(1).tolower())
    $SurName = (($SurName).Substring(0,1).toupper()+$SurName.Substring(1).tolower())
    $DisplayName = ("$FirstName" + " " + "$SurName")
    $UPN = ("$SamAccountName" + "@mycompany.org.au")

    #Set variable for RDS Group
    $RDSGroup = "CN=RDS_Users,OU=Groups,OU=ad.mycompany.org.au,DC=ad,DC=mycompany,DC=org,DC=au"

    #Create user account, set password
       try
       {
          Write-Verbose -Message "Creating user account..."
           New-ADUser -Name $DisplayName `
                      -SamAccountName $SamAccountName `
                      -UserPrincipalName $UPN `
                      -DisplayName $DisplayName `
                      -GivenName $FirstName `
                      -Surname $SurName `
                      -EmailAddress $UPN `
                      -Enabled $true `
                      -Server $DomainController `
                      -Path $DestOU `
                      -AccountPassword (ConvertTo-SecureString "$TempPassword" -AsPlainText -Force) `
                      -ErrorAction Stop

       }
       catch
       {
           Write-Output "Failed to create AD User.."
           Write-Output -InputObject $Error[0]
       }

       #Set new user proxy addresses (Aliases)
       if (-not ([string]::IsNullOrEmpty($Aliases)))
       {
        try
        {
         Write-Verbose -Message "Setting proxy addresses..."
            #Set additional proxy addresses
            Foreach ($address in $Aliases)
                        {
                        Set-ADUser -Identity $SamAccountName `
                                   -Add @{proxyAddresses = ("smtp:" + "$address")} `
                                   -ErrorAction Stop
                        }
            #Set primary proxy address, done last to move it to the top
            Set-ADUser -Identity $SamAccountName `
                       -Add @{proxyAddresses = ("SMTP:" + "$UPN")} `
                       -ErrorAction Stop
                        
        }
        catch
        {
           Write-Output "Failed to add aliases (proxy addresses).."
           Write-Output -InputObject $Error[0]
        }
       }

       #Add user to RDS group if required
       if ($RDSAccess.IsPresent)
       {
        try
        {
         Write-Verbose -Message "Adding to RDS Group..."
            if ($RDSAccess)
                        {
                        Add-ADGroupMember -Identity $RDSGroup `
                                          -Members $SamAccountName `
                                          -ErrorAction Stop
                        }
        }
        catch
        {
           Write-Output "Failed to add user to RDS group.."
           Write-Output -InputObject $Error[0]
        }
       }
       #Pause script for a minute to allow for AD changes before stating AD Sync
       Write-Verbose -Message "Sleeping script for 20 seconds to allow for object to be created..."
       Start-Sleep -Seconds 20

       #Sync AzureAD
       Write-Verbose -Message "Starting Azure AD Sync Cycle to replicate object to AAD..."
       Start-ADSyncSyncCycle -PolicyType Initial -Verbose
    }
    End
    {
      Write-Verbose -Message "Finished!"
    }
}
Export-ModuleMember -Function New-MyCompanyUser