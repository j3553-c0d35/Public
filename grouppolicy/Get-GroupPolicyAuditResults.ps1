# global variables
[string]$domain = '<domain>'
[string]$server = '<server>'
[string]$csvPath = '<csvpath>'
[float]$timeZoneUtcDifference = '<utcdifferene>'


# helper function, convert Canonical to DN function ConvertFrom-CanonicalOU {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [ValidateNotNullOrEmpty()]
        [string]$CanonicalName
    )
    process {
        $obj = $CanonicalName.Split('/')
        [string]$DN = 'OU=' + $obj[$obj.count - 1]
        for ($i = $obj.count - 2; $i -ge 1; $i--) { $DN += ',OU=' + $obj[$i] }
        $obj[0].split('.') | ForEach-Object { $DN += ',DC=' + $_ }
        return $DN
    }
}

# helper function, get AD sites
function Get-AllAdSites {
$getADO = @{
     LDAPFilter = "(Objectclass=site)"
     properties = "Name"
     SearchBase = (Get-ADRootDSE -Server $server).ConfigurationNamingContext
 }
 $allSites = Get-ADObject @getADO -Server $server  return $allSites }


# helper function, get all AD site GPO links
function Get-AllAdGpoSiteLinks {    

# Get all AD Sites
$sites = Get-AllAdSites

# Setup com opbject to get site gpo objects $gpm = New-Object -ComObject "GPMGMT.GPM"
$gpmConstants = $gpm.GetConstants()
$gpmdomain = $gpm.GetDomain("$domain", "", $gpmConstants.UseAnyDC) $SiteContainer = $gpm.GetSitesContainer("$domain", "$domain", $null, $gpmConstants.UseAnyDC)

$s2 = @()
            
foreach ($site in $sites) {

        $comSite = $SiteContainer.GetSite($($site.name))

        $s1 = $comSite.GetGPOLinks()

        if ($s1) {

        $s2 += ($s1 | Select-Object -Property @{name="SiteName";expression={"$($site.name)"}},gpoid,enabled,enforced,somlinkorder)
    
        } #end if $s1
    
    } #end foreach site
return $s2
} # end Get-AllAdGpoSiteLinks



$GPOs = Get-GPO -All -Domain $domain -Server $server | Sort-Object -Property DisplayName

$AllAdGpoSiteLinks = Get-AllAdGpoSiteLinks

foreach ($GPO in $GPOs) {

        Write-host "Working with $($GPO.DisplayName)" -ForegroundColor Yellow

        [xml]$Report = $GPO | Get-GPOReport -ReportType XML -Domain $domain -Server $server

        if ($report.GPO.Identifier.Identifier.'#text' -in $AllAdGpoSiteLinks.GPOID) {$siteLink = $true}else{$siteLink = $false}
                
        # get gpo security filtering
        $GPOPermission = (Get-GPPermission -Name $GPO.DisplayName -DomainName $domain -all -Server $server | Where-Object {$_.permission -eq 'GPOApply'}).Trustee.Name
        $perms = ($GPOPermission -join ",")

                # get gpo enforced and order
                if ($Report.GPO.LinksTo) {

                    foreach ($link in $Report.GPO.LinksTo) {

                        $ouDN = Get-ADOrganizationalUnit -Identity (ConvertFrom-CanonicalOU -CanonicalName $Link.SOMPath) -Server $server
                        $gpEnforcedOrder = (Get-GPInheritance -Domain $domain -Server $server -Target $ouDN.DistinguishedName).GpoLinks | Select-Object -Property enforced,order,gpoid

                        foreach ($g in $gpEnforcedOrder) {
                            
                            # set gpoid for comparison
                            $reportGPOID = $Report.GPO.Identifier.Identifier.'#text'
                            $reportGPOID = $reportGPOID.TrimStart('{')
                            $reportGPOID = $reportGPOID.TrimEnd('}')

                            if ($g.gpoid -eq $reportGPOID) {

                                if ($Report.GPO.FilterName) {

                                [PSCustomObject]@{            
                                GPOName = $Report.GPO.Name 
                                GPOLinkPath = $Link.SOMPath
                                GPOLinkEnabled = $Link.Enabled
                                GPOEnforced = $g.Enforced
                                GPOOrder = $g.Order
                                SiteOrOuLink = 'OU'
                                ComputerEnabled = $Report.GPO.Computer.Enabled
                                UserEnabled = $Report.GPO.User.Enabled
                                WMIFilterName = $Report.GPO.FilterName
                                WMIFilterDescription = $Report.GPO.FilterDescription
                                SecurityFiltering = $perms
                                CreatedTime = $($createdDate = [datetime]$report.GPO.CreatedTime ; $createdDate.AddHours($timeZoneUtcDifference))
                                ModifiedTime = $($modifiedDate = [datetime]$report.GPO.ModifiedTime ; $modifiedDate.AddHours($timeZoneUtcDifference))
                                } | Export-Csv -Path $csvPath -NoTypeInformation -NoClobber -Append
                    
                    } # end if report GPO filter name
                            else {

                                [PSCustomObject]@{            
                                GPOName = $Report.GPO.Name 
                                GPOLinkPath = $Link.SOMPath
                                GPOLinkEnabled = $Link.Enabled
                                GPOEnforced = $g.Enforced
                                GPOOrder = $g.Order
                                SiteOrOuLink = 'OU'
                                ComputerEnabled = $Report.GPO.Computer.Enabled
                                UserEnabled = $Report.GPO.User.Enabled
                                WMIFilterName = "Not Applicable"
                                WMIFilterDescription = "Not Applicable"
                                SecurityFiltering = $perms
                                CreatedTime = $($createdDate = [datetime]$report.GPO.CreatedTime ; $createdDate.AddHours($timeZoneUtcDifference))
                                ModifiedTime = $($modifiedDate = [datetime]$report.GPO.ModifiedTime ; $modifiedDate.AddHours($timeZoneUtcDifference))
                                } | Export-Csv -Path $csvPath -NoTypeInformation -NoClobber -Append
                        

                    } # end else report GPO filter name

                } # if gpoid match

            } # end foreach $g

        } # end foreach link

    } # end if report gpo links to
                            elseif ($Report.GPO.FilterName -and -not $Report.GPO.LinksTo -and -not $siteLink) {

                                [PSCustomObject]@{            
                                GPOName = $Report.GPO.Name 
                                GPOLinkPath = "Not Applicable"
                                GPOLinkEnabled = "Not Applicable"
                                GPOEnforced = "Not Applicable"
                                GPOOrder = "Not Applicable"
                                SiteOrOuLink = 'Not Applicable'
                                ComputerEnabled = $Report.GPO.Computer.Enabled
                                UserEnabled = $Report.GPO.User.Enabled
                                WMIFilterName = $Report.GPO.FilterName
                                WMIFilterDescription = $Report.GPO.FilterDescription
                                SecurityFiltering = $perms
                                CreatedTime = $($createdDate = [datetime]$report.GPO.CreatedTime ; $createdDate.AddHours($timeZoneUtcDifference))
                                ModifiedTime = $($modifiedDate = [datetime]$report.GPO.ModifiedTime ; $modifiedDate.AddHours($timeZoneUtcDifference))
                                } | Export-Csv -Path $csvPath -NoTypeInformation -NoClobber -Append

    } # end elseif report gpo filter name and not report gpo links to
                            elseif (-not $Report.GPO.FilterName -and -not $Report.GPO.LinksTo -and -not $siteLink) {

                                [PSCustomObject]@{            
                                GPOName = $Report.GPO.Name 
                                GPOLinkPath = "Not Applicable"
                                GPOLinkEnabled = "Not Applicable"
                                GPOEnforced = "Not Applicable"
                                GPOOrder = "Not Applicable"
                                SiteOrOuLink = 'Not Applicable'
                                ComputerEnabled = $Report.GPO.Computer.Enabled
                                UserEnabled = $Report.GPO.User.Enabled
                                WMIFilterName = "Not Applicable"
                                WMIFilterDescription = "Not Applicable"
                                SecurityFiltering = $perms
                                CreatedTime = $($createdDate = [datetime]$report.GPO.CreatedTime ; $createdDate.AddHours($timeZoneUtcDifference))
                                ModifiedTime = $($modifiedDate = [datetime]$report.GPO.ModifiedTime ; $modifiedDate.AddHours($timeZoneUtcDifference))
                                } | Export-Csv -Path $csvPath -NoTypeInformation -NoClobber -Append
                                

                            } # end elseif report gpo filter name and not report gpo linksto

                            elseif ($siteLink -and -not $Report.GPO.FilterName) {
                                
                                foreach ($site in $AllAdGpoSiteLinks) {

                                    if ( $site.GPOID -eq $report.GPO.Identifier.Identifier.'#text') {

                                        [PSCustomObject]@{            
                                        GPOName = $Report.GPO.Name 
                                        GPOLinkPath = $site.SiteName
                                        GPOLinkEnabled = $site.Enabled
                                        GPOEnforced = $site.Enforced
                                        GPOOrder = $site.SOMLinkOrder
                                        SiteOrOuLink = 'SITE'
                                        ComputerEnabled = $Report.GPO.Computer.Enabled
                                        UserEnabled = $Report.GPO.User.Enabled
                                        WMIFilterName = "Not Applicable"
                                        WMIFilterDescription = "Not Applicable"
                                        SecurityFiltering = $perms
                                        CreatedTime = $($createdDate = [datetime]$report.GPO.CreatedTime ; $createdDate.AddHours($timeZoneUtcDifference))
                                        ModifiedTime = $($modifiedDate = [datetime]$report.GPO.ModifiedTime ; $modifiedDate.AddHours($timeZoneUtcDifference))
                                        } | Export-Csv -Path $csvPath -NoTypeInformation -NoClobber -Append
                            
                            } # end if site.gpoid eq report.gpo.identifier

                        } # end foreach site in alladgpositelinks
                                    
                    } # end elseif sitelink and not report gpo filter name

                            elseif ($siteLink -and $Report.GPO.FilterName) {

                                foreach ($site in $AllAdGpoSiteLinks) {

                                    if ( $site.GPOID -eq $report.GPO.Identifier.Identifier.'#text') {

                                        [PSCustomObject]@{            
                                        GPOName = $Report.GPO.Name 
                                        GPOLinkPath = $site.SiteName
                                        GPOLinkEnabled = $site.Enabled
                                        GPOEnforced = $site.Enforced
                                        GPOOrder = $site.SOMLinkOrder
                                        SiteOrOuLink = 'SITE'
                                        ComputerEnabled = $Report.GPO.Computer.Enabled
                                        UserEnabled = $Report.GPO.User.Enabled
                                        WMIFilterName = $Report.GPO.FilterName
                                        WMIFilterDescription = $Report.GPO.FilterDescription
                                        SecurityFiltering = $perms
                                        CreatedTime = $($createdDate = [datetime]$report.GPO.CreatedTime ; $createdDate.AddHours($timeZoneUtcDifference))
                                        ModifiedTime = $($modifiedDate = [datetime]$report.GPO.ModifiedTime ; $modifiedDate.AddHours($timeZoneUtcDifference))
                                        } | Export-Csv -Path $csvPath -NoTypeInformation -NoClobber -Append

                            } # end if site.gpoid eq report.gpo.identifier

                        } # end foreach site in alladgpositelinks
                                
                    } # end elseif sitelink and report gpo filter name


} # end foreach gpo