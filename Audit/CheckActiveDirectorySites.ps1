<#
Check Active Directory Sites
http://blogs.metcorpconsulting.com/tech/?p=366
#>

# Set variables
[array] $SitesWithNoSubnet = @()
[array] $SitesWithNoSiteLinks = @()
[array] $SitesWithNoISTG = @()
[array] $DCsnotGC = @()
[array] $SitesWithNoGC = @()
#Get AD Domain (lightweight & fast method)
$DomainDNS = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().Name 
import-module activedirectory
Write-Verbose "Get AD Site List `r"
[array] $ADSites = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().Sites
$ADSitesCount = $ADSites.Count
Write-Output "There are $ADSitesCount AD Sites `r"
## Check AD Sites
        Write-Verbose "Checking for AD Site issues… `r "
        ForEach ($Site in $ADSites)
            {  ## OPEN ForEach Site in ADSites
                $SiteName = $Site.Name
                [array] $SiteSubnets = $Site.Subnets
                [array] $SiteServers = $Site.Servers
                [array] $SiteAdjacentSites = $Site.AdjacentSites
                [array] $SiteLinks = $Site.SiteLinks
                $SiteInterSiteTopologyGenerator = $Site.InterSiteTopologyGenerator
              # Check for missing subnets                      
                IF (!$SiteSubnets)
                    {  ## OPEN IF there are no Site Subnets
                        Write-Verbose "The site $SiteName does not have a configured subnet. `r "
                        [array] $SitesWithNoSubnet += $SiteName
                    }  ## OPEN IF there are no Site Subnets
                # Check for missing site link 
                IF (!$SiteLinks)
                    {  ## OPEN IF there are no Site Links for this site
                        Write-Verbose "The site $SiteName does not have an associated site link. `r "
                        [array] $SitesWithNoSiteLinks += $SiteName
                    }  ## OPEN IF there are no Site Links for this site
                # Check for missing ISTG     
                IF (!$SiteInterSiteTopologyGenerator)
                    {  ## OPEN IF there are no ISTG  for this site
                        Write-Verbose "The site $SiteName does not have a configured InterSite Topology Generator server `r "
                        [array] $SitesWithNoISTG += $SiteName
                    }  ## OPEN IF there are no ISTG for this site     
                # Find AD Sites with no GCs
                $SiteDC = Get-ADDomainController -filter { (Site -eq $Site) -and (IsGlobalCatalog -eq $True) } 
                IF (!$SiteDC)
                    {  ## OPEN IF there are no GCs for this site
                        Write-Verbose "The site $SiteName does not have a Global Catalog associated with it `r "
                        [array] $SitesWithNoGC += $SiteName
                    }  ## OPEN IF there are no GCs for this site
            }  ## CLOSE ForEach Site in ADSites
        $SitesWithNoSubnetCount = $SitesWithNoSubnet.Count
        IF ($SitesWithNoSubnetCount -ge 1)
            {  ## IF Count is >= 1
                Write-Output "The following $SitesWithNoSubnetCount sites do not have subnets associated with them `r "
                $SitesWithNoSubnet
                Write-Output " `r "
            }  ## IF Count is >= 1
        $SitesWithNoSiteLinksCount = $SitesWithNoSiteLinks.Count
        IF ($SitesWithNoSiteLinksCount -ge 1)
            {  ## IF Count is >= 1
                Write-Output "The following $SitesWithNoSiteLinksCount sites do not have Site Links associated with them `r "
                $SitesWithNoSiteLinks
                Write-Output " `r "
            }  ## IF Count is >= 1
        $SitesWithNoISTGCount = $SitesWithNoISTG.Count
        IF ($SitesWithNoISTGCount -ge 1)
            {  ## IF Count is >= 1
                Write-Output "The following $SitesWithNoISTGCount sites do not an ISTG associated with them `r "
                $SitesWithNoISTG
                Write-Output " `r "
            }  ## IF Count is >= 1
        $SitesWithNoGCCount = $SitesWithNoGC.Count
        IF ($SitesWithNoGCCount -ge 1)
            {  ## IF Count is >= 1
                Write-Output "The following $SitesWithNoGCCount sites do not have a GC associated with them `r "
                $SitesWithNoGC
                Write-Output " `r "
            }  ## IF Count is >= 1
