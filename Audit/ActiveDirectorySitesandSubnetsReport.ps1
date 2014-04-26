<#
  Active Directory Sites and Subnets Report

  Release 1.0 Written by Jeremy@jhouseconsulting.com 13th September 2013
#>

#-------------------------------------------------------------
# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$SitesReport = $(&$ScriptPath) + "\ActiveDirectorySitesReport.csv"
$SubnetsReport = $(&$ScriptPath) + "\ActiveDirectorySubnetsReport.csv"

#-------------------------Site Report-------------------------
# This module was Written by Brian Seltzer
# List Sites and Subnets in Active Directory using PowerShell
# http://www.itadmintools.com/2011/08/list-sites-and-subnets-in-active.html
$siteDescription=@{}
$siteSubnets=@{}
$sitesDN="LDAP://CN=Sites," + $([adsi] "LDAP://RootDSE").Get("ConfigurationNamingContext")
$subnetsDN="LDAP://CN=Subnets,CN=Sites," + $([adsi] "LDAP://RootDSE").Get("ConfigurationNamingContext")
#get the site names and descriptions
foreach ($site in $([adsi] $sitesDN).psbase.children){
 if($site.objectClass -eq "site"){
  $siteName=([string]$site.cn).toUpper()
  $siteDescription[$siteName]=$site.Description
  $siteSubnets[$siteName]=@()
 }
}
#get the subnets and associate them with the sites
foreach ($subnet in $([adsi] $subnetsDN).psbase.children){
 $site=[adsi] "LDAP://$($subnet.siteObject)"
 if($site.cn -ne $null){
  $siteName=([string]$site.cn).toUpper()
  $siteSubnets[$siteName] += $subnet.cn
 }else{
  $siteDescription["Orphaned"]="Subnets not associated with any site"
  if($siteSubnets["Orphaned"] -eq $null){ $siteSubnets["Orphaned"] = @() }
  $siteSubnets["Orphaned"] += $subnet.cn
 }
}
#write output to screen
foreach ($siteName in $siteDescription.keys | sort){
 "$siteName  $($siteDescription[$siteName])"
 foreach ($subnet in $siteSubnets[$siteName]){
  "`t$subnet"
 }
}

#-------------------------Site Report-------------------------
$allsites = @()
$sitesDN="LDAP://CN=Sites," + $([adsi] "LDAP://RootDSE").Get("ConfigurationNamingContext")
#get the site names and descriptions
foreach ($site in $([adsi] $sitesDN).psbase.children){
 if($site.objectClass -eq "site"){
   $data = "" | select Name, Description, Location
   $data.Name = $($site.Name)
   $data.Description = $($site.Description)
   $data.Location = $($site.Location)
   $allsites += $data
 }
}
Write-Host -ForegroundColor Green $allsites.count "sites have been exported to $SitesReport"
$allsites | Sort-Object Name | Export-Csv -notype "$SitesReport"
# Remove the quotes
(get-content "$SitesReport") |% {$_ -replace '"',""} | out-file "$SitesReport" -Fo -En ascii

#-----------------------Subnet Report-------------------------
$allsubnets = @()
$subnetsDN="LDAP://CN=Subnets,CN=Sites," + $([adsi] "LDAP://RootDSE").Get("ConfigurationNamingContext")
foreach ($subnet in $([adsi] $subnetsDN).psbase.children){
  $net = [ADSI]"$($subnet.Path)"
  $data = "" | select Site, Name, Description, Location
  If ($($net.cn).Contains("CNF:") -eq $False) {
    $data.name = $($net.cn)
  } else {
    $data.name = [string]::join("\0A",($($net.cn).Split("`n")))
  }
  $data.Location = $($net.location)
  $data.Description = $($net.description)
  If ($net.siteobject -ne $NULL) {
    $st = $($net.siteobject).split(",")
    $data.site = $st[0].Replace("CN=","")
  } Else {
    $st = "*Orphaned"
    $data.site = $st
  }
  $allsubnets += $data
}
Write-Host -ForegroundColor Green $allsubnets.count "subnets have been exported to $SubnetsReport"
$allsubnets | Sort-Object Site | Export-Csv -notype "$SubnetsReport"
# Remove the quotes
(get-content "$SubnetsReport") |% {$_ -replace '"',""} | out-file "$SubnetsReport" -Fo -En ascii
