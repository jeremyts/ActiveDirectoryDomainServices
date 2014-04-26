<#

  Active Directory Site Link Reporting
  
  Original script written by Ashley McGlone, Microsoft PFE, June 2011
  Report and Edit AD Site Links From PowerShell:
    http://blogs.technet.com/b/ashleymcglone/archive/2011/06/29/report-and-edit-ad-site-links-from-powershell-turbo-your-ad-replication.aspx

#>

#-------------------------------------------------------------
# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$SiteLinksReport = $(&$ScriptPath) + "\ActiveDirectorySiteLinksReport.csv"

#----------------------Site Link Report-----------------------

Import-Module ActiveDirectory

# Report of all site links and related settings
$AllSiteLinks = Get-ADObject -Filter 'objectClass -eq "siteLink"' -Searchbase (Get-ADRootDSE).ConfigurationNamingContext -Property Description, Options, Cost, ReplInterval, SiteList, Schedule | Select-Object Name, Description, @{Name="SiteCount";Expression={$_.SiteList.Count}}, Cost, ReplInterval, @{Name="Schedule";Expression={If($_.Schedule){If(($_.Schedule -Join " ").Contains("240")){"NonDefault"}Else{"24x7"}}Else{"24x7"}}}, Options

#$AllSiteLinks | Format-Table * -AutoSize
$AllSiteLinks | Export-Csv -notype "$SiteLinksReport"

#Remove the quotes
(get-content "$SiteLinksReport") |% {$_ -replace '"',""} | out-file "$SiteLinksReport" -Fo -En ascii
