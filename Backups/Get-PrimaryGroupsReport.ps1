<#
  This script will enumerate all user accounts in a Domain, and report
  on the following attributes:

  Script Name: Get-UserReport.ps1
  Release 1.0
  Written by Jeremy@jhouseconsulting.com 27/12/2013

#>

#-------------------------------------------------------------

# Set this value to true if you want to see the progress bar.
$ProgressBar = $True

# Set this to true to process disabled user accounts.
$ProcessDisabledUsers = $True

#-------------------------------------------------------------

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$ReferenceFile = $(&$ScriptPath) + "\UserReport.csv"

$array = @()
$primarygroups = @{}
$TotalUsersProcessed = 0
$UserCount = 0

$ADRoot = ([System.DirectoryServices.DirectoryEntry]"LDAP://RootDSE")
$DefaultNamingContext = $ADRoot.defaultNamingContext

# Derive FQDN Domain Name
$TempDefaultNamingContext = $DefaultNamingContext.ToString().ToUpper()
$DomainName = $TempDefaultNamingContext.Replace(",DC=",".")
$DomainName = $DomainName.Replace("DC=","")

If ($ProcessDisabledUsers -eq $False) {
  # Create an LDAP search for all enabled users not marked as criticalsystemobjects to avoid system accounts
  $ADFilter = "(&(objectClass=user)(objectcategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2)(!(isCriticalSystemObject=TRUE))(!name=IUSR*)(!name=IWAM*)(!name=ASPNET))"
} else {
  # Create an LDAP search for all users not marked as criticalsystemobjects to avoid system accounts
  $ADFilter = "(&(objectClass=user)(objectcategory=person)(!(isCriticalSystemObject=TRUE))(!name=IUSR*)(!name=IWAM*)(!name=ASPNET))"
}
# There is a known bug in PowerShell requiring the DirectorySearcher
# properties to be in lower case for reliability.
$ADPropertyList = @("distinguishedname","samaccountname","objectsid","primarygroupid")
$ADScope = "SUBTREE"
$ADPageSize = 1000
$ADSearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($DefaultNamingContext)") 
$ADSearcher = New-Object System.DirectoryServices.DirectorySearcher
$ADSearcher.SearchRoot = $ADSearchRoot
$ADSearcher.PageSize = $ADPageSize 
$ADSearcher.Filter = $ADFilter 
$ADSearcher.SearchScope = $ADScope
if ($ADPropertyList) {
  foreach ($ADProperty in $ADPropertyList) {
    [Void]$ADSearcher.PropertiesToLoad.Add($ADProperty)
  }
}
$colResults = $ADSearcher.Findall()
$UserCount = $colResults.Count

if ($UserCount -ne 0) {
  foreach($objResult in $colResults) {
    $lastLogonTimeStamp = ""
    $lastLogon = ""
    $UserDN = $objResult.Properties.distinguishedname[0]
    $samAccountName = $objResult.Properties.samaccountname[0]

    # Get user SID
    $arruserSID = New-Object System.Security.Principal.SecurityIdentifier($objResult.Properties.objectsid[0], 0)
    $userSID = $arruserSID.Value

    # Get the SID of the Domain the account is in
    $AccountDomainSid = $arruserSID.AccountDomainSid.Value

    # Get Primary Group by binding to the user account
    $objUser = [ADSI]("LDAP://" + $UserDN)
    $primarygroupID = $objUser.primarygroupid
    If ($primarygroupID -ne $NULL) {
      # Primary group can be calculated by merging the account domain SID and primary group ID
      $primarygroupSID = $AccountDomainSid + "-" + $primarygroupID.ToString()
      $primarygroup = [adsi]("LDAP://<SID=$primarygroupSID>")
      $primarygroupname = $primarygroup.name[0]
      $objUser = $null
    } else {
      $primarygroupname = "NULL"
    }

    # Create a hashtable to capture a count of each Primary Group
    If (!($primarygroups.ContainsKey($primarygroupname))) {
      $primarygroups = $primarygroups + @{$primarygroupname = 1}
    } else {
      $value = $primarygroups.Get_Item($primarygroupname)
      $value ++
      $primarygroups.Set_Item($primarygroupname,$value)
    } # end if

    $TotalUsersProcessed ++
    If ($ProgressBar) {
      Write-Progress -Activity 'Processing Users' -Status ("Username: {0}" -f $samAccountName) -PercentComplete (($TotalUsersProcessed/$UserCount)*100)
    }
  }
  write-host -ForegroundColor Green "`nA breakdown of the $($primarygroups.count) Primary Groups applied to $TotalUsersProcessed user objects:"
  $primarygroups.GetEnumerator() | Sort-Object Name | % { 
      write-host -ForegroundColor Green "- $($_.key): $($_.value)"
    }
}
