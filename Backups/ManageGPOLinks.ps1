<#
  This script will Export and Import GPO Links

  Syntax examples:
    Export:
      ManageGPOLinks.ps1 -Action Export -ReferenceFile GPOLinksExport.csv
    Import:
      ManageGPOLinks.ps1 -Action Import -ReferenceFile GPOLinksExport.csv

  It is based on the following script by Microsoft's Manny Murguia:
    - Migrating GPO Links between Domains with PowerShell:
      http://blogs.technet.com/b/manny/archive/2013/05/18/migrating-gpo-links-between-domains-with-powershell.aspx

  Release 1.0
  Written by Jeremy@jhouseconsulting.com 13th September 2013
#>

#-------------------------------------------------------------
param([String]$Action,[String]$ReferenceFile,[String]$LogFile)

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}

if ([String]::IsNullOrEmpty($Action)) {
  write-host -ForeGroundColor Red "Action is a required parameter. Exiting Script.`n"
  exit
} else {
  switch ($Action)
  {
    "Import" {$Import = $true;$Export = $false}
    "Export" {$Import = $false;$Export = $true}
    default {$Import = $false;$Export = $false}
  }
  if ($Import -eq $false -AND $Export -eq $false) {
    write-host -ForeGroundColor Red "The Action parameter is invalid. Exiting Script.`n"
    exit
  }
}

if ([String]::IsNullOrEmpty($ReferenceFile)) {
  write-host -ForeGroundColor Red "ReferenceFile is a required parameter. Exiting Script.`n"
  exit
}

if ([String]::IsNullOrEmpty($LogFile))
{
    $LogFile = $(&$ScriptPath) + "\LinkWMIFilters.txt";
}
set-content $LogFile $NULL;

#-------------------------------------------------------------

# Import the Active Directory Module
Import-Module ActiveDirectory -WarningAction SilentlyContinue
if($Error.Count -eq 0) {
   #Write-Host "Successfully loaded Active Directory Powershell's module" -ForeGroundColor Green
}else{
   Write-Host "Error while loading Active Directory Powershell's module : $Error" -ForeGroundColor Red
   exit
}

#-------------------------------------------------------------
If ($Import -eq $true) {

    $FileExists = Test-Path $ReferenceFile
    if ($FileExists -eq $false)
    {
    write-host -ForegroundColor Red "Input file does not exist"
    Write-Host -ForegroundColor Red "Exiting Script"
    return
    }

    $thisDomain = Get-ADDomain
    $thisDomainDN = $thisDomain.DistinguishedName
    $header = "objectDN","domainDN","Links"
    $Links = $null
    $Links = import-csv $ReferenceFile -Delimiter "`t" -Header $header
    if ($Links -eq $null)
    {
    write-host -ForegroundColor Red "No input was detected"
    Write-Host -ForegroundColor Red "Exiting Script"
    return
    }

    foreach ($Link in $Links)
    {
    $currentObject = $NULL
    $Link.objectDN = $Link.objectDN -replace $Link.domainDN,$thisDomainDN
    [Array]$objectLinks = $Link.Links.Split("`v")
    $NewLink = $null
    $currentDN = $Link.objectDN
    add-content $LogFile "$currentDN was configured with the following links:"
    foreach ($objectLink in $objectLinks)
        {
        $currentLink = $NULL
        $linkName = $objectLink.TrimEnd("0","1","2")
        $linkName = $linkName.TrimEnd(";")
            if ($linkName)
            {
            add-content $LogFile $linkName
            $currentLink = Get-ADObject -Filter {objectClass -eq "groupPolicyContainer" -and displayName -eq $linkName}
                if ($currentLink -ne $NULL)
                {
                $NewLink = $NewLink + "[LDAP://" + $currentLink + ";" + $objectLink.Substring($objectLink.Length -1,1) + "]"
                }
                else
                {
                Add-Content $LogFile "Error: $linkname does not appear to exist in the destination domain. Please re-import it or create a new GPO with the same name."
                }
            }
        }
        try
        {
        $currentObject = Get-ADObject $Link.objectDN -Properties gpLink
        $currentObject.gpLink = $NewLink
            if ($NewLink)
            {
            Set-ADObject $Link.objectDN -Replace @{gpLink = $NewLink}
            }
            else
            {
            $currentObjectDN = $Link.objectDN
            Add-Content $LogFile "Error: It appears none of the GPO's previously linked to '$currentObjectDN' exist. Please re-import the GPO's to the destination domain."
            }
        }
        catch
        {
        $currentObjectDN = $Link.objectDN
        add-content $LogFile "Error: $currentObjectDN does not exist. Create the object and try again."
        }
            if ($NewLink)
            {
            add-content $LogFile "gPLink will be set to: $NewLink"
            }
            else
            {
            Add-Content $LogFile "gPLink will not be modified on this object."
            }
        add-content $LogFile "---END---"
    }
    write-host -ForegroundColor Yellow "A log file has been saved at $LogFile"
}


#-------------------------------------------------------------
If ($Export -eq $true) {

    set-content $ReferenceFile $null
    $Links = $NULL
    $thisDomain = Get-ADDomain
    $thisDomainDN = $thisDomain.DistinguishedName
    $thisDomainConfigurationPartition = "CN=Configuration," + $thisDomainDN

    $Links = Get-ADObject -Filter {gpLink -LIKE "[*]"} -Properties gpLink
    $Links += Get-ADObject -Filter {gpLink -LIKE "[*]"} -Searchbase $thisDomainConfigurationPartition -Properties gpLink

    $NewLine = $null
    if ($Links)
    {
        foreach ($Link in $Links)
        {
          if ($Link -ne $NULL) {
          $NewLine = $null
          $LinkList = $Link.gpLink.Split('\[|\]')

	        foreach ($LinkItem in $LinkList)
	        {
		        if ($LinkItem)
		        {
		        $LinkSplit = $LinkItem.Split(";")
		        $LinkItem = $LinkItem.TrimStart("LDAP://")
		        $LinkItem = $LinkItem.TrimEnd(';0|;1|;2')
		        $LinkItem = get-adobject $LinkItem -Properties displayName
		        $NewLine = $NewLine + $LinkItem.DisplayName + ";" + $LinkSplit[1] + "`v"
		        }
	        }
	
	        $NewLine = $Link.DistinguishedName + "`t" + $thisDomainDN + "`t" + $NewLine
	        add-content $ReferenceFile $NewLine
          }
        }
            write-host -ForegroundColor Yellow "The output file has been saved at $ReferenceFile"
    }
    else
    {
        write-host -ForegroundColor Red "No GPO Links exist in this domain"
        write-host -ForegroundColor Red "Exiting script"
        Set-Content $ReferenceFile $NULL
        return
    }
}