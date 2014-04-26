<#
  This script will modify the employeeID value by panning it out
  with zeros to match the length of the HR system. It will also
  report on users that have an employeeID of invalid length.

  Syntax examples:
    To process all users:
      UpdatingEmployeeID.ps1

    To process all users from a particular OU structure:
      UpdatingEmployeeID.ps1 -SearchBase "OU=Users,OU=Corp,DC=mydemosthatrock,DC=com"

    You must use quotes around the SearchBase parameter otherwise the
    comma will be replaced with a space. This is because the comma is a
    special symbol in PowerShell.

  Release 1.0
  Written by Jeremy@jhouseconsulting.com 3rd April 2014

#>

#-------------------------------------------------------------
param([String]$SearchBase)

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}

$UsedefaultNamingContext = $False
if ([String]::IsNullOrEmpty($SearchBase)) {
  $UsedefaultNamingContext = $True
}

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
$defaultNamingContext = (get-adrootdse).defaultnamingcontext
$DistinguishedName = (Get-ADDomain).DistinguishedName
$DomainName = (Get-ADDomain).NetBIOSName
$DNSRoot = (Get-ADDomain).DNSRoot

If ($UsedefaultNamingContext -eq $True) {
  $SearchBase = $defaultNamingContext
} else {
  $TestSearchBase = Get-ADobject "$SearchBase"
  If ($TestSearchBase -eq $Null) {
    $SearchBase = $defaultNamingContext
  }
}

#-------------------------------------------------------------

# Set this to the length of the current employeeID
$CurrentemployeeIDLength = 6

# Set this to the length of the new valid employeeID string
$ValidemployeeIDLength = 8

$ReportOnly = $True
#-------------------------------------------------------------

$TotalProcessed = 0
$filter = "(employeeID=*)"
write-host -ForegroundColor Green "Finding all users with an employeeID attribute...`n"
Get-ADUser -LDAPFilter $filter -SearchBase $SearchBase -Properties * | foreach-object {$Count=0} {$Count++;[array]$Results+=$_} {$Results | select-object Name,SamAccountName,EmployeeID | % { 

    Write-Progress -Activity 'Processing Users' -Status ("Username: {0}" -f $($_.SamAccountName)) -PercentComplete (($TotalProcessed/$Count)*100)

    $line  = "Processing $($_.SamAccountName)..."
    $ID = $_.EmployeeID
    $StripSpaces = $False
    $PadWithZeros = $False
    $Update = $False
    $NewID = ""

    If ($ID.length -ne $ID.Trim().length) {
      $StripSpaces = $True
    }
    If ($ID.Trim().Length -ne $ValidemployeeIDLength) {
        $PadWithZeros = $True
    }

    If ($PadWithZeros -OR $StripSpaces) {
      If ($PadWithZeros -AND $ID.Trim().Length -eq $CurrentemployeeIDLength) {
        $NewID = $ID.Trim().PadLeft($ValidemployeeIDLength,"0")
        $Update = $True
      } elseif ($StripSpaces) {
        $NewID = $ID.Trim()
        $Update = $True
      }
    }

    If ($Update -AND $NewID -ne "") {
      If ($ReportOnly -eq $False) {
        write-host -ForegroundColor Green "$line `n - Setting employeeID attribute to $NewID"
        Set-ADUser $_.SamAccountName -employeeID $NewID
      }
    } else {
        If ($ID.Length -ne $ValidemployeeIDLength) {
        write-host -ForegroundColor Red "$line `n - Contains an invalid employeeID of $ID"
      } else {
        #write-host -ForegroundColor Green "$line `n - No changes required"
      }
    }

    $TotalProcessed ++
  }
}


