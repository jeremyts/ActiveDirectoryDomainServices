<#
  This script will enumerate all user accounts that contain an
  employeeID value. It can either take action and update them
  by padding it out with zeros to match the length of the HR
  system, or report only. It will test to see which users have
  an employeeID of invalid length, which will need to be manually
  fixed in line with the HR system.

  Syntax examples:
    To process all users:
      UpdateEmployeeID.ps1 -Action Update

    To report on all users:
      UpdateEmployeeID.ps1 -Action Report

    To process all users from a particular OU structure:
      UpdateEmployeeID.ps1 -Action Update -SearchBase "OU=Users,OU=Corp,DC=mydemosthatrock,DC=com"

    You must use quotes around the SearchBase parameter otherwise the
    comma will be replaced with a space. This is because the comma is a
    special symbol in PowerShell.

  Script Name: UpdateEmployeeID.ps1
  Release: 1.1
  Written by Jeremy@jhouseconsulting.com 3rd April 2014
  Modified by Jeremy@jhouseconsulting.com 7th April 2014

#>

#-------------------------------------------------------------
param([String]$Action,[String]$SearchBase)

if ([String]::IsNullOrEmpty($Action)) {
  write-host -ForeGroundColor Red "Action is a required parameter. Exiting Script.`n"
  exit
} else {
  switch ($Action)
  {
    "Report" {$ReportOnly = $True}
    "Update" {$ReportOnly = $False}
    default {$ReportOnly = $True}
  }
}

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

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$ReferenceFile = $(&$ScriptPath) + "\UpdatingEmployeeID.csv"

#-------------------------------------------------------------

# Set this to the length of the current employeeID
$CurrentemployeeIDLength = 6

# Set this to the length of the new valid employeeID string
$ValidemployeeIDLength = 8

#-------------------------------------------------------------

if (Test-Path -path $ReferenceFile) {
  remove-item $ReferenceFile -force -confirm:$false
}
$header = "Action,SamAccountName,EmployeeID"
$header | Out-File -filepath "$ReferenceFile"

$TotalProcessed = 0
$filter = "(employeeID=*)"
write-host -ForegroundColor Green "Finding all users with an employeeID attribute..."
Get-ADUser -LDAPFilter $filter -SearchBase $SearchBase -Properties * | foreach-object {$Count=0} {$Count++;[array]$Results+=$_} {$Results | select-object Name,SamAccountName,EmployeeID | % { 

    Write-Progress -Activity 'Processing Users' -Status ("Username: {0}" -f $($_.SamAccountName)) -PercentComplete (($TotalProcessed/$Count)*100)

    $line = ""
    $StripSpaces = $False
    $PadWithZeros = $False
    $Update = $False
    $NewID = ""
    $ID = $_.EmployeeID

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
      $Action = "Change"
      $EmployeeID = $NewID
      If ($ReportOnly -eq $False) {
        Set-ADUser $_.SamAccountName -employeeID $NewID
      }
    } else {
      If ($ID.Length -ne $ValidemployeeIDLength) {
        $Action = "Warning"
      } else {
        $Action = "Informational"
      }
      $EmployeeID = $ID
    }

    $line = $Action+","+$($_.SamAccountName)+","+$EmployeeID

    $line | Out-File -filepath "$ReferenceFile" -append -noclobber

    $TotalProcessed ++
  }
}
write-host -ForegroundColor Green "Review '$ReferenceFile' for the results."
