<#
  This script will find duplicate employeeID values and export them
  to a CSV file.

  Syntax examples:
    To process all users:
      UpdatingEmployeeID.ps1

    To process all users from a particular OU structure:
      UpdatingEmployeeID.ps1 -SearchBase "OU=Users,OU=Corp,DC=mydemosthatrock,DC=com" -ReferenceFile "c:\MyOutput.csv"

    You must use quotes around the SearchBase parameter otherwise the
    comma will be replaced with a space. This is because the comma is a
    special symbol in PowerShell.

  Release 1.0
  Written by Jeremy@jhouseconsulting.com 3rd April 2014

#>

#-------------------------------------------------------------
param([String]$SearchBase,[String]$ReferenceFile)

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}

if ([String]::IsNullOrEmpty($ReferenceFile)) {
  $ReferenceFile = $(&$ScriptPath) + "\DuplicateEmployeeIDs.csv";
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

#-------------------------------------------------------------

function Get-LastLoggedOnDate ([string] $Date) {
  If ($Date -eq $NULL -OR $Date -eq "") {
    $Date = "Never logged on before"
  }
  $Date
}

$LastLoggedOnDate = @{Name='LastLoggedOnDate'; Expression={ Get-LastLoggedOnDate $_.LastLogonDate } }

$EmployeeID = @{Name='EmployeeID'; Expression={ $_.EmployeeID.Trim() } }

$TotalProcessed = 0
$filter = "(employeeID=*)"
write-host -ForegroundColor Green "Finding all users with an employeeID attribute and exporting the duplicates to '$ReferenceFile'`n"
Get-ADUser -LDAPFilter $filter -SearchBase $SearchBase -Properties * | select-object $EmployeeID,SamAccountName,Name,Enabled,$LastLoggedOnDate,whenCreated | Group-Object EmployeeID | ?{ $_.Count -gt 1 } | Select-Object -Expand Group | export-csv -notype "$ReferenceFile"

# Remove the quotes
(get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii
