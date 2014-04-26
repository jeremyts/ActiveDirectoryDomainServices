<#
  The following script finds all GPOs in the domain that have been modified this month.
  It then takes these GPOs backs them up and generates a settings report  for each.
  Finally it lists out all of the GPOs that were backed up.

  Syntax examples:
    Full backup:
      GPOBackup.ps1
    Backup all GPOs modified within the last 7 days:
      GPOBackup.ps1 -Days 7

  It is based on the following two scripts:
    1) PowerShell Script: Backup all GPOs that have been modified this month by the Microsoft Group Policy Team:
       http://blogs.technet.com/b/grouppolicy/archive/2009/03/26/powershell-script-backup-all-gpos-that-have-been-modified-this-month.aspx
    2) Backing up Group Policy Objects using Windows PowerShell by Jan Egil Ring:
       http://blog.powershell.no/2010/06/15/backing-up-group-policy-objects-using-windows-powershell/

  Release 1.2
  Modified by Jeremy@jhouseconsulting.com 7th June 2011
  Modified by Jeremy@jhouseconsulting.com 13th September 2013
#>

#-------------------------------------------------------------
param([switch]$Days,[int]$NumberOfDays,[String]$BackupLocation,[String]$LogFile);

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}

if ([String]::IsNullOrEmpty($BackupLocation))
{
    # Set the paths
    $BackupLocation = $(&$ScriptPath) + "\GPOs"
}

if ([String]::IsNullOrEmpty($LogFile))
{
    $LogFile = $(&$ScriptPath) + "\ExportAllGPOsLog.txt";
}
set-content $LogFile $NULL;

#-------------------------------------------------------------
Write-Host -ForegroundColor Green "Importing the PowerShell modules..."

# Import the Active Directory Module
Import-Module ActiveDirectory -WarningAction SilentlyContinue
if($Error.Count -eq 0) {
   Write-Host "Successfully loaded Active Directory Powershell's module" -ForeGroundColor Green
}else{
   Write-Host "Error while loading Active Directory Powershell's module : $Error" -ForeGroundColor Red
   exit
}

# Import the Group Policy Module
Import-Module GroupPolicy -WarningAction SilentlyContinue
if($Error.Count -eq 0) {
   Write-Host "Successfully loaded Group Policy Powershell's module" -ForeGroundColor Green
}else{
   Write-Host "Error while loading Group Policy Powershell's module : $Error" -ForeGroundColor Red
   exit
}

#-------------------------------------------------------------
# Get the Domain DNS name
$DNSRoot = (Get-ADDomain).DNSRoot

# Set the paths
$GPOPath = $BackupLocation + "\Backups"
$ReportPath = $BackupLocation + "\Reports"

#-------------------------------------------------------------
# Create the folders, if they do not already exist
if (Test-Path -path $BackupLocation)
{
  # Delete existing backup
  remove-item $BackupLocation\* -force -recurse -confirm:$false
}
if (!(Test-Path -path $BackupLocation))
{
  New-Item $BackupLocation -type directory | out-Null
}
if (!(Test-Path -path $GPOPath))
{
  New-Item $GPOPath -type directory | out-Null
}
if (!(Test-Path -path $ReportPath))
{
  New-Item $ReportPath -type directory | out-Null
}
write-host " "

# Get the current date
get-Date | Out-File $LogFile

# Get the GPOs
if ($Days) {
  If ($NumberOfDays -eq 0) {$NumberOfDays = 1}
  $Timespan = (Get-Date).AddDays(-$NumberOfDays)
  $GPOs = Get-GPO -domain $DNSRoot -all | Where-Object {$_.ModificationTime -gt $Timespan}
} else {
  $GPOs = Get-GPO -domain $DNSRoot -all
}

$Count = $GPOs | Measure-Object | %{$_.Count}

If ($Count -gt 0) {

  # Count GPOs to be backed up
  $Message = "Backing up " + $Count + " GPOs..."
  write-host -ForeGroundColor Green $Message
  $Message | Out-File $LogFile -append

  # Loop through all GPOs
  Foreach ($GPO in $GPOs) { 
    $Message = " - " + $GPO.DisplayName
    write-host -ForeGroundColor Green $Message
    $Message | Out-File $LogFile -append

    # Backup the GPO to the specified path 
    $GPOBackup = backup-GPO $GPO.DisplayName -path $GPOPath -Domain $DNSRoot

    # Generate a report of the backed up settings.
    $ReportName = $ReportPath + "\" + $GPO.Displayname + "_" + $GPO.ModificationTime.Month + "-"+ $GPO.ModificationTime.Day + "-" + $GPO.ModificationTime.Year + "_" +  $GPOBackup.Id + ".html" 
    get-GPOReport -Name $GPO.DisplayName -path $ReportName -ReportType HTML 
  }

  $Message = " "
  write-host -ForeGroundColor Green $Message
  $Message | Out-File $LogFile -append
  $Message = "Go to the '" + $ReportPath + "' folder to view the settings reports for the backed up GPOs."
  write-host -ForeGroundColor Green $Message
  $Message | Out-File $LogFile -append

} else {

  $Message = "There are no GPOs to be backed up that have been modified in the last $NumberOfDays days."
  write-host -ForeGroundColor Green $Message
  $Message | Out-File $LogFile -append

}
