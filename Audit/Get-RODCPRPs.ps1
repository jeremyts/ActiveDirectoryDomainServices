<#
  This script will find all Read-Only Domain Controllers (RODCs) and report
  on the Allowed and Denied Password Replication Policies (PRPs) on each one.

  Script Name: Get-RODCPRPs.ps1
  Release 1.0
  Written by Jeremy@jhouseconsulting.com 29/1/2014

  References:
  - http://technet.microsoft.com/en-us/library/cc730883(v=ws.10).aspx
  - http://blogs.technet.com/b/heyscriptingguy/archive/2013/12/17/use-powershell-to-work-with-rodc-accounts.aspx
  - http://gallery.technet.microsoft.com/scriptcenter/Get-ADRodcAuthenticatedNotR-daf51490
  - http://gunnalag.com/2011/12/25/rodc-filtered-attribute-set-credential-caching-and-the-authentication-process-with-an-rodc/
  - http://blogs.dirteam.com/blogs/paulbergson/archive/2010/09/22/rodc-password-replication-group-management.aspx
  - http://technet.microsoft.com/nl-nl/library/cc835090(v=ws.10).aspx
  - http://blogs.metcorpconsulting.com/tech/?p=1096
  - Refer to the section on Password Replication Policies in chapter 9 of the
    O'Reilly Active Directoy 5th Edition book under the Read-Only Domain
    Controllers section.
#>

#-------------------------------------------------------------

# Import the Active Directory Module
Import-Module ActiveDirectory -WarningAction SilentlyContinue
if ($Error.Count -eq 0) {
  #Write-Host "Successfully loaded Active Directory Powershell's module`n" -ForeGroundColor Green
} else {
  Write-Host "Error while loading Active Directory Powershell's module : $Error`n" -ForeGroundColor Red
  exit
}

#-------------------------------------------------------------

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$ReferenceFile = $(&$ScriptPath) + "\RODCPRPReport.csv"

$array = @()

$RODCs = Get-ADDomainController -Filter {IsReadOnly -eq $true}

ForEach ($RODC in $RODCs) {

  $Allowed = Get-ADDomainControllerPasswordReplicationPolicy -Allowed -Identity $RODC

  $Denied = Get-ADDomainControllerPasswordReplicationPolicy -Denied -Identity $RODC

  $AllowArray = @()
  $DeniedArray = @()

  ForEach ($Allow in $Allowed) {
    $AllowArray += $Allow.Name
  }
  ForEach ($Deny in $Denied) {
    $DeniedArray += $Deny.Name
  }

  $output = New-Object PSObject
  $output | Add-Member NoteProperty -Name "RODC" $RODC.Name
  $output | Add-Member NoteProperty -Name "Allowed" ($AllowArray -join "|")
  $output | Add-Member NoteProperty -Name "Denied" ($DeniedArray -join "|")
  $array += $output

}

# Write-Output $array | Format-Table
$array | export-csv -notype -path "$ReferenceFile"

# Remove the quotes
(get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii
