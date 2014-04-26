# This script will raise the Active Directory Functional Level.
# The functionality level should already be set to Windows 2008 R2 during the DCPromo process.

$RaiseTo = "windows2008R2Forest"

Import-Module ActiveDirectory

$ForestInfo = Get-ADForest
$ForestName = $ForestInfo.Name
$RaiseFrom = $ForestInfo.ForestMode

If ($RaiseFrom -ne $RaiseTo) {
  write-host -ForegroundColor green "Raising the $ForestName forest from $RaiseFrom to $RaiseTo mode..."
  # Use either one of the following lines:
  Set-ADForestMode -ForestMode $RaiseTo –confirm:$false
  #Set-ADForestMode –Identity $ForestName -ForestMode $RaiseTo –confirm:$false
} Else {
  write-host -ForegroundColor yellow "The $ForestName forest is already set to $RaiseFrom functionality level."
}
