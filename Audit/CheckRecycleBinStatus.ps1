# Check the Active Directory Recycle Bin Status

Import-Module ActiveDirectory

$ForestName = (Get-ADForest).Name
$ForestDomainNamingMaster = $((Get-ADForest -Current LocalComputer).DomainNamingMaster)

$ForestInfo = Get-ADOptionalFeature "Recycle Bin Feature" -server $ForestDomainNamingMaster
If ($ForestInfo.EnabledScopes -eq $NULL -OR $ForestInfo.EnabledScopes -eq "") {
  write-host -ForegroundColor yellow "The Recycle Bin Feature is not enabled on the $ForestName forest."
} Else {
  write-host -ForegroundColor green "The Recycle Bin Feature is enabled on the $ForestName forest with the following scopes:"
  ForEach ($scope in $ForestInfo.EnabledScopes) {
    write-host -ForegroundColor green " - $scope"
  }
}
