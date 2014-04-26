# This script will list the installed Windows Features

import-module ServerManager
$InstalledFeatures = Get-WindowsFeature | where-object {$_.Installed -eq $True}
write-host "Installed Features:"
ForEach ($Feature in $InstalledFeatures){
  $Message = " - "+$Feature.DisplayName+" ("+$Feature.Name+")"
  write-host $Message
}
