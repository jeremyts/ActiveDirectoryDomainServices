<#
  This script will create your ADMX Central Store by using a master
  source and the local store on existing management servers.

  Script Name: CreateADMXCentralStore.ps1
  Release 1.3
  Modified by Jeremy@jhouseconsulting.com 23rd February 2014
  Written by Jeremy@jhouseconsulting.com 14th February 2014

  Notes:
  - I've found that some ADML files are more language generic.
    For Example: The OpsMgs (SCOM) HealthService.adml is located
      under the "EN" folder instead of the "en-us" folder.
  - I've found that some ADML files are accompanied by a dll.
    For Example: The OpsMgr (SCOM) HealthService.adml also has a
      HealthServiceADML.Dll.
    I've not been able to find any information on this, so have
    made sure this script copies across any existing dlls that
    accompany the ADML.

  ADMX Central Store references:
  - For further information refer to Managing Group Policy ADMX Files Step-by-Step Guide:
    http://msdn.microsoft.com/en-us/library/bb530196.aspx
  - How to create a Central Store for Group Policy Administrative Templates in Window Vista
    http://support.microsoft.com/kb/929841

  Compare-Object cmdlet limitations:
  - The output of the compare-object cmdlet may be incorrect if
    you're comparing collections of more than 11 elements. To
    address this issue we set the SyncWindow parameter to half the
    size of the smaller object.
    http://dmitrysotnikov.wordpress.com/2008/06/06/compare-object-gotcha/

  Copy-Item cmdlet limitations:
  - The Copy-Item cmdlet is quite limiting in its behavior. There
    is no "overwrite if newer", or "keep newest version" parameter.
    If the destination file exists, it will not be overwritten
    unless you use the -force paratemeter. So to work around this
    I've added a check to compare the lastwritetime property of
    the source and destination files to decide on which one is the
    newer file.

  Get-ChildItem cmdlet confusion:
  - The Include parameter is effective only when the command includes
    the "-recurse" parameter OR the path leads to the contents of a
    directory such as C:\Windows\*

#>

#-------------------------------------------------------------

# Set this to the location where your ADMX master files are kept.
# If you use a relative path, the script will prepend the script
# path to create an absolute path.
$MasterReferenceLocation = "ADMXCentralStore\Used"

# Set array to the language so that we copy across the relevant
# ADML files. Note that but setting this to an * (asterix), it
# will copy the ADML files from all language folders.
$languages = @("EN","en-us")

# Set this to the servers that you want to use to build the ADMX
# Central Store. They are typically the servers that contain the
# latest versions of ADMX files, as well as the customised and
# 3rd party ones you're currently using in any GPOs.
$SourceServers = @("dc01","ctx01","adm01")

#-------------------------------------------------------------

# Get the current domain name 
$FQDN = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().name

If (!($MasterReferenceLocation.Contains(':\')) -AND !($MasterReferenceLocation.Contains('\\'))) {
  $ScriptPath = (Split-Path -Path ((Get-Variable -Name MyInvocation).Value).MyCommand.Path)
  If (!($MasterReferenceLocation.StartsWith('\'))) {
    $MasterReferenceLocation = $ScriptPath + "\" + $MasterReferenceLocation
  } Else {
    $MasterReferenceLocation = $ScriptPath + $MasterReferenceLocation
  }
}

# We can either prepend of append the $MasterReferenceLocation to the $SourceServers
# array. If we append it, we should then reverse the array so that it's processed first.
$SourceServers = ,$MasterReferenceLocation + $SourceServers
#$SourceServers += $MasterReferenceLocation
#[array]::Reverse($SourceServers)

write-host -ForegroundColor green "`nCreating or adding to the ADMX Central Store..."

[string]$t = "\\$FQDN\SYSVOL\$FQDN\Policies\PolicyDefinitions"
If (-not(Test-Path -Path "$t")) {
  write-host -ForegroundColor green "`n`tCreating the '$t' folder..."
  New-Item -Path "$t" -ItemType Directory | out-Null
} else {
  write-host -ForegroundColor yellow "`n`tThe '$t' folder already exists."
}

$target = Get-ChildItem $t | Where {$_.psIsContainer -eq $false}

ForEach ($SourceServer in $SourceServers ) {
  If ($SourceServer.Contains('\')) {
    [string]$s = $SourceServer
  } else {
    If ($SourceServer -ne ($env:computername)) {
      [string]$s = "\\" + $SourceServer + "\admin$\PolicyDefinitions"
    } else {
      [string]$s = "$($env:systemroot)\PolicyDefinitions"
    }
  }

  write-host -ForegroundColor green "`n`tProcessing source files from $SourceServer..."

  If (Test-Path -Path $s) {
    $source = Get-ChildItem $s | Where {$_.psIsContainer -eq $false}
    If (($languages -eq "*") -OR ($languages -contains "*")) {
      $languages = @()
      $folders = Get-ChildItem $s | Where {$_.psIsContainer -eq $true}
      ForEach ($folder in $folders) {
        $languages += $folder.name
        If (-not(Test-Path -Path "$t\$($folder.name)")) {
          write-host -ForegroundColor green "`t- Creating the '$t\$($folder.name)' folder..."
          New-Item -Path "$t\$($folder.name)" -ItemType Directory | out-Null
        } else {
          write-host -ForegroundColor yellow "`t- The '$t\$($folder.name)' folder already exists."
        }
      }
    } else {
      ForEach ($language in $languages) {
        If (-not(Test-Path -Path "$t\$language")) {
          write-host -ForegroundColor green "`t- Creating the '$t\$language' folder..."
          New-Item -Path "$t\$language" -ItemType Directory | out-Null
        } else {
          #write-host -ForegroundColor yellow "`t- The '$t\$language' folder already exists."
        }
      }
    }

    # Set the SyncWindow to half the size of the smaller object
    $TargetCount = ($target | Measure-object).Count
    $SourceCount = ($source | Measure-object).Count
    If ($TargetCount -le $SourceCount) {
      $SyncWindow = $TargetCount / 2
    } Else {
      $SyncWindow = $SourceCount / 2
    }
    If ($SyncWindow -gt 5) {
      # Use the modulus operator to divide it by 2 to determine if it's an
      # odd or even number. An even number will not have a remainer of 0,
      # whilst an odd number has a remainder of 0.5, so we use the [int]
      # DataType to round it down to a A 32-bit signed whole number.
      If (($SyncWindow % 2) -ne 0) {
        $SyncWindow = [int]$SyncWindow
      }
    } Else {
      $SyncWindow = 5
    }

    If ($TargetCount -eq 0) {
      # If there are no files in the target folder, the Compare-Object cmdlet
      # will fail with the following error:
      # Cannot bind argument to parameter 'DifferenceObject' because it is null.
      # To work around this issue we create a starter file, re-create the
      # target object and then delete the starter file. Now we have a difference
      # object that is not null.
      New-Item $t\StarterFile.txt -type file | out-Null
      $target = Get-ChildItem $t | Where {$_.psIsContainer -eq $false}
      Remove-Item $t\StarterFile.txt | out-Null
    }

    $results = @(Compare-Object -ReferenceObject $source -DifferenceObject $target -SyncWindow $SyncWindow |Where-Object { $_.SideIndicator -eq '<=' } )
    If (($results | Measure-object).Count -ne 0) {
      write-host -ForegroundColor green "`t- Processing results from $SourceServer..."
      foreach($result in $results) {
        If (!($result.InputObject.PSIsContainer)) {
          #$SourceADMXFile = "$($result.InputObject.FullName)"
          $SourceADMXFile = "$($result.InputObject.DirectoryName)\$($result.InputObject.BaseName).admx"
          $ADMLFilePresent = $False
          ForEach ($language in $languages) {
            $SourceADMLFile = "$($result.InputObject.DirectoryName)\$language\$($result.InputObject.BaseName).adml"
            $SourceADMLLibraryFile = "$($result.InputObject.DirectoryName)\$language\$($result.InputObject.BaseName)ADML.dll"
            If (Test-Path -Path $SourceADMLFile) {
              $ADMLFilePresent = $True
              $DestinationADMLFile = "$t\$language\$($result.InputObject.BaseName).adml"
              $DestinationADMLLibraryFile = "$t\$language\$($result.InputObject.BaseName)ADML.dll"
              if (Test-Path -Path $DestinationADMLFile) {
                $SourceADMLFileTime = [datetime](Get-ItemProperty -Path $SourceADMLFile -Name LastWriteTime).lastwritetime
                $DestinationADMLFileTime = [datetime](Get-ItemProperty -Path $DestinationADMLFile -Name LastWriteTime).lastwritetime
                If ($SourceADMLFileTime -gt $DestinationADMLFileTime ) {
                  write-host -ForegroundColor green "`t- Overwriting from source: $SourceADMLFile"
                  copy-item "$SourceADMLFile" -destination "$t\$language" -force
                } else {
                  write-host -ForegroundColor yellow "`t`t- Destination file is newer: $SourceADMLFile"
                }
              } else {
                write-host -ForegroundColor green "`t`t- Copying from source: $SourceADMLFile"
                copy-item "$SourceADMLFile" -destination "$t\$language"
              }
              # Copy a matching ADML library file if present.
              If ($ADMLFilePresent -AND (Test-Path -Path $SourceADMLLibraryFile)) {
                if (Test-Path -Path $DestinationADMLLibraryFile) {
                  $SourceADMLLibraryFileTime = [datetime](Get-ItemProperty -Path $SourceADMLLibraryFile -Name LastWriteTime).lastwritetime
                  $DestinationADMLLibraryFileTime = [datetime](Get-ItemProperty -Path $DestinationADMLLibraryFile -Name LastWriteTime).lastwritetime
                  If ($SourceADMLLibraryFileTime -gt $DestinationADMLLibraryFileTime ) {
                    write-host -ForegroundColor green "`t- Overwriting from source: $SourceADMLLibraryFile"
                    copy-item "$SourceADMLLibraryFile" -destination "$t\$language" -force
                  } else {
                    write-host -ForegroundColor yellow "`t`t- Destination file is newer: $SourceADMLLibraryFile"
                  }
                } else {
                  write-host -ForegroundColor green "`t`t- Copying from source: $SourceADMLLibraryFile"
                  copy-item "$SourceADMLLibraryFile" -destination "$t\$language"
                }
              }
            }
          }
          # Only copy the ADMX if an ADML is present.
          If ($ADMLFilePresent) {
            $DestinationADMXFile = "$t\$($result.InputObject.BaseName).admx"
            if (Test-Path -Path $DestinationADMXFile) {
              $SourceADMXFileTime = [datetime](Get-ItemProperty -Path $SourceADMXFile -Name LastWriteTime).lastwritetime
              $DestinationADMXFileTime = [datetime](Get-ItemProperty -Path $DestinationADMXFile -Name LastWriteTime).lastwritetime
              If ($SourceADMXFileTime -gt $DestinationADMXFileTime ) {
                write-host -ForegroundColor green "`t`t- Overwriting from source: $SourceADMXFile"
                copy-item "$SourceADMXFile" -destination "$t" -force
              } else {
                write-host -ForegroundColor yellow "`t`t- Destination file is newer: $SourceADMXFile"
              }
            } else {
              write-host -ForegroundColor green "`t`t- Copying from source: $SourceADMXFile"
              copy-item "$SourceADMXFile" -destination "$t"
            }
          } else {
            write-host -ForegroundColor yellow "`t- No matching ADML file was found for: $SourceADMXFile"
          }
        }
      }
    } else {
      write-host -ForegroundColor yellow "`t- No files to be added from $SourceServer."
    }
  } else {
    write-host -ForegroundColor red "`t- The $SourceServer location does not exist."
  }
}

write-host -ForegroundColor green "`nSummary:"
$TotalADMX = (Get-ChildItem $t | Where {$_.psIsContainer -eq $false}| Measure-object).Count
write-host -ForegroundColor green "- Total ADMX files in '$t': $TotalADMX"
$folders = Get-ChildItem $t | Where {$_.psIsContainer -eq $true}
ForEach ($folder in $folders) {
  $language = $folder.name
  $TotalADML = (Get-ChildItem "$t\$language\*" -include *.adml | Measure-object).Count
  If ($TotalADML -ne 0) {
    write-host -ForegroundColor green "- Total ADML files in '$t\$language': $TotalADML"
    $TotalADMLDLL = (Get-ChildItem "$t\$language\*" -include *adml.dll | Measure-object).Count
    If ($TotalADMLDLL -ne 0) {
      write-host -ForegroundColor green "- Total ADML dll files in '$t\$language': $TotalADMLDLL"
    }
  }
}

write-host -ForegroundColor green "`nFinished."
