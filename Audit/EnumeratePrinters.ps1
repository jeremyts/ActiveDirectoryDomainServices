<#
  This script will enumerate the print queues on a Cluster Resource and/or standalone print server and
  then output them to a file.

  Cluster Notes:

  1) Whilst we can use the Win32_Printer WMI class to get the printers from a standalone print server
     WMI is not cluster aware, so we can't access printers with WMI in a clustered environment
     http://blogs.msdn.com/b/alejacma/archive/2011/11/09/we-can-t-manage-printers-with-wmi-in-a-clustered-environment.aspx

  2) Cluster database CLUSDB: A hive under HKLM\Cluster that is physically stored as the file
     "%Systemroot%\Cluster\Clusdb". When a node joins a cluster, it obtains the cluster configuration
     from the quorum and downloads it to this local cluster database.

  3) Cluster Printers are located on the cluster service under the following key:
     HKEY_LOCAL_MACHINE\Cluster\Resources\<Resource GUID>\Parameters\Printers
     Note that this registry key structure is also available on each Node.

  4) To get the path we read the uNCName string value from under the DsSpooler subkey.
     So the data we are looking for is stored under the uNCName value under the folowing key:
     HKEY_LOCAL_MACHINE\Cluster\Resources\<Resource GUID>\Parameters\Printers\<printer>\DsSpooler\uNCName

  Release 1.0
  Written by Jeremy@jhouseconsulting.com 2nd May 2012
#>

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}

#---------------------Variables To Set------------------------

# Enable or Disable the enumeration of printers from Cluster Resources
$EnableClusterResource = $False

# Add each Cluster Resouce to the array. Alternatively, you can add one Node
# from each Cluster Resource to the array.
$ClusterResources = @("")

# Enable or Disable the enumeration of printers from Standalone Print Servers
$EnablePrintServer = $True

# Add each Standalone Print Server to the array
$Printservers = @("PTHMSPSM01")

# The output file
$OutFile = $(&$ScriptPath) + "\Printers.csv"

#-------------------------------------------------------------

if(Test-Path -Path "$OutFile")
{
  Remove-Item "$OutFile"
}

$array = @()

If ($EnableClusterResource -eq $true)
{
  foreach($Resource in $ClusterResources)
  {
    Write-Host -ForegroundColor Green "Enumerating print queues from the $Resource cluster resource/node. Please be patient and refer to the $OutFile when finished."
    $RegPath = "Cluster\Resources"
    $RemoteKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine", $Resource)
    $SubKey = $RemoteKey.OpenSubKey($RegPath,$false)
    If(!$SubKey){
      Return
    } else {
      $SubKeyValues = $SubKey.GetSubKeyNames()
      if($SubKeyValues)
      {
        foreach($SubKeyValue in $SubKeyValues)
        {
          $newsubkey = $RegPath + "\" + $SubKeyValue + "\Parameters\Printers"
          $SubKey = $RemoteKey.OpenSubKey($newsubkey,$false)
          If ($SubKey){
            $SubSubKeyValues = $SubKey.GetSubKeyNames()
            if($SubSubKeyValues)
            {
              foreach($SubSubKeyValue in $SubSubKeyValues)
              {
                $newsubsubkey = $newsubkey + "\" + $SubSubKeyValue + "\DsSpooler"
                $SubKey = $RemoteKey.OpenSubKey($newsubsubkey,$false)
                If ($SubKey){
                  $Values = $SubKey.GetValueNames()
                  if($Values)
                  {
                    foreach($Value in $Values)
                    {
                      If ($value -eq "uNCName") {
                        $Share = $SubKey.GetValue("$value")
                      } else {
                        $Share = ""
                      }
                      If ($value -eq "printerName") {
                        $Name = $SubKey.GetValue("$value")
                      } else {
                        $Name = ""
                      }
                      If ($value -eq "location") {
                        $Location = $SubKey.GetValue("$value")
                      } else {
                        $Location = ""
                      }
                      If ($value -eq "description") {
                        $Comment = $SubKey.GetValue("$value")
                      } else {
                        $Comment = ""
                      }
                      If ($value -eq "driverName") {
                        $DriverName = $SubKey.GetValue("$value")
                      } else {
                        $DriverName = ""
                      }
                      if($Share -ne "")
                      {
                        $output = New-Object PSObject
                        $output | Add-Member NoteProperty -Name "Name" "$Name"
                        $output | Add-Member NoteProperty -Name "Share" "$Share"
                        $output | Add-Member NoteProperty -Name "Location" "$Location"
                        $output | Add-Member NoteProperty -Name "Comment" "$Comment"
                        $output | Add-Member NoteProperty -Name "Driver" "$DriverName"
                        $array += $output
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }
}

If ($EnablePrintServer -eq $true)
{
  foreach($Printserver in $Printservers)
  {
    Write-Host -ForegroundColor Green "Enumerating print queues from the $Printserver print server. Please be patient and refer to the $OutFile when finished."
    # Get printer information
    $Printers = Get-WMIObject Win32_Printer -computername $Printserver
    foreach ($Printer in $Printers)
    {
      If ($Printer.Shared)
      {
        $Name = $Printer.Name
        $Share = "\\$Printserver\" + ($Printer.ShareName)
        $Location = $Printer.Location
        $Comment = $Printer.Comment
        $DriverName = $Printer.DriverName
        $output = New-Object PSObject
        $output | Add-Member NoteProperty -Name "Name" "$Name"
        $output | Add-Member NoteProperty -Name "Share" "$Share"
        $output | Add-Member NoteProperty -Name "Location" "$Location"
        $output | Add-Member NoteProperty -Name "Comment" "$Comment"
        $output | Add-Member NoteProperty -Name "Driver" "$DriverName"
        $array += $output
      }
    }
  }
}

$array | export-csv -notype "$OutFile" -Delimiter ';'
# Remove the quotes
(get-content "$OutFile") |% {$_ -replace '"',""} | out-file "$OutFile" -Fo -En ascii
