<#
  This script will find all Windows servers with the DNS Server service installed,
  report on it's state and any local zones files (non-Active Directory integrated).

  Note that for servers we filter out Cluster Name Objects (CNOs) and
  Virtual Computer Objects (VCOs) by checking the objects serviceprincipalname
  property for a value of MSClusterVirtualServer. The CNO is the cluster
  name, whereas a VCO is the client access point for the clustered role.
  These are not actual computers, so we exlude them to assist with
  accuracy.

  Script Name: FindDNSServersAndLocalZones.ps1
  Release 1.0
  Written by Jeremy@jhouseconsulting.com 10/01/2014

#>

#-------------------------------------------------------------
# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$OnlineFile = $(&$ScriptPath) + "\DNSServers-online.csv"
$OfflineFile = $(&$ScriptPath) + "\DNSServers-offline.csv"

#-------------------------------------------------------------

Import-Module ActiveDirectory 

# Get Domain Controllers
#$Computers = Get-ADDomainController -Filter * | Sort-Object Name

# Get All Servers filtering out Cluster Name Objects (CNOs) and Virtual computer Objects (VCOs) 
$Computers = Get-ADComputer -Filter * -Properties Name,Operatingsystem,servicePrincipalName | Where-Object {($_.Operatingsystem -like '*server*') -AND !($_.serviceprincipalname -like '*MSClusterVirtualServer*')} | Sort-Object Name

#-------------------------------------------------------------

$onlinearray = @()
$offlinearray = @()

$Count = 0
$Count = $Computers.Count

write-Host -ForegroundColor Green "There are $Count servers to process.`n"

ForEach($Computer in $Computers){
  $ComputerError = "$false"
  $ComputerName = $Computer.Name
  if (Test-Connection -Cn $Computer.Name -BufferSize 16 -Count 1 -ea 0 -quiet) {
    write-Host -ForegroundColor Green "Checking for DNS Server service on $ComputerName"
    $ServiceName = "DNS"
    Try {
      $serviceObj = Get-Service -ComputerName $ComputerName | ?{ $_.ServiceName -eq $serviceName } | Select-Object Name, Status
      If ($serviceObj -ne $NULL) {
        If ($serviceObj.Status -eq "Running") {
          write-host -ForegroundColor green "- Service found in a $($serviceObj.Status) state."
        } Else {
          write-host -ForegroundColor red "- Service found in a $($serviceObj.Status) state."
        }
        # Path to DNS
        $path = "\\$ComputerName\admin$\System32\dns"
        # Testing the $path
        IF ((Test-Path -Path $path) -and ((Get-Item -Path $path).Length -ne $null)) {

          IF ((Get-ChildItem "$path\*.dns" -Exclude "CACHE.DNS"| Measure-Object).Count -gt 0) {
            # Get all zone file (*.dns) and exclude the CACHE.DNS file.
            write-host -ForegroundColor yellow "- Local zones files found."
            $ZoneFiles = Get-ChildItem "$path\*.dns" -Exclude "CACHE.DNS"
            $LocalZones = ""
            ForEach ($ZoneFile in $ZoneFiles) {
              If ($LocalZones -eq "" ) {
                $LocalZones = $ZoneFile.Name
              } Else {
                $LocalZones = $LocalZones +";"+ $ZoneFile.Name
              }
            }
          } Else {
            write-host -ForegroundColor yellow "- No local zones files found."
            $LocalZones = "none found."
          }
        } Else {
          $ComputerError = "$true"
          $ErrorDescription = "Not reachable via the $path path."
          write-Host -ForegroundColor Red "- $ErrorDescription"
        }
        $output = New-Object PSObject
        $output | Add-Member NoteProperty -Name "ComputerName" $ComputerName
        $output | Add-Member NoteProperty -Name "Service" $serviceObj.Name
        $output | Add-Member NoteProperty -Name "Status" $serviceObj.Status
        $output | Add-Member NoteProperty -Name "LocalZoneFiles" $LocalZones
        $onlinearray += $output
      } Else {
        write-host -ForegroundColor yellow "- $ServiceName service not installed"
      }
    }
    Catch {
      $ComputerError = "$true"
      $ErrorDescription = "Error connecting using the Get-Service cmdlet."
      write-Host -ForegroundColor Red "- $ErrorDescription"
    }
  } Else {
    $ComputerError = "$true"
    $ErrorDescription = "Unable to ping server"
    write-Host -ForegroundColor Red "$ComputerName is offline"
  }
  if ($ComputerError -eq $true) {
    $outputbad = New-Object PSObject
    $outputbad | Add-Member NoteProperty ComputerName $ComputerName
    $outputbad | Add-Member NoteProperty ErrorDescription $ErrorDescription
    $offlinearray += $outputbad
  }
}
$onlinearray | export-csv -notype "$OnlineFile"
$offlinearray | export-CSV -notype "$OfflineFile"

# Remove the quotes
(get-content "$OnlineFile") |% {$_ -replace '"',""} | out-file "$OnlineFile" -Fo -En ascii
(get-content "$OfflineFile") |% {$_ -replace '"',""} | out-file "$OfflineFile" -Fo -En ascii
