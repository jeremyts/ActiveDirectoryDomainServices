<#
 This script will compare the 

 Shay Levy posted the Get the AD site name of a computer function here:
 http://www.powershellmagazine.com/2013/04/23/pstip-get-the-ad-site-name-of-a-computer/

 Microsoft posted a method for retrieving the DynamicSiteName here:
 Problem retrieving Value for DynamicSiteName from Registry using PS: http://support.microsoft.com/kb/2801452

  https://github.com/joethemongoose/PowerCLI/blob/master/Export-VMs-With-SiteCode.ps1

#>

#-------------------------------------------------------------
# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$OnlineFile = $(&$ScriptPath) + "\SiteNameConsistency-online.csv"
$OfflineFile = $(&$ScriptPath) + "\SiteNameConsistency-offline.csv"

#-------------------------------------------------------------

Import-Module ActiveDirectory 

# Get Domain Controllers
$Computers = Get-ADDomainController -Filter * | Sort-Object Name

# Get All Servers
#$Computers = Get-ADComputer -Filter * -Properties Name,Operatingsystem | Where-Object {$_.Operatingsystem -like "*server*"} | Sort-Object Name

# Get All Authorized DHCP Servers
$defaultNC = ([ADSI]"LDAP://RootDSE").defaultNamingContext.Value
$configurationNC = "cn=configuration," + $defaultNC
$AuthorizedDHCPServers = Get-ADObject -SearchBase $configurationNC -Filter "objectclass -eq 'dhcpclass' -AND Name -ne 'dhcproot'"

#-------------------------------------------------------------

function Get-ComputerSite($ComputerName)
{
  $site = nltest /server:$ComputerName /dsgetsite 2>$null
  if($LASTEXITCODE -eq 0){ $site[0] }
}

function Get-ComputerSiteValue($ComputerName,$Value)
{
  $ValueData = ""
  $key = "System\CurrentControlSet\Services\NetLogon\Parameters"
  $type = [Microsoft.Win32.RegistryHive]::LocalMachine
  Try
    {
      $regKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($type, $ComputerName)
      $regKey = $regKey.OpenSubKey($key)
      $ValueData= $regKey.GetValue($($Value)).Split([char]0)[0]
    }
  Catch
    {
      write-host -ForegroundColor Red "- $($Value): Could Not Connect to Remote Registry!"
    }
  $ValueData
}

$onlinearray = @()
$offlinearray = @()

ForEach($Computer in $Computers){
  $ComputerError = "$false"
  $ComputerName = $Computer.Name
  if (Test-Connection -Cn $Computer.Name -BufferSize 16 -Count 1 -ea 0 -quiet) {
    write-Host -ForegroundColor Green "Checking Site for $ComputerName"
    $ComputerSite = Get-ComputerSite $ComputerName
    $ComputerDynamicSiteName = Get-ComputerSiteValue $ComputerName DynamicSiteName
    $ComputerSiteName = Get-ComputerSiteValue $ComputerName SiteName

    $DynamicSiteName = $False
    $SiteName = $False

    If ($ComputerDynamicSiteName -ne "") {$DynamicSiteName = $True}
    If ($ComputerSiteName -ne "") {$SiteName = $True}

    If ($SiteName -eq $True) {
      If ($ComputerSite -eq $ComputerSiteName) {
        write-host -ForegroundColor Green "- Sites match"
      } Else {
        write-host -ForegroundColor Red "- Sites do not match"
      }
    } ElseIf ($DynamicSiteName -eq $True) {
        If ($ComputerSite -eq $ComputerDynamicSiteName) {
          write-host -ForegroundColor Green "- Sites match"
        } Else {
          write-host -ForegroundColor Red "- Sites do not match"
        }
    } Else {
      write-host -ForegroundColor Red "- Both the DynamicSiteName and SiteName registry values are missing."
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
