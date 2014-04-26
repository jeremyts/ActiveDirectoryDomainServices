<#
  This script will find all Windows servers with the DHCP Server service installed,
  and report on it's state, if a local database (DHCP.mdb) exists, and if it's
  authorized in Active Directory.

  Note that for servers we filter out Cluster Name Objects (CNOs) and
  Virtual Computer Objects (VCOs) by checking the objects serviceprincipalname
  property for a value of MSClusterVirtualServer. The CNO is the cluster
  name, whereas a VCO is the client access point for the clustered role.
  These are not actual computers, so we exlude them to assist with
  accuracy.

  Thanks to Michael B Smith for the Get-AuthorizedDHCPServers function.

  There are multiple ways of getting information on a remote service:
  - Get-Service with the –ComputerName parameter (RPC)
  - Get-WMIObject with a –ComputerName parameter (WMI)
  - Invoke-Command to execute Get-Service remotely (WinRM/WSMAN)
  - Get-CimInstance with the –ComputerName parameter (WinRM/WSMAN)

  The Get-CimInstance is only available from v3.

  I have found that using WinRM is hit and miss, as it's not always correctly
  setup across a server fleet, and therefore you receive the following errors:

  - Get-CimInstance : WinRM cannot complete the operation. Verify that the specified
    computer name is valid, that the computer is accessible over the network, and
    that a firewall exception for the WinRM service is enabled and allows access from
    this computer. By default, the WinRM firewall exception for public profiles
    limits access to remote computers within the same local subnet.
  - New-PSSession : [<computername>] Connecting to remote server <computername> failed
    with the following error message : WinRM cannot complete the operation. Verify that
    the specified computer name is valid, that the computer is accessible over the
    network, and that a firewall exception for the WinRM service is enabled and allows
    access from this computer. By default, the WinRM firewall exception for public
    profiles limits access to remote computers within the same local subnet. For more
    information, see the about_Remote_Troubleshooting Help topic.

  Ironically I get more reliability using the Get-Service or Get-WMIObject cmdlets.

  Getting things to work correctly using WinRM/WSMAN requires some cleverness:
  - http://jdhitsolutions.com/blog/2013/04/get-ciminstance-from-powershell-2-0/
  - http://richardspowershellblog.wordpress.com/2012/01/14/cim-cmdlets-and-remote-machines/
  - http://blogs.technet.com/b/josebda/archive/2010/04/02/comparing-rpc-wmi-and-winrm-for-remote-server-management-with-powershell-v2.aspx

  Script Name: FindDHCPServers.ps1
  Release 1.2
  Written by Jeremy@jhouseconsulting.com 15/01/2014
  Modified by Jeremy@jhouseconsulting.com 26/03/2014

#>

#-------------------------------------------------------------
# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$OnlineFile = $(&$ScriptPath) + "\DHCPServers-online.csv"
$OfflineFile = $(&$ScriptPath) + "\DHCPServers-offline.csv"

#-------------------------------------------------------------

Import-Module ActiveDirectory 

# Get Domain Controllers
#$Computers = Get-ADDomainController -Filter * | Sort-Object Name

# Get All Servers filtering out Cluster Name Objects (CNOs) and Virtual computer Objects (VCOs) 
$Computers = Get-ADComputer -Filter * -Properties Name,Operatingsystem,servicePrincipalName | Where-Object {($_.Operatingsystem -like '*server*') -AND !($_.serviceprincipalname -like '*MSClusterVirtualServer*')} | Sort-Object Name

$Filter = "(&(objectCategory=computer)(operatingSystem=*server*)(!(servicePrincipalName=*MSClusterVirtualServer*)))"

# Get All Authorized DHCP Servers
# This has been purposely commented out as I am now using the
# Get-AuthorizedDHCPServers function.
#$defaultNC = ([ADSI]"LDAP://RootDSE").defaultNamingContext.Value
#$configurationNC = "cn=configuration," + $defaultNC
#$AuthorizedDHCPServers = Get-ADObject -SearchBase "CN=NetServices,CN=Services,$configurationNC" -Filter "objectclass -eq 'dhcpclass' -AND Name -ne 'dhcproot'"

#-------------------------------------------------------------
function Get-AuthorizedDHCPServers
{
  # The configNC is replicated to all DCs in a forest, similar
  # to the schema, so this function will get all authorized
  # DHCP servers in the forest. 
  $adsi = [ADSI]( "LDAP://RootDSE" )
  $configNC = $adsi.configurationNamingContext.Item( 0 )
  $names = @()
  $cn = [ADSI]("LDAP://CN=NetServices,CN=Services," + $configNC )
  foreach( $object in $cn.children )
  {
    if( $object.Properties[ 'dhcpIdentification' ] -ne $null )
    {
      if( $object.dhcpIdentification -eq 'DHCP Server Object' )
      {
        $obj = New-Object PSObject
        $obj | Add-Member NoteProperty -Name "DistinguishedName" $object.distinguishedName[0]
        $obj | Add-Member NoteProperty -Name "Name" $object.Name[0]
        $names += $obj
      }
    }
  }
  $names
}

# Get All Authorized DHCP Servers
$AuthorizedDHCPServers = Get-AuthorizedDHCPServers

#-------------------------------------------------------------

$onlinearray = @()
$offlinearray = @()

ForEach($Computer in $Computers){
  $ComputerError = "$false"
  #$ComputerName = $Computer.Name
  $ComputerName = $Computer.DNSHostName
  if (Test-Connection -Cn $ComputerName -BufferSize 16 -Count 1 -ea 0 -quiet) {
    write-Host -ForegroundColor Green "Checking for DHCP Server service on $ComputerName"
    $ServiceName = "DHCPServer"
    Try {
      # If using v3, we can use Get-CimInstance. otherwise we need to use Get-Service.
      If ($PSVersionTable.PSVersion.Major -eq 3) {
        $serviceObj = Get-CimInstance Win32_Service -Computer $ComputerName | ?{ $_.Name -eq $serviceName } | Select-Object Name, State
      } else {
        #$sessions = New-PSSession -ComputerName $ComputerName
        #$serviceObj = Invoke-Command -Session $sessions -ScriptBlock {Get-Service | ?{ $_.ServiceName -eq $serviceName } | Select-Object -Property Name, @{Name="State";Expression={$_.Status}}}
        #Remove-PSSession $sessions
        $serviceObj = Get-Service -ComputerName $ComputerName | ?{ $_.ServiceName -eq $serviceName } | Select-Object -Property Name, @{Name="State";Expression={$_.Status}}
      }
      If ($serviceObj -ne $NULL) {
        $State = $serviceObj.State
        If ($State -eq "Running") {
          write-host -ForegroundColor green "- Service found in a $State state."
        } Else {
          write-host -ForegroundColor red "- Service found in a $State state."
        }
        # Path to DHCP
        $path = "\\$ComputerName\admin$\System32\dhcp"
        # Testing the $path
        IF ((Test-Path -Path $path) -and ((Get-Item -Path $path).Length -ne $null)) {
          IF ((Get-ChildItem "$path\Dhcp.mdb" | Measure-Object).Count -gt 0) {
            write-host -ForegroundColor green "- Database file found."
            $DatabaseExists = $True
          } Else {
            write-host -ForegroundColor red "- Database file not found."
            $DatabaseExists = $False
          }
        } Else {
          $ComputerError = "$true"
          $ErrorDescription = "Not reachable via the $path path."
          write-Host -ForegroundColor Red "$ErrorDescription"
        }
        $ISAuthorized = $False
        ForEach ($AuthorizedDHCPServer in $AuthorizedDHCPServers) {
          If (($AuthorizedDHCPServer.Name).ToLower().Contains($ComputerName.ToLower())) {
            $ISAuthorized = $True
            write-host -ForegroundColor green "- Authorized in Active Directory."
          }
        }
        If ($ISAuthorized -eq $False) {
          write-host -ForegroundColor red "- Not authorized in Active Directory."
        }
        $output = New-Object PSObject
        $output | Add-Member NoteProperty -Name "ComputerName" $ComputerName
        $output | Add-Member NoteProperty -Name "Service" $serviceObj.Name
        $output | Add-Member NoteProperty -Name "State" $State
        $output | Add-Member NoteProperty -Name "DatabaseExists" $DatabaseExists
        $output | Add-Member NoteProperty -Name "ISAuthorized" $ISAuthorized
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
