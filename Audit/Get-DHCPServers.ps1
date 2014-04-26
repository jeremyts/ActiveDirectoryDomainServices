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
        $names += $object.Name
      }
    }
  }
  $names
}

$Servers = Get-AuthorizedDHCPServers


#$ComputerName = "PTHADDS03.FMG.local"
#$ServiceName = "DHCPServer"
#$serviceObj = Get-Service -ComputerName $ComputerName | ?{ $_.ServiceName -eq $serviceName } | Select-Object Name, @{Name="Jeremy";Expression={$._Status}}, Status
#$serviceObj

#http://blogs.technet.com/b/heyscriptingguy/archive/2012/11/12/force-a-domain-wide-update-of-group-policy-with-powershell.aspx


$sessions = New-PSSession -ComputerName $Servers
$Sessions
$serviceObj = Invoke-Command -Session $sessions -ScriptBlock {
  $ServiceName = "DHCPServer"
  Get-Service | ?{ $_.ServiceName -eq $serviceName } | Select-Object -Property Name, @{Name="State";Expression={$_.Status}}
}

$serviceObj

Remove-PSSession $sessions
