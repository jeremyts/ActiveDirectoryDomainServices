# This script will list all the authorized DHCP servers in Active Directory.

# http://blogs.technet.com/b/heyscriptingguy/archive/2013/01/10/use-powershell-to-query-ad-ds-for-dhcp-servers.aspx

import-module ActiveDirectory

$defaultNC = ([ADSI]"LDAP://RootDSE").defaultNamingContext.Value

$configurationNC = "cn=configuration," + $defaultNC

Get-ADObject -SearchBase $configurationNC -Filter "objectclass -eq 'dhcpclass' -AND Name -ne 'dhcproot'"
