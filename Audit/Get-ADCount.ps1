<#
================================================================================
PURPOSE:	Count AD Files Size and Objects
AUTHOR:		Axel Limousin
VERSION:	1.0 
DATE:		03/01/2013
SYNTAX:		Get-ADCount <OutFormat> <DomainNC> <Scope>
EXAMPLE:	Get-ADCount Host
			Get-ADCount Host "DC=mydom,DC=com" Files
THANKS:		Technet
			Sean Metcalf, Bryan Sweeney, Alex Verboon, PScottC, jrv,
			Cédric Bigini, Olivier de Lagarde Montlezun, Freddy Elmaleh
COMMENTS:	Enhancements in next releases : 
			- OutputFormat: ConvertTo-HTML
			- Count: DHCP, DNS, Bridgeheads, RODC, GC, GPLinks, DFS,
			check GC for DomainObjects of all Domains in Forest
			- Values comparison to highlight potential issues
================================================================================
#>

Import-Module ActiveDirectory

$OutFormat = $args[0]
$DomainNC = $args[1]
$Scope = $args[2]

function Count-NTDS
{
	$Key = "SYSTEM\CurrentControlSet\Services\NTDS\Parameters"
	$ValueName = "DSA Database file"
	$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $env:COMPUTERNAME)
	$RegKey = $Reg.opensubkey($Key)
	$NTDSPath = $RegKey.getvalue($ValueName)
	
	$NTDSSize = (Get-Item $NTDSPath).length
	$NTDSSize = ($NTDSSize / 1MB)
	$NTDSSize = “{0:N2}” -f $NTDSSize + " MB"
	
	write-output $NTDSSize
}

function Count-SysVol
{
	$Key = "SYSTEM\CurrentControlSet\Services\Netlogon\Parameters"
	$ValueName = "SysVol"
	$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $env:COMPUTERNAME)
	$RegKey = $Reg.opensubkey($Key)
	$SysVolPath = $RegKey.getvalue($ValueName)
	$SysVolFolder = $SysVolPath.Remove($SysVolPath.Length -6,6) + "domain"
	
	$SysVolSize = Get-ChildItem $SysVolFolder -Recurse | Measure-Object -Property Length -Sum
	$SysVolSize = ($SysVolSize.sum / 1MB)
	$SysVolSize = “{0:N2}” -f $SysVolSize + " MB"
	
	write-output $SysVolSize
}

function Count-DomainPartitions
{
	$ConfigurationNC = (Get-ADRootDSE).ConfigurationNamingContext
	
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://CN=Partitions,$ConfigurationNC"
	$ADSearch.SearchScope = "OneLevel"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(&(objectClass=crossRef)(systemFlags=3))"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-Partitions
{
	$ConfigurationNC = (Get-ADRootDSE).ConfigurationNamingContext
	
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://CN=Partitions,$ConfigurationNC"
	$ADSearch.SearchScope = "OneLevel"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(objectClass=crossRef)"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-ApplicationPartitions
{
	$ConfigurationNC = (Get-ADRootDSE).ConfigurationNamingContext
	
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://CN=Partitions,$ConfigurationNC"
	$ADSearch.SearchScope = "OneLevel"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(&(objectClass=crossRef)(systemFlags=5))"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-Sites
{
	$ConfigurationNC = (Get-ADRootDSE).ConfigurationNamingContext
	
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://CN=Sites,$ConfigurationNC"
	$ADSearch.SearchScope = "OneLevel"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(objectClass=site)"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-Subnets
{
	$ConfigurationNC = (Get-ADRootDSE).ConfigurationNamingContext
	
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://CN=Subnets,CN=Sites,$ConfigurationNC"
	$ADSearch.SearchScope = "OneLevel"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(objectClass=subnet)"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-SiteLinks
{
	$ConfigurationNC = (Get-ADRootDSE).ConfigurationNamingContext
	
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://CN=Inter-Site Transports,CN=Sites,$ConfigurationNC"
	$ADSearch.SearchScope = "Subtree"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(objectClass=siteLink)"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-SiteLinkBridges
{
	$ConfigurationNC = (Get-ADRootDSE).ConfigurationNamingContext
	
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://CN=Inter-Site Transports,CN=Sites,$ConfigurationNC"
	$ADSearch.SearchScope = "Subtree"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(objectClass=siteLinkBridge)"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-Objects
{   
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://$DomainNC"
	$ADSearch.SearchScope = "Subtree"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(&distinguishedName=*)"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-Computers
{   
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://$DomainNC"
	$ADSearch.SearchScope = "Subtree"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(&(objectCategory=computer)(objectClass=user)(!servicePrincipalName=MSClusterVirtualServer*))"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-DomainControllers
{   
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://$DomainNC"
	$ADSearch.SearchScope = "Subtree"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(&(objectCategory=computer)(|(userAccountControl:1.2.840.113556.1.4.803:=8192)(userAccountControl:1.2.840.113556.1.4.803:=67108864)))"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-WorkstationsAndServers
{   
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://$DomainNC"
	$ADSearch.SearchScope = "Subtree"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(&(objectCategory=computer)(userAccountControl:1.2.840.113556.1.4.803:=4096)(!servicePrincipalName=MSClusterVirtualServer*))"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count

	Write-Output $Count
}

function Count-Users
{   
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://$DomainNC"
	$ADSearch.SearchScope = "Subtree"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(&(objectCategory=person)(objectClass=user))"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-InetOrgPersons
{   
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://$DomainNC"
	$ADSearch.SearchScope = "Subtree"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(objectClass=inetOrgPerson)"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-GlobalGroups
{   
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://$DomainNC"
	$ADSearch.SearchScope = "Subtree"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(&(objectCategory=group)(groupType:1.2.840.113556.1.4.803:=2147483650))"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-DomainLocalGroups
{   
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://$DomainNC"
	$ADSearch.SearchScope = "Subtree"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(&(objectCategory=group)(groupType:1.2.840.113556.1.4.803:=2147483652))"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-UniversalGroups
{   
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://$DomainNC"
	$ADSearch.SearchScope = "Subtree"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(&(objectCategory=group)(groupType:1.2.840.113556.1.4.803:=2147483656))"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-DistributionGroups
{   
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://$DomainNC"
	$ADSearch.SearchScope = "Subtree"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(&(objectCategory=group)(|(groupType:1.2.840.113556.1.4.803:=2)(groupType:1.2.840.113556.1.4.803:=4)(groupType:1.2.840.113556.1.4.803:=8))(!(groupType:1.2.840.113556.1.4.803:=2147483648)))"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-GroupPolicies
{   
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://CN=Policies,CN=System,$DomainNC"
	$ADSearch.SearchScope = "OneLevel"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(objectClass=groupPolicyContainer)"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-DomainTrusts
{   
	$ADSearch = New-Object DirectoryServices.DirectorySearcher
	$ADSearch.SearchRoot = "LDAP://CN=System,$DomainNC"
	$ADSearch.SearchScope = "Base"
        $ADSearch.PageSize = 1000
	$ADSearch.Filter = "(objectClass=trustedDomain)"
	$AllObjects = $ADSearch.FindAll()
	
	$Count = $AllObjects.Count
	
	Write-Output $Count
}

function Count-toHost
{	
	switch ($Scope) 
	{
		"Files"
		{
			$NTDS = Count-NTDS
			$SysVol = Count-SysVol
				
			Write-Host	"File ntds.dit: $NTDS",
						"Directory SysVol: $SysVol" -BackgroundColor "Black" -ForegroundColor "White" -Separator "`n"
		}
		"Forest"
		{
			$DF = Count-DomainPartitions
			$PF = Count-Partitions
			$APF = Count-ApplicationPartitions
			
			$SF = Count-Sites
			$SbF = Count-Subnets
			$SLF = Count-SiteLinks
			$SLBF = Count-SiteLinkBridges
			
			Write-Host	"Domains in Forest: $DF",
						"Partitions in Forest: $PF",
						"Application Partitions in Forest: $APF" -BackgroundColor "Black" -ForegroundColor "Red" -Separator "`n"
						
			Write-Host	"Sites in Forest: $SF",
						"Subnets in Forest: $SbF",
						"Site Links in Forest: $SLF",
						"Site Link Bridges in Forest: $SLBF" -BackgroundColor "Black" -ForegroundColor "Green" -Separator "`n"
		}
		"Domain"
		{
			$OD = Count-Objects $DomainNC
			
			$ACD = Count-Computers $DomainNC
			$DCD = Count-DomainControllers $DomainNC
			$AnDCD = Count-WorkstationsAndServers $DomainNC
			
			$AUD = Count-Users $DomainNC
			$LUD = Count-InetOrgPersons $DomainNC
			
			$SGGD = Count-GlobalGroups $DomainNC
			$SDLGD = Count-DomainLocalGroups $DomainNC
			$SUGD = Count-UniversalGroups $DomainNC
			$ADGD = Count-DistributionGroups $DomainNC
			
			$GPD = Count-GroupPolicies $DomainNC
			
			$TD = Count-DomainTrusts $DomainNC
			
			Write-Host	"Objects in Domain: $OD" -BackgroundColor "Black" -ForegroundColor "Gray"
			
			Write-Host	"All Computers in Domain: $ACD",
						"DCs in Domain: $DCD",
						"All nonDCs in Domain: $AnDCD" -BackgroundColor "Black" -ForegroundColor "Yellow" -Separator "`n"
			
			Write-Host	"AD-Users in Domain: $AUD",
						"LDAP-Users in Domain: $LUD" -BackgroundColor "Black" -ForegroundColor "Cyan" -Separator "`n"

			Write-Host	"Security Global Groups in Domain: $SGGD",
						"Security DLocal Groups in Domain: $SDLGD",
						"Security Universal Groups in Domain: $SUGD",
						"All Distribution Groups in Domain: $ADGD" -BackgroundColor "Black" -ForegroundColor "Blue" -Separator "`n"
			
			Write-Host	"Group Policies in Domain: $GPD" -BackgroundColor "Black" -ForegroundColor "Magenta"
			
			Write-Host	"Trusts in Domain: $TD" -BackgroundColor "Black" -ForegroundColor "DarkRed"
		}			
		"Computers"
		{
			$ACD = Count-Computers $DomainNC
			$DCD = Count-DomainControllers $DomainNC
			$AnDCD = Count-WorkstationsAndServers $DomainNC
			
			Write-Host	"All Computers in Domain: $ACD",
						"DCs in Domain: $DCD",
						"All nonDCs in Domain: $AnDCD" -BackgroundColor "Black" -ForegroundColor "Yellow" -Separator "`n"
		}			
		"Users"
		{
			$AUD = Count-Users $DomainNC
			$LUD = Count-InetOrgPersons $DomainNC
			
			Write-Host	"AD-Users in Domain: $AUD",
						"LDAP-Users in Domain: $LUD" -BackgroundColor "Black" -ForegroundColor "Cyan" -Separator "`n"
		}	
		"Groups"
		{
			$SGGD = Count-GlobalGroups $DomainNC
			$SDLGD = Count-DomainLocalGroups $DomainNC
			$SUGD = Count-UniversalGroups $DomainNC
			$ADGD = Count-DistributionGroups $DomainNC
			
			Write-Host	"Security Global Groups in Domain: $SGGD",
						"Security DLocal Groups in Domain: $SDLGD",
						"Security Universal Groups in Domain: $SUGD",
						"All Distribution Groups in Domain: $ADGD" -BackgroundColor "Black" -ForegroundColor "Blue" -Separator "`n"
		}	
		"GPOs"
		{
			$GPD = Count-GroupPolicies $DomainNC
			
			Write-Host	"Group Policies in Domain: $GPD" -BackgroundColor "Black" -ForegroundColor "Magenta" -Separator "`n"
		}
		"All"
		{
			$NTDS = Count-NTDS
			$SysVol = Count-SysVol
			
			$DF = Count-DomainPartitions
			$PF = Count-Partitions
			$APF = Count-ApplicationPartitions
			
			$SF = Count-Sites
			$SbF = Count-Subnets
			$SLF = Count-SiteLinks
			$SLBF = Count-SiteLinkBridges
			
			$OD = Count-Objects $DomainNC
			
			$ACD = Count-Computers $DomainNC
			$DCD = Count-DomainControllers $DomainNC
			$AnDCD = Count-WorkstationsAndServers $DomainNC
			
			$AUD = Count-Users $DomainNC
			$LUD = Count-InetOrgPersons $DomainNC
			
			$SGGD = Count-GlobalGroups $DomainNC
			$SDLGD = Count-DomainLocalGroups $DomainNC
			$SUGD = Count-UniversalGroups $DomainNC
			$ADGD = Count-DistributionGroups $DomainNC
			
			$GPD = Count-GroupPolicies $DomainNC
			
			$TD = Count-DomainTrusts $DomainNC
			
			Write-Host	"File ntds.dit: $NTDS",
						"Directory SysVol: $SysVol" -BackgroundColor "Black" -ForegroundColor "White" -Separator "`n"
										
			Write-Host	"Domains in Forest: $DF",
						"Partitions in Forest: $PF",
						"Application Partitions in Forest: $APF" -BackgroundColor "Black" -ForegroundColor "Red" -Separator "`n"

			Write-Host	"Sites in Forest: $SF",
						"Subnets in Forest: $SbF",
						"Site Links in Forest: $SLF",
						"Site Link Bridges in Forest: $SLBF" -BackgroundColor "Black" -ForegroundColor "Green" -Separator "`n"
			
			Write-Host	"Objects in Domain: $OD" -BackgroundColor "Black" -ForegroundColor "Gray"
			
			Write-Host	"All Computers in Domain: $ACD",
						"DCs in Domain: $DCD",
						"All nonDCs in Domain: $AnDCD" -BackgroundColor "Black" -ForegroundColor "Yellow" -Separator "`n"
			
			Write-Host	"AD-Users in Domain: $AUD",
						"LDAP-Users in Domain: $LUD" -BackgroundColor "Black" -ForegroundColor "Cyan" -Separator "`n"

			Write-Host	"Security Global Groups in Domain: $SGGD",
						"Security DLocal Groups in Domain: $SDLGD",
						"Security Universal Groups in Domain: $SUGD",
						"All Distribution Groups in Domain: $ADGD" -BackgroundColor "Black" -ForegroundColor "Blue" -Separator "`n"
			
			Write-Host	"Group Policies in Domain: $GPD" -BackgroundColor "Black" -ForegroundColor "Magenta"
			
			Write-Host	"Trusts in Domain: $TD" -BackgroundColor "Black" -ForegroundColor "DarkRed" 	
		}
		default
		{
			Write-Host	"Scope of counters is expected, please choose Files, Forest, Domain, Computers, Users, Groups, GPOs or All",
						"Syntax: Get-ADCount <OutFormat> <DomainNC> <Scope>",
						"Example: Get-ADCount Host ""DC=mydom,DC=com"" Files" -BackgroundColor "Black" -ForegroundColor "DarkGreen" -Separator "`n"
		}
	}	
}

function Count-toCsv
{	
	switch ($Scope) 
	{
		"Files"
		{
			$Count = New-Object psobject
			
			$Count | Add-Member NoteProperty "File ntds.dit" (Count-NTDS)
			$Count | Add-Member NoteProperty "Directory SysVol" (Count-SysVol)
			
			$Count | Export-Csv "$env:USERPROFILE\Desktop\Files.csv" -NoTypeInformation
		}
		"Forest"
		{
			$Count = New-Object psobject
			
			$Count | Add-Member NoteProperty "Domains in Forest" (Count-DomainPartitions)
			$Count | Add-Member NoteProperty "Partitions in Forest" (Count-Partitions)
			$Count | Add-Member NoteProperty "Application Partitions in Forest" (Count-ApplicationPartitions)
			
			$Count | Add-Member NoteProperty "Sites in Forest" (Count-Sites)
			$Count | Add-Member NoteProperty "Subnets in Forest" (Count-Subnets)
			$Count | Add-Member NoteProperty "Site Links in Forest" (Count-SiteLinks)
			$Count | Add-Member NoteProperty "Site Link Bridges in Forest" (Count-SiteLinkBridges)
			
			$Count | Export-Csv "$env:USERPROFILE\Desktop\Forest.csv" -NoTypeInformation
		}
		"Domain"
		{
			$Count = New-Object psobject
			
			$Count | Add-Member NoteProperty "Objects in Domain" (Count-Objects $DomainNC)
			
			$Count | Add-Member NoteProperty "All Computers in Domain" (Count-Computers $DomainNC)
			$Count | Add-Member NoteProperty "DCs in Domain" (Count-DomainControllers $DomainNC)
			$Count | Add-Member NoteProperty "All nonDCs in Domain" (Count-WorkstationsAndServers $DomainNC)
			
			$Count | Add-Member NoteProperty "AD-Users in Domain" (Count-Users $DomainNC)
			$Count | Add-Member NoteProperty "LDAP-Users in Domain" (Count-InetOrgPersons $DomainNC)
			
			$Count | Add-Member NoteProperty "Security Global Groups in Domain" (Count-GlobalGroups $DomainNC)
			$Count | Add-Member NoteProperty "Security DLocal Groups in Domain" (Count-DomainLocalGroups $DomainNC)
			$Count | Add-Member NoteProperty "Security Universal Groups in Domain" (Count-UniversalGroups $DomainNC)
			$Count | Add-Member NoteProperty "All Distribution Groups in Domain" (Count-DistributionGroups $DomainNC)
			
			$Count | Add-Member NoteProperty "Group Policies in Domain" (Count-GroupPolicies $DomainNC)
			
			$Count | Add-Member NoteProperty "Trusts in Domain" (Count-DomainTrusts $DomainNC)
			
			$Count | Export-Csv "$env:USERPROFILE\Desktop\Domain.csv" -NoTypeInformation
		}			
		"Computers"
		{
			$Count = New-Object psobject
			
			$Count | Add-Member NoteProperty "All Computers in Domain" (Count-Computers $DomainNC)
			$Count | Add-Member NoteProperty "DCs in Domain" (Count-DomainControllers $DomainNC)
			$Count | Add-Member NoteProperty "All nonDCs in Domain" (Count-WorkstationsAndServers $DomainNC)
			
			$Count | Export-Csv "$env:USERPROFILE\Desktop\Computers.csv" -NoTypeInformation
		}			
		"Users"
		{
			$Count = New-Object psobject
			
			$Count | Add-Member NoteProperty "AD-Users in Domain" (Count-Users $DomainNC)
			$Count | Add-Member NoteProperty "LDAP-Users in Domain" (Count-InetOrgPersons $DomainNC)
			
			$Count | Export-Csv "$env:USERPROFILE\Desktop\Users.csv" -NoTypeInformation
		}	
		"Groups"
		{
			$Count = New-Object psobject
			
			$Count | Add-Member NoteProperty "Security Global Groups in Domain" (Count-GlobalGroups $DomainNC)
			$Count | Add-Member NoteProperty "Security DLocal Groups in Domain" (Count-DomainLocalGroups $DomainNC)
			$Count | Add-Member NoteProperty "Security Universal Groups in Domain" (Count-UniversalGroups $DomainNC)
			$Count | Add-Member NoteProperty "All Distribution Groups in Domain" (Count-DistributionGroups $DomainNC)
			
			$Count | Export-Csv "$env:USERPROFILE\Desktop\Groups.csv" -NoTypeInformation
		}	
		"GPOs"
		{	
			$Count = New-Object psobject
			
			$Count | Add-Member NoteProperty "Group Policies in Domain" (Count-GroupPolicies $DomainNC)
			
			$Count | Export-Csv "$env:USERPROFILE\Desktop\GPOs.csv" -NoTypeInformation
		}
		"All"
		{
			$Count = New-Object psobject
			
			$Count | Add-Member NoteProperty "File ntds.dit" (Count-NTDS)
			$Count | Add-Member NoteProperty "Directory SysVol" (Count-SysVol)
						
			$Count | Add-Member NoteProperty "Domains in Forest" (Count-DomainPartitions)
			$Count | Add-Member NoteProperty "Partitions in Forest" (Count-Partitions)
			$Count | Add-Member NoteProperty "Application Partitions in Forest" (Count-ApplicationPartitions)
			
			$Count | Add-Member NoteProperty "Sites in Forest" (Count-Sites)
			$Count | Add-Member NoteProperty "Subnets in Forest" (Count-Subnets)
			$Count | Add-Member NoteProperty "Site Links in Forest" (Count-SiteLinks)
			$Count | Add-Member NoteProperty "Site Link Bridges in Forest" (Count-SiteLinkBridges)
						
			$Count | Add-Member NoteProperty "Objects in Domain" (Count-Objects $DomainNC)
			
			$Count | Add-Member NoteProperty "All Computers in Domain" (Count-Computers $DomainNC)
			$Count | Add-Member NoteProperty "DCs in Domain" (Count-DomainControllers $DomainNC)
			$Count | Add-Member NoteProperty "All nonDCs in Domain" (Count-WorkstationsAndServers $DomainNC)
			
			$Count | Add-Member NoteProperty "AD-Users in Domain" (Count-Users $DomainNC)
			$Count | Add-Member NoteProperty "LDAP-Users in Domain" (Count-InetOrgPersons $DomainNC)
			
			$Count | Add-Member NoteProperty "Security Global Groups in Domain" (Count-GlobalGroups $DomainNC)
			$Count | Add-Member NoteProperty "Security DLocal Groups in Domain" (Count-DomainLocalGroups $DomainNC)
			$Count | Add-Member NoteProperty "Security Universal Groups in Domain" (Count-UniversalGroups $DomainNC)
			$Count | Add-Member NoteProperty "All Distribution Groups in Domain" (Count-DistributionGroups $DomainNC)
			
			$Count | Add-Member NoteProperty "Group Policies in Domain" (Count-GroupPolicies $DomainNC)
			
			$Count | Add-Member NoteProperty "Trusts in Domain" (Count-DomainTrusts $DomainNC)

			$Count | Export-Csv "$env:USERPROFILE\Desktop\All.csv" -NoTypeInformation
		}
		default
		{
			Write-Host	"Scope of counters is expected, please choose Files, Forest, Domain, Computers, Users, Groups, GPOs or All",
						"Syntax: Get-ADCount <OutFormat> <DomainNC> <Scope>",
						"Example: Get-ADCount Host ""DC=mydom,DC=com"" Files" -BackgroundColor "Black" -ForegroundColor "DarkGreen" -Separator "`n"
		}
	}
}

function CountAD
{
	switch ($OutFormat)
	{
		"Host"
		{
			if ($DomainNC)
			{
				Count-toHost $DomainNC $Scope
			}
			else
			{
				$DomainNC = ([adsi]'').distinguishedName
				$Scope = "Domain"
				Count-toHost $DomainNC $Scope
			}
		}
		"Csv"
		{
			if ($DomainNC)
			{
				Count-toCsv $DomainNC $Scope
			}
			else
			{
				$DomainNC = ([adsi]'').distinguishedName
				$Scope = "Domain"
				Count-toCsv $DomainNC $Scope
			}
		}
		default
		{
			Write-Host	"Output format expected, please choose Host or Csv",
						"Syntax: Get-ADCount <OutFormat> <DomainNC> <Scope>",
						"Example: Get-ADCount Host ""DC=mydom,DC=com"" Files" -BackgroundColor "Black" -ForegroundColor "DarkGreen" -Separator "`n"
		}
	}
}

CountAD $OutFormat $DomainNC $Scope