<#
.SYNOPSIS
	The script is exporting all Active Directory subnets to a CSV file with the following datas:
		Name		: subnet name
		Location	: location attribute
		Type		: IPv4/IPv6
		IpStart		: first IP of the subnet
		IpEnd		: last IP of the subnet
		RangeStart	: first IP of the subnet (decimal format)
		RangeEnd	: last IP of the subnet (decimal format)
	
	The script is exporting the list of all subnet overlaps within the forest with the following datas:
		Subnet1		: the 'CN' attribute of the first subnet
		Site1		: the associated site of the first subnet
		Subnet2		: the 'CN' attribute of the second subnet
		Site2		: the associated site of the second subnet
		IsSameSite	: check if site1 and site2 are the same (boolean)
		
	The script is exporting the list of possible superscopes within a site with the following datas:
		Site		: concerned site
		Superscope	: IP and Mask of the superscope
		IpStart		: first IP of the superscope
		IpEnd		: last IP of the superscope
		Subnets		: list of concerned subnets
	
.NOTES
	Author		: Alexandre Augagneur (www.alexwinner.com)
	File Name	: Audit-ADSubnets.ps1
	
.EXAMPLE
	.\Audit-ADSubnets.ps1 -Path C:\Temp -CheckOverlap $true -CheckSuperscope $true
	
.PARAMETER Path
	Path for the exported files

.PARAMETER CheckOverlap
	If set to $true, the script is checking the existence of subnet overlaps

.PARAMETER CheckSuperScope
	If set to $true, the script is evaluating superscope creation within AD sites
#>

param (	
	[Parameter(Mandatory=$true)]
	[ValidateScript({Test-Path $_ -PathType Container})]
	[String] $Path,
	
	[Parameter()]
	[Boolean] $CheckOverlap = $false,
	
	[Parameter()] 
	[Boolean] $CheckSuperscope = $false
)

#Region Functions

####################################################
# Functions
####################################################

#---------------------------------------------------
# Converting an array of bytes to IPv6 format
#---------------------------------------------------
function Convert-BytesToIPv6 ( $arrayBytes )
{
	$String = $null
	$j = 0
	
	foreach ( $Item in $arrayBytes )
	{ 
		if ( $j -eq 2)
		{
			$String += ":"+[system.bitconverter]::Tostring($Item)
			$j = 1
		}
		else
		{ 
			$String += [system.bitconverter]::Tostring($Item)
			$j++
		}
	}
	Return $String
}

#---------------------------------------------------
# Create the IPv6 object
#---------------------------------------------------
function Compute-IPv6 ( $Obj, $ObjInputAddress, $Prefix )
{
	try
	{
		# Compute IP length
    	[int] $IntIPLength = 128 - $Prefix
		
		$Obj | Add-Member -type NoteProperty -name Type -value "IPv6"
		
		# Compute available IPs based on the IPv6 Prefix
		[System.Numerics.BigInteger] $NumberOfIPs = [System.Math]::Pow(2, $IntIPLength)
		$Obj | Add-Member -type NoteProperty -name NumberOfIPs -value $NumberOfIPs
		
		# Convert InputAddress to decimal value
		$ArrBytesInputAddress = $ObjInputAddress.GetAddressBytes()
		[array]::Reverse($ArrBytesInputAddress)
		$BigIntRangeStart = [System.Numerics.BigInteger] $ArrBytesInputAddress
		$Obj | Add-Member -type NoteProperty -name RangeStart -value $BigIntRangeStart

		# Compute the lastest available IP in decimal value
		[System.Numerics.BigInteger] $BigIntRangeEnd = ($BigIntRangeStart + $NumberOfIPs) - 1
		$Obj | Add-Member -type NoteProperty -name RangeEnd -value $BigIntRangeEnd

		# Convert the decimal value of the range start to IPv6 format
		$Obj | Add-Member -type NoteProperty -name IpStart -value $ObjInputAddress.IPAddressToString

		# Convert the lastest available IP in IPv6 format
		$ArrBytesRangeEnd = $BigIntRangeEnd.ToByteArray()
		[array]::Reverse($ArrBytesRangeEnd)
		$IpEnd = Convert-BytesToIpv6 $ArrBytesRangeEnd
		$Obj | Add-Member -type NoteProperty -name IpEnd -value $IpEnd
		
		return $Obj
	}
	catch
	{
		Write-Host $_.Exception.Message
	}
}

#---------------------------------------------------
# Create the IPv4 object
#---------------------------------------------------
function Compute-IPv4 ( $Obj, $ObjInputAddress, $Prefix )
{
	$Obj | Add-Member -type NoteProperty -name Type -value "IPv4"
	
	# Compute IP length
    [int] $IntIPLength = 32 - $Prefix
	
	$NumberOfIPs = ([System.Math]::Pow(2, $IntIPLength)) -1
	$ArrBytesInputAddress = $ObjInputAddress.GetAddressBytes()
	
	[Array]::Reverse($ArrBytesInputAddress)
	$IpStart = ([System.Net.IPAddress]($ArrBytesInputAddress -join ".")).Address

	If (($IpStart.Gettype()).Name -ine "double")
	{
		$IpStart = [Convert]::ToDouble($IpStart)
	}

	$IpStart = [System.Net.IPAddress] $IpStart
	$Obj | Add-Member -type NoteProperty -name IpStart -value $IpStart

	$ArrBytesIpStart = $IpStart.GetAddressBytes()
	[array]::Reverse($ArrBytesIpStart)
	$RangeStart = [system.bitconverter]::ToUInt32($ArrBytesIpStart,0)
	

	$IpEnd = $RangeStart + $NumberOfIPs

	If (($IpEnd.Gettype()).Name -ine "double")
	{
		$IpEnd = [Convert]::ToDouble($IpEnd)
	}

	$IpEnd = [System.Net.IPAddress] $IpEnd
	$Obj | Add-Member -type NoteProperty -name IpEnd -value $IpEnd

	$Obj | Add-Member -type NoteProperty -name RangeStart -value $RangeStart
	
	$ArrBytesIpEnd = $IpEnd.GetAddressBytes()
	[array]::Reverse($ArrBytesIpEnd)
	$Obj | Add-Member -type NoteProperty -name RangeEnd -value ([system.bitconverter]::ToUInt32($ArrBytesIpEnd,0))
	
	Return $Obj
}

#---------------------------------------------------
# Compute the subnet/prefix length
#---------------------------------------------------
function Compute-Prefix ( $IntStart, $IntEnd, $SubnetType )
{
	if ( $SubnetType -eq "IPv4" )
	{
		[Double] $NumberOfIPs = [Double] $IntEnd - [Double] $IntStart
	}
	else
	{
		[System.Numerics.BigInteger] $NumberOfIPs = [System.Numerics.BigInteger] $IntEnd - [System.Numerics.BigInteger] $IntStart
	}
	
	$IPlength = [Math]::Ceiling([Math]::Log($NumberOfIPs,2))
	$Prefix = 32 - $IPlength
	Return $Prefix
}

#EndRegion

$Message = 
@"

---------------------------------------------------------------------------------------------------
The script will not treat IPv6 subnets.

The script needs to use the CLR .Net Framework 4.0 to manage IPv6 subnets.

If you have already installed it, you have to indicate to PowerShell to use it.

Commands to run: 
	reg add hklm\software\microsoft\.netframework /v OnlyUseLatestCLR /t REG_DWORD /d 1 
	reg add hklm\software\wow6432node\microsoft\.netframework /v OnlyUseLatestCLR /t REG_DWORD /d 1
---------------------------------------------------------------------------------------------------

"@ 

# Check the CLR version (only version 4 can be used) 
if ( ([environment]::Version).Major -ne 4 ) 
{
    Write-Host $Message -Foreground 'Yellow'
	$script:CLRVersion4 = $false
} 
else
{ 
    try 
    { 
        # Add dll System.Numerics to use the BigInteger structure 
        Add-Type -Path "C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Numerics.dll"
		$script:CLRVersion4 = $true
	}
	catch
	{
		Write-Error "Adding DLL Sytem.Numerics: $($_.Exception.Message)"
		$script:CLRVersion4 = $false
	}
}

Write-Host "Retrieving AD subnets..."

# Connect to Active Directory and retrieve subnet objects
$objRootDSE = [System.DirectoryServices.DirectoryEntry] "LDAP://rootDSE"

$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://cn=subnets,cn=sites,"+$objRootDSE.ConfigurationNamingContext)
$Searcher.PageSize = 10000
$Searcher.SearchScope = "Subtree"
$Searcher.Filter = "(objectClass=subnet)"

$Properties = @("cn","location","siteobject")
$Searcher.PropertiesToLoad.AddRange(@($Properties))
$Subnets = $Searcher.FindAll()

$selectedProperties = $Properties | ForEach-Object {@{name="$_";expression=$ExecutionContext.InvokeCommand.NewScriptBlock("`$_['$_']")}}
[Regex] $RegexCN = "CN=(.*?),.*"
$SubnetsArray = @()

foreach ( $Subnet in $Subnets )
{
	# Construct the subnet object
	$SubnetObj = New-Object -TypeName PsObject
	$SubnetObj | Add-Member -type NoteProperty -name Name -value ([string] $Subnet.Properties['cn'])
	$SubnetObj | Add-Member -type NoteProperty -name Location -value ([string] $Subnet.Properties['location'])
	$SubnetObj | Add-Member -type NoteProperty -name Site -value ([string] $RegexCN.Match( $Subnet.Properties['siteobject']).Groups[1].Value)
     
	$InputAddress = (($SubnetObj.Name).Split("/"))[0]
	$Prefix = (($SubnetObj.Name).Split("/"))[1]
	
	# Construct System.Net.IPAddress 
    $ObjInputAddress = [System.Net.IPAddress] $InputAddress
	
	# Check if IP is a IPv4 or IPv6
	if ( ($ObjInputAddress.AddressFamily -match "InterNetworkV6") )
	{
		if ($script:CLRVersion4)
    	{
        	# Compute network address and IP ranges
        	$SubnetObj = Compute-IPv6 $SubnetObj $ObjInputAddress $Prefix
			$SubnetsArray += $SubnetObj
		}
    } 
    elseif ( $ObjInputAddress.AddressFamily -match "InterNetwork" )
    {
		$SubnetObj = Compute-IPv4 $SubnetObj $ObjInputAddress $Prefix
		$SubnetsArray += $SubnetObj
    }
}

$SubnetsArray | Export-Csv "$($Path)\ADSubnets-Export.csv" -Delimiter ";" -NoTypeInformation -Force
Write-Host "List of AD subnets exported to file : $($Path)\ADSubnets-Export.csv" -ForegroundColor Green

$ReportObj = New-Object -TypeName PsObject
$ReportObj | Add-Member -type NoteProperty -name 'Number of subnets' -value ($SubnetsArray.Count)
$ReportObj | Add-Member -type NoteProperty -name 'Number of subnets with no site' -value (($SubnetsArray | Where-Object { [string]::IsNullOrEmpty($_.Site) }).Count)

# Check if overlaps are existing between every subnets of the forest
if ( $CheckOverlap )
{
	Write-Host "Checking subnets overlap..."
	
	# Working of the existing array
	$Subnets = $SubnetsArray | Sort-Object -Property RangeStart
	$OverlapsArray = @()

	# Compare a subnet against all the others to check if it is overlapping one of them
	for ( $i=0; $i -lt $Subnets.Count; $i++ )
	{
		foreach ( $Item in $Subnets )
		{
			# Compare subnets ranges (decimal values of first IP and last IP) of the same type (IPv4/IPv6)
			if (($Item.Type -match $Subnets[$i].Type) -and ($Item.rangeStart -ge $Subnets[$i].rangeStart) -and ($Item.rangeEnd -le $Subnets[$i].rangeEnd) -and ($Item.Name -notmatch $Subnets[$i].Name) )
			{
				$OverlapObj = New-Object -TypeName PsObject
				$OverlapObj | Add-Member -type NoteProperty -name Subnet1 -value $Subnets[$i].Name
				$OverlapObj | Add-Member -type NoteProperty -name Site1 -value $Subnets[$i].Site
				$OverlapObj | Add-Member -type NoteProperty -name Subnet2 -value $Item.Name
				$OverlapObj | Add-Member -type NoteProperty -name Site2 -value $Item.Site
				
				if ( $OverlapObj.Site1 -eq $OverlapObj.Site2 )
				{
					$OverlapObj | Add-Member -type NoteProperty -name IsSameSite -value $true
				}
				else
				{
					$OverlapObj | Add-Member -type NoteProperty -name IsSameSite -value $false
				}
				
				$OverlapsArray += $OverlapObj
			}
		}
	}

	$OverlapsArray | Export-Csv "$($Path)\ADSubnets-Overlaps.csv" -Delimiter ";" -NoTypeInformation -Force
	Write-Host "List of overlapped AD subnets exported to file : $($Path)\ADSubnets-Overlaps.csv" -ForegroundColor Green
	
	$ReportObj | Add-Member -type NoteProperty -name 'Number of overlaps' -value ($OverlapsArray.Count)
}

# Check superscope creation per site. The script doesn't check the "location" attribute
if ( $CheckSuperscope )
{
	Write-Host "Checking subnets superscope..."
	
	$Subnets = $SubnetsArray | Sort-Object -Property Site,RangeStart
	$Sites = $Subnets | select Site -Unique
	$ArraySuperScopes = @()
	
	foreach ( $Site in $Sites )
	{	
		# Treatment of IPv6 if CLR 4.0 is used
		if ( $script:CLRVersion4 )
		{
			$SubnetsOfSite = $Subnets | Where-Object { $_.Site -eq $Site.Site } | sort RangeStart
		}
		else
		{
			$SubnetsOfSite = $Subnets | Where-Object { ($_.Site -eq $Site.Site) -and ($_.Type -eq "IPv4") } | sort RangeStart
		}
		
		# Check if there is more than 1 subnet associated to a site
		if ( $SubnetsOfSite.Count -gt 1 )
		{	
			# Treatment of each subnet in a site
			for ( $i=0; $i -lt ($SubnetsOfSite.Count-1) ; $i++ )
			{
				# Cast the last IP of the current subnet and the last IP of the next subnet
				if ( $SubnetsOfSite[$i].Type -eq "IPv4" )
				{
					[double] $LastIP = $SubnetsOfSite[$i].RangeEnd
					$LastIP++
					[double] $FirstIP = $SubnetsOfSite[$i+1].RangeStart
				}
				else
				{
					[System.Numerics.BigInteger] $LastIP = $SubnetsOfSite[$i].RangeEnd
					$LastIP += 1
					[System.Numerics.BigInteger] $FirstIP = $SubnetsOfSite[$i+1].RangeStart
				}
				
				# Check if we can merge the current subnet with the next subnet
				if ( ($LastIP -ge $FirstIP) -and ($SubnetsOfSite[$i].Type -eq $SubnetsOfSite[$i+1].Type) )
				{
					$SubnetType = $SubnetsOfSite[$i].Type
					
					# Check if we have to create a new superscope object
					if ( !($SuperScopeObj) )
					{
						$SuperScopeObj = New-Object -TypeName PsObject
						$SuperScopeObj | Add-Member -type NoteProperty -name Site -value $Site.Site
						
						$SuperScopeObj | Add-Member -type NoteProperty -name RangeStart -value ($SubnetsOfSite[$i].RangeStart)
						
						if ( $SubnetType -eq "IPv4" )
						{
							$SuperScopeObj | Add-Member -type NoteProperty -name IpStart -value ([System.Net.IPAddress] "$($SubnetsOfSite[$i].RangeStart)")
						}
						else
						{
							[System.Numerics.BigInteger] $BigIntIP = $SubnetsOfSite[$i].RangeStart
							$ArrBytesIP = $BigIntIP.ToByteArray()
							[array]::Reverse($ArrBytesIP)
							$IP = Convert-BytesToIpv6 $ArrBytesIP
							$SuperScopeObj | Add-Member -type NoteProperty -name IpStart -value ([System.Net.IPAddress] $IP)
						}
						
						$arrSubnets = @()
					}
					
					$arrSubnets += $SubnetsOfSite[$i].Name
				
					$SuperScopeObj | Add-Member -type NoteProperty -name RangeEnd -value ($SubnetsOfSite[$i+1].RangeEnd) -Force
					
					if ( $SubnetType -eq "IPv4" )
					{
						$SuperScopeObj | Add-Member -type NoteProperty -name IpEnd -value ([System.Net.IPAddress] "$($SubnetsOfSite[$i+1].RangeEnd)") -Force
					}
					else
					{
						[System.Numerics.BigInteger] $BigIntIP = $SubnetsOfSite[$i+1].RangeEnd
						$ArrBytesIP = $BigIntIP.ToByteArray()
						[array]::Reverse($ArrBytesIP)
						$IP = Convert-BytesToIpv6 $ArrBytesIP
						$SuperScopeObj | Add-Member -type NoteProperty -name IpEnd -value ([System.Net.IPAddress] $IP) -Force
					}
				}
				# Current subnet can not be merge with the next one
				else
				{
					# If property RangeEnd is not null then add the superscope to the superscopes array
					if ( $SuperScopeObj.RangeEnd )
					{
						$arrSubnets += $SubnetsOfSite[$i].Name
						$SuperScopeObj | Add-Member -type NoteProperty -name Subnets -value ([string] $arrSubnets)
						$SubnetLength = Compute-Prefix $SuperScopeObj.RangeStart $SuperScopeObj.RangeEnd $SubnetType
						$SuperScopeObj | Add-Member -type NoteProperty -name Superscope -value "$($SuperScopeObj.IpStart)/$Subnetlength"
						$ArraySuperScopes += $SuperScopeObj
						Remove-Variable -Name SuperScopeObj
					}
				}
			}
			# Special treatment for the lastest subnet which is part of a superscope
			if ( $SuperScopeObj.RangeEnd )
			{
				$arrSubnets += $SubnetsOfSite[$i].Name
				$SuperScopeObj | Add-Member -type NoteProperty -name Subnets -value ([string] $arrSubnets)
				$SubnetLength = Compute-Prefix $SuperScopeObj.RangeStart $SuperScopeObj.RangeEnd $SubnetType
				$SuperScopeObj | Add-Member -type NoteProperty -name Superscope -value "$($SuperScopeObj.IpStart)/$Subnetlength"
				$ArraySuperScopes  += $SuperScopeObj
				Remove-Variable -Name SuperScopeObj
			}
		}
	}

	$ArraySuperScopes | Select-Object Site,Superscope,IpStart,IpEnd,Subnets | Export-Csv "$($Path)\ADSubnets-Superscopes.csv" -Delimiter ";" -NoTypeInformation
	Write-Host "AD subnet superscopes evaluation exported to file : $($Path)\ADSubnets-Superscopes.csv" -ForegroundColor Green
	$ReportObj | Add-Member -type NoteProperty -name 'Number of superscopes' -value ($ArraySuperScopes.Count)
}

$ReportObj | fl