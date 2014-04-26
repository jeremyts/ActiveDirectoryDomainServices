<#
  This script will create the Time Server GPOs and WMI Filters for the Domain Controllers
  to ensure your time server hierarchy remains correct for transfer and seizure of the PDC(e)
  emulator FSMO role holder. The policies will apply on the next policy refresh or by forcing
  a group policy refresh.

  Script Name: CreateTimeServerGPOs.ps1
  Release 1.1
  Written by Jeremy@jhouseconsulting.com 14/01/2014

  Original script was written by Carl Webster:
  - Carl Webster, CTP and independent consultant
  - webster@carlwebster.com
  - @carlwebster on Twitter
  - http://www.CarlWebster.com
  - It can be found here:
    http://carlwebster.com/creating-a-group-policy-using-microsoft-powershell-to-configure-the-authoritative-time-server/

  WMI Filters are created via the New-ADObject cmdlet in the Active Directory module, which
  makes them of type "Microsoft.ActiveDirectory.Management.ADObject". However, the Group
  Policy module requires that you use an object of type "Microsoft.GroupPolicy.WmiFilter"
  when adding a wmifilter using the New-GPO cmdlet. Therefore there is no default way to use
  the Group Policy PowerShell cmdlets to add WMI Filters to GPOs without a bit or trickery.
  As Carl documented there is a "Group Policy WMI filter cmdlet module" available for download
  from here: http://gallery.technet.microsoft.com/scriptcenter/Group-Policy-WMI-filter-38a188f3
  But if you reverse engineer the code Bin Yi from Microsoft created, you'll see that he has
  simply and cleverly converted a "Microsoft.ActiveDirectory.Management.ADObject" object type
  to a "Microsoft.GroupPolicy.WmiFilter" object type. I didn't want to include the whole module
  for the simple task I needed, so have directly used the ConvertTo-WmiFilter function from the
  GPWmiFilter.psm1 module and tweaked it for my requirements. Many thanks to Bin.

  If your Active Directory is based on Windows 2003 or has been upgraded from Windows 2003, you
  may may have an issue with System Owned Objects. If you receive an error message along the
  lines of "The attribute cannot be modified because it is owned by the system", you'll need to
  set the following registry value:
    Key: HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\services\NTDS\Parameters
    Type: REG_DWORD
    Value: Allow System Only Change
    Data: 1

  Disable the Hyper-V time synchronization integration service:
  - The time source of "VM IC Time Synchronization Provider" (vmictimeprovider.dll) is enabled
    on Virtual Machines as part of the Hyper-V Integration Services. The following articles
    explain it in more depth and how it should be configured:
    - Time Sync Recommendations For Virtual DCs On Hyper-V – Change In Recommendations (AGAIN)
      http://jorgequestforknowledge.wordpress.com/2013/11/17/time-sync-recommendations-for-virtual-dcs-on-hyper-v-change-in-recommendations-again/
    - Time Synchronization in Hyper-V:
      http://blogs.msdn.com/b/virtual_pc_guy/archive/2010/11/19/time-synchronization-in-hyper-v.aspx
    - Hyper V Time Synchronization on a Windows Based Network:
      http://kevingreeneitblog.blogspot.com.au/2011/01/hyper-v-time-synchronization-on-windows.html

#>
Set-StrictMode -Version 2

#-------------------------------------------------------------
# Define variables specific to your Active Directory environment

# Set this to the NTP Servers the PDCe will sync with
$TimeServers = "0.au.pool.ntp.org,0x8 1.au.pool.ntp.org,0x8 2.au.pool.ntp.org,0x8 3.au.pool.ntp.org,0x8"

# This is the name of the GPO for the PDCe policy
$PDCeGPOName = "Set PDCe Domain Controller as Authoritative Time Server v1.0"

# This is the WMI Filter for the PDCe Domain Controller
$PDCeWMIFilter = @("PDCe Domain Controller",
                   "Queries for the domain controller that holds the PDC emulator FSMO role",
                   "root\CIMv2",
                   "Select * from Win32_ComputerSystem where DomainRole=5")

# This is the name of the GPO for the non-PDCe policy
$NonPDCeGPOName = "Set Time Settings on non-PDCe Domain Controllers v1.0"

# This is the WMI Filter for the non-PDCe Domain Controllers
$NonPDCeWMIFilter = @("Non-PDCe Domain Controllers",
                      "Queries for all domain controllers except for the one that holds the PDC emulator FSMO role",
                      "root\CIMv2",
                      "Select * from Win32_ComputerSystem where DomainRole=4")

# Set this to True to include the registry value to disable the Hyper-V Time Synchronization
$DisableHyperVTimeSynchronization = $True

# Set this to True if you need to set the "Allow System Only Change" value.
$AllowSystemOnlyChange = $False

# Set this to the number of seconds you would like to wait for Active Directory replication
# to complete before retrying to add the WMI filter to the Group Policy Object (GPO).
$SleepTimer = 10

#-------------------------------------------------------------
# Import the Active Directory Module
Import-Module ActiveDirectory -WarningAction SilentlyContinue
if ($Error.Count -eq 0) {
  #Write-Host "Successfully loaded Active Directory Powershell's module`n" -ForeGroundColor Green
} else {
  Write-Host "Error while loading Active Directory Powershell's module : $Error`n" -ForeGroundColor Red
  exit
}

# Import the Group Policy Module
Import-Module GroupPolicy -WarningAction SilentlyContinue
if ($Error.Count -eq 0) {
  #Write-Host "Successfully loaded Group Policy Powershell's module`n" -ForeGroundColor Green
} else {
  Write-Host "Error while loading Group Policy Powershell's module : $Error`n" -ForeGroundColor Red
  exit
}

#-------------------------------------------------------------

$defaultNC = ([ADSI]"LDAP://RootDSE").defaultNamingContext.Value
$TargetOU = "OU=Domain Controllers," + $defaultNC

function ConvertTo-WmiFilter([Microsoft.ActiveDirectory.Management.ADObject[]] $ADObject)
{
  # The concept of this function has been taken directly from the GPWmiFilter.psm1 module
  # written by Bin Yi from Microsoft. I have modified it to allow for the challenges of
  # Active Directory replication. It will return the WMI filter as an object of type
  # "Microsoft.GroupPolicy.WmiFilter".
  $gpDomain = New-Object -Type Microsoft.GroupPolicy.GPDomain
  $ADObject | ForEach-Object {
    $path = 'MSFT_SomFilter.Domain="' + $gpDomain.DomainName + '",ID="' + $_.Name + '"'
    $filter = $NULL
    try
      {
        $filter = $gpDomain.GetWmiFilter($path)
      }
    catch
      {
        write-host -ForeGroundColor Red "The WMI filter could not be found."
      }
    if ($filter)
      {
        [Guid]$Guid = $_.Name.Substring(1, $_.Name.Length - 2)
        $filter | Add-Member -MemberType NoteProperty -Name Guid -Value $Guid -PassThru | Add-Member -MemberType NoteProperty -Name Content -Value $_."msWMI-Parm2" -PassThru
      } else {
        write-host -ForeGroundColor Yellow "Waiting $SleepTimer seconds for Active Directory replication to complete."
        start-sleep -s $SleepTimer
        write-host -ForeGroundColor Yellow "Trying again to retrieve the WMI filter."
        ConvertTo-WmiFilter $ADObject
      }
  }
}

Function Create-Policy {
  param($GPOName,$NtpServer,$AnnounceFlags,$Type,$WMIFilter)

  If ($AllowSystemOnlyChange) {
    new-itemproperty "HKLM:\System\CurrentControlSet\Services\NTDS\Parameters" `
      -name "Allow System Only Change" -value 1 -propertyType dword -EA 0
  }

  $UseAdministrator = $False
  If ($UseAdministrator -eq $False) {
    $msWMIAuthor = (Get-ADUser $env:USERNAME).Name
  } Else {
    $msWMIAuthor = "Administrator@" + [System.DirectoryServices.ActiveDirectory.Domain]::getcurrentdomain().name
  }

  # Create WMI Filter
  $WMIGUID = [string]"{"+([System.Guid]::NewGuid())+"}"
  $WMIDN = "CN="+$WMIGUID+",CN=SOM,CN=WMIPolicy,CN=System,"+$defaultNC
  $WMICN = $WMIGUID
  $WMIdistinguishedname = $WMIDN
  $WMIID = $WMIGUID
 
  $now = (Get-Date).ToUniversalTime()
  $msWMICreationDate = ($now.Year).ToString("0000") + ($now.Month).ToString("00") + ($now.Day).ToString("00") + ($now.Hour).ToString("00") + ($now.Minute).ToString("00") + ($now.Second).ToString("00") + "." + ($now.Millisecond * 1000).ToString("000000") + "-000" 
  $msWMIName = $WMIFilter[0]
  $msWMIParm1 = $WMIFilter[1] + " "
  $msWMIParm2 = "1;3;" + $WMIFilter[2].Length.ToString() + ";" + $WMIFilter[3].Length.ToString() + ";WQL;" + $WMIFilter[2] + ";" + $WMIFilter[3] + ";"

  # msWMI-Name: The friendly name of the WMI filter
  # msWMI-Parm1: The description of the WMI filter
  # msWMI-Parm2: The query and other related data of the WMI filter
  $Attr = @{"msWMI-Name" = $msWMIName;"msWMI-Parm1" = $msWMIParm1;"msWMI-Parm2" = $msWMIParm2;"msWMI-Author" = $msWMIAuthor;"msWMI-ID"=$WMIID;"instanceType" = 4;"showInAdvancedViewOnly" = "TRUE";"distinguishedname" = $WMIdistinguishedname;"msWMI-ChangeDate" = $msWMICreationDate; "msWMI-CreationDate" = $msWMICreationDate} 
  $WMIPath = ("CN=SOM,CN=WMIPolicy,CN=System,"+$defaultNC) 

  $ExistingWMIFilters = Get-ADObject -Filter 'objectClass -eq "msWMI-Som"' -Properties "msWMI-Name","msWMI-Parm1","msWMI-Parm2"
  $array = @()

  If ($ExistingWMIFilters -ne $NULL) {
    foreach ($ExistingWMIFilter in $ExistingWMIFilters) {
      $array += $ExistingWMIFilter."msWMI-Name"
    }
  } Else {
    $array += "no filters"
  }

  if ($array -notcontains $msWMIName) {
    write-host -ForeGroundColor Green "Creating the $msWMIName WMI Filter..."
    $WMIFilterADObject = New-ADObject -name $WMICN -type "msWMI-Som" -Path $WMIPath -OtherAttributes $Attr
  } Else {
    write-host -ForeGroundColor Yellow "The $msWMIName WMI Filter already exists."
  }
  $WMIFilterADObject = $NULL

  # Get WMI filter
  $WMIFilterADObject = Get-ADObject -Filter 'objectClass -eq "msWMI-Som"' -Properties "msWMI-Name","msWMI-Parm1","msWMI-Parm2" | 
                Where {$_."msWMI-Name" -eq "$msWMIName"}

  $ExistingGPO = get-gpo $GPOName -ea "SilentlyContinue"   
  If ($ExistingGPO -eq $NULL) {            
    write-host -ForeGroundColor Green "Creating the $GPOName Group Policy Object..."

    # Create new GPO shell
    $GPO = New-GPO -Name $GPOName

    # Disable User Settings
    $GPO.GpoStatus = "UserSettingsDisabled"

    # Add the WMI Filter
    $GPO.WmiFilter = ConvertTo-WmiFilter $WMIFilterADObject

    # Set the three registry keys in the Preferences section of the new GPO
    Set-GPPrefRegistryValue -Name $GPOName -Action Update -Context Computer `
      -Key "HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Config" `
      -Type DWord -ValueName "AnnounceFlags" -Value $AnnounceFlags | out-null
 
    Set-GPPrefRegistryValue -Name $GPOName -Action Update -Context Computer `
      -Key "HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" `
      -Type String -ValueName "NtpServer" -Value "$NtpServer" | out-null
 
    Set-GPPrefRegistryValue -Name $GPOName -Action Update -Context Computer `
      -Key "HKLM\SYSTEM\CurrentControlSet\Services\W32Time\Parameters" `
      -Type String -ValueName "Type" -Value "$Type" | out-null

    If ($DisableHyperVTimeSynchronization) {
      # Disable the Hyper-V time synchronization integration service.
      Set-GPPrefRegistryValue -Name $GPOName -Action Update -Context Computer `
        -Key "HKLM\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\VMICTimeProvider" `
        -Type DWord -ValueName "Enabled" -Value 0 | out-null
    }

    # Link the new GPO to the Domain Controllers OU
    write-host -ForeGroundColor Green "Linking the $GPOName Group Policy Object to the $TargetOU OU..."
    New-GPLink -Name $GPOName `
      -Target "$TargetOU" | out-null
  } Else {
    write-host -ForeGroundColor Yellow "The $GPOName Group Policy Object already exists."
    write-host -ForeGroundColor Green "Adding the $msWMIName WMI Filter..."
    $ExistingGPO.WmiFilter = ConvertTo-WmiFilter $WMIFilterADObject
  }
  write-host -ForeGroundColor Green "Completed.`n"
  $ObjectExists = $NULL
}

Write-Host -ForeGroundColor Green "Creating the WMI Filters and Policies...`n"
Create-Policy "$PDCeGPOName" "$TimeServers" 5 "NTP" $PDCeWMIFilter
Create-Policy "$NonPDCeGPOName" "" 10 "NT5DS" $NonPDCeWMIFilter
