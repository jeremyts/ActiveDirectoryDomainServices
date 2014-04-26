<#
  This script will create a report of all computer accounts. It helps
  idenify all domain joined servers or workstations in your environment
  for licensing and reporting requirements, including any potentially
  stale Computer accounts that can be deleted.

  Note that for servers we filter out Cluster Name Objects (CNOs) and
  Virtual Computer Objects (VCOs) by checking the objects serviceprincipalname
  property for a value of MSClusterVirtualServer. The CNO is the cluster
  name, whereas a VCO is the client access point for the clustered role.
  These are not actual computers, so we exlude them to assist with
  accuracy.

  A cluster updates the lastLogonTimeStamp of the CNO/VNO when it brings
  a clustered network name resource online. So it could be running for
  months without an update to the lastLogonTimeStamp attribute.

  For servers it reports on whether or not it's a virtual machine, and
  which hypervisor it's running on.

  The CSV report has 18 columns, 7 of which are optional depending on
  the script variables you set.
  - ComputerDomain - Domain name of the computer. A handy column when auditing multiple domains and merging spread sheets. 
  - ComputerName – Needs no explanation.
  - OperatingSystem – The Operating System registered in Active Directory.
  - IsVirtual - Validates if the server is virtual.
  - Hypervisor - The hypervisor the virtual machine is running on.
  - AccountEnabled – If the AD account is enabled.
  - IsPingable – If it's pingable (AKA alive). Will be accurate unless a firewall or access list is blocking it.
  - PasswordLastChanged – When the computer password was last changed. Should typically never go beyond 90 days.
  - StaleAccount – is derived from 3 values:
                   1. PasswordLastChanged  > 90 days ago
                   2. LastLogonDate > 30 days ago
                   3. IsPingable = False
  - LastLogonDate - The lastLogonTimeStamp attribute from Active Directory, which can be up to 14 days out.
  - RealLastLogonDate – An accurate date and time when it last logged on (authenticated). This get the lastLogon attribute from each Domain Controller.
  - LastLogonDC – The Domain Controller it last authenticated against, which is determined in the process of collecting the lastLogon attribute.
  - CcmExecInstalled - Checks to see if the SMS Agent Host (SCCM) Service is installed.
  - CcmExecStatus - Checks the status of the SMS Agent Host (SCCM) Service.
  - HealthServiceInstalled - Checks to see if the Microsoft Monitoring Agent (SCOM) Service is installed.
  - HealthServiceStatus - Checks the status of the Microsoft Monitoring Agent (SCOM) Service.
  - FRSSaaSClientAgentInstalled - Checks to see if the FrontRange SaaS ClientAgent Service is installed.
  - FRSSaaSClientAgentStatus - Checks the status of the FrontRange SaaS ClientAgent Service.

  Note that the services will depend on the services you list in the
  $Services array.

  You may notice a question mark (?) in various attributes, such as the
  OperatingSystem string. Refer to Microsoft KB829856 for an explanation.

  To provide accurate data we determine a stale computer account using
  three different parameters.
  For Servers:
    1) Password last set more than 90 days ago
    2) Last logged in more than 30 days ago
    3) It's not pingable
  For Workstations:
    1) Password last set more than 90 days ago
    2) Last logged in more than 30 days ago

  Computers change their passwword if and when they feel like it. The
  domain doesn't initiate the change. It is controlled by three values
  under the following registry key:
  - HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters
    - DisablePasswordChange
    - MaximumPasswordAge
    - RefusePasswordChange
  If these values are not present, the default value of 30 days will be
  used.

  Script Name: Get-AllComputersAccounts.ps1
  Release 1.4
  Modified by Jeremy@jhouseconsulting.com 27th February 2014
  Written by Jeremy@jhouseconsulting.com 31st January 2014

  References:
  - http://blogs.technet.com/b/ken_brumfield/archive/2008/09/16/identifying-stale-user-and-computer-accounts.aspx
  - http://blogs.technet.com/b/askds/archive/2011/08/23/cluster-and-stale-computer-accounts.aspx
  - http://blogs.technet.com/b/askds/archive/2009/02/15/test2.aspx
  - http://blogs.metcorpconsulting.com/tech/?p=1369
  - http://blogs.msdn.com/b/clustering/archive/2011/08/17/10197069.aspx
  - http://gallery.technet.microsoft.com/scriptcenter/Get-Active-Directory-User-bbcdd771
  - http://gallery.technet.microsoft.com/scriptcenter/Determine-if-a-computer-is-cdd20473
#>

#-------------------------------------------------------------

# Set this to true to process workstations only. Setting it to
# false will process servers only.
$Workstations = $True

# Set this to true to get an accurate last logon time. Only set
# to true if needed.
$ExactLastLogon = $False

# Set this to true to validate if the server is virtual. This is
# only valid for servers.
$ValidateVirtualStatus = $True

# Set this to true to check for the existence of all services
# in the $Services array. This is only valid for servers.
$CheckForService = $True

# Set this to the names of the services to be checked. This is
# only valid for servers.
$Services = @("CcmExec","HealthService","FRSSaaSClientAgent")

# Set this value to true if you want to see the progress bar.
$ProgressBar = $True

#-------------------------------------------------------------

# Import the Active Directory Module
Import-Module ActiveDirectory -WarningAction SilentlyContinue
if ($Error.Count -eq 0) {
  #Write-Host "Successfully loaded Active Directory Powershell's module`n" -ForeGroundColor Green
} else {
  Write-Host "Error while loading Active Directory Powershell's module : $Error`n" -ForeGroundColor Red
  exit
}

#-------------------------------------------------------------

Function Get-RemoteServerVirtualStatus
{
    <#
    .SYNOPSIS
        Validate if a remote server is virtual or physical
    .DESCRIPTION
        Uses wmi (along with an optional credential) to determine if a remote computers, or list of remote computers are virtual.
        If found to be virtual, a best guess effort is done on which type of virtual platform it is running on.
    .PARAMETER ComputerName
        Computer or IP address of machine
    .PARAMETER PromptForCredential
        Set this if you want the function to prompt for alternate credentials.
    .PARAMETER Credential
        Provide an alternate credential
    .EXAMPLE
        $Credential = Get-Credential
        Get-RemoteServerVirtualStatus 'Server1','Server2' -Credential $Credential | select ComputerName,IsVirtual,VirtualType | ft
        
        Description:
        ------------------
        Using an alternate credential, determine if server1 and server2 are virtual. Return the results along with the type of virtual machine it might be.
    .EXAMPLE
        (Get-RemoteServerVirtualStatus server1).IsVirtual
        
        Description:
        ------------------
        Determine if server1 is virtual and returns either true or false.

    .LINK
        http://www.the-little-things.net/
    .LINK
        http://nl.linkedin.com/in/zloeber
    .NOTES
        
        Name       : Get-RemoteServerVirtualStatus
        Version    : 1.0.0 07/27/2013
                     - First release
        Author     : Zachary Loeber
        Disclaimer : This script is provided AS IS without warranty of any kind. I 
                     disclaim all implied warranties including, without limitation,
                     any implied warranties of merchantability or of fitness for a 
                     particular purpose. The entire risk arising out of the use or
                     performance of the sample scripts and documentation remains
                     with you. In no event shall I be liable for any damages 
                     whatsoever (including, without limitation, damages for loss of 
                     business profits, business interruption, loss of business 
                     information, or other pecuniary loss) arising out of the use of or 
                     inability to use the script or documentation. 

        Copyright  : I believe in sharing knowledge, so this script and its use is 
                     subject to : http://creativecommons.org/licenses/by-sa/3.0/
    #>
    [cmdletBinding(SupportsShouldProcess = $true)]
    param(
        [parameter( Position=0,
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    HelpMessage="Computer or IP address of machine to test")]
        [string[]]$ComputerName,
        [parameter( HelpMessage="Set this if you want the function to prompt for alternate credentials.")]
        [switch]$PromptForCredential,
        [parameter( HelpMessage="Pass an alternate credential")]
        [System.Management.Automation.PSCredential]$Credential = $null
    )
    BEGIN
    {
        if ($PromptForCredential)
        {
            $Credential = Get-Credential
        }
        $WMISplat = @{}
        if ($Credential -ne $null)
        {
            $WMISplat.Credential = $Credential
        }
    }
    PROCESS
    {
        $results = @()
        $computernames = @()
        $computernames += $ComputerName
        
        foreach($computer in $computernames)
        {
            $WMISplat.ComputerName = $computer
            try
            {
                $wmibios = Get-WmiObject Win32_BIOS @WMISplat -ErrorAction Stop | Select-Object version,serialnumber
                $wmisystem = Get-WmiObject Win32_ComputerSystem @WMISplat -ErrorAction Stop | Select-Object model,manufacturer
                $CanConnect = $true
            }
            catch
            {
                $CanConnect = $false
            }
            if ($CanConnect)
            {
                $ResultProps = @{ 
                    ComputerName = $computer
                    BIOSVersion = $wmibios.Version
                    SerialNumber = $wmibios.serialnumber
                    Manufacturer = $wmisystem.manufacturer
                    Model = $wmisystem.model
                    IsVirtual = $false
                    VirtualType = $null
                }
                if ($wmibios.Version -match "VIRTUAL") 
                {
                    $ResultProps.IsVirtual = $true
                    $ResultProps.VirtualType = "Hyper-V"
                }
                elseif ($wmibios.Version -match "A M I") 
                {
                    $ResultProps.IsVirtual = $true
                    $ResultProps.VirtualType = "Virtual PC"
                }
                elseif ($wmibios.Version -like "*Xen*") 
                {
                    $ResultProps.IsVirtual = $true
                    $ResultProps.VirtualType = "Xen"
                }
                elseif ($wmibios.SerialNumber -like "*VMware*")
                {
                    $ResultProps.IsVirtual = $true
                    $ResultProps.VirtualType = "VMWare"
                }
                elseif ($wmisystem.manufacturer -like "*Microsoft*")
                {
                    $ResultProps.IsVirtual = $true
                    $ResultProps.VirtualType = "Hyper-V"
                }
                elseif ($wmisystem.manufacturer -like "*VMWare*")
                {
                    $ResultProps.IsVirtual = $true
                    $ResultProps.VirtualType = "VMWare"
                }
                elseif ($wmisystem.model -like "*Virtual*")
                {
                    $ResultProps.IsVirtual = $true
                    $ResultProps.VirtualType = "Unknown Virtual Machine"
                }
                $results += New-Object PsObject -Property $ResultProps
            }
            else
            {
               #"Cannot connect via WMI to determine if it's virtual"
            }
        }
    }
    END
    {
        return $results
    }
}

#-------------------------------------------------------------

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
If ($Workstations -eq $False) {
  $ReferenceFile = $(&$ScriptPath) + "\AllServerAccountsReport.csv"
} else {
  $ReferenceFile = $(&$ScriptPath) + "\AllWorkstationAccountsReport.csv"
}

$DNSRoot = (Get-ADDomain).DNSRoot

If ($Workstations -eq $False) {
  # Get All Servers filtering out Cluster Name Objects (CNOs) and Virtual computer Objects (VCOs) 
  $Computers = Get-ADComputer -Filter * -Properties Name,Operatingsystem,passwordLastSet,LastLogonDate,servicePrincipalName | Where-Object {($_.Operatingsystem -like '*server*') -AND !($_.serviceprincipalname -like '*MSClusterVirtualServer*')} | Sort-Object Name
} Else {
  # Get All Workstations
  $Computers = Get-ADComputer -Filter * -Properties Name,Operatingsystem,passwordLastSet,LastLogonDate | Where-Object {($_.Operatingsystem -like '*windows*') -AND !($_.Operatingsystem -like '*server*')} | Sort-Object Name
}

# Get Domain Controllers
$DomainControllers = Get-ADDomainController -Filter * | Sort-Object Name

$array = @()
$TotalProcessed = 0
$Count = 0
$Count = $Computers.Count

write-Host -ForegroundColor Green "There are $Count computer objects to process.`n"

ForEach ($Computer in $Computers) {
  write-Host -ForegroundColor Green "Processing"($Computer.Name)

  # On the very rare occasion I've found that the DNSHostName attribute may not be set.
  # When that happens we assume that the $ComputerDomain equals the DNS domain name
  # from where the script is being run from.
  If ($Computer.DNSHostName -ne $NULL) {
    $ComputerDomain = $Computer.DNSHostName.Substring(($Computer.DNSHostName.Split(".")[0].length + 1))
  } Else {
    $ComputerDomain = $DNSRoot
  }
  write-Host -ForegroundColor Green "- Domain: $ComputerDomain"
  write-Host -ForegroundColor Green "- OperatingSystem: "($Computer.OperatingSystem)

  $output = New-Object PSObject
  $output | Add-Member NoteProperty -Name "ComputerDomain" $ComputerDomain
  $output | Add-Member NoteProperty -Name "ComputerName" $Computer.Name
  $output | Add-Member NoteProperty -Name "OperatingSystem" $Computer.OperatingSystem

  If ($Workstations -eq $False) {
    # Check to see if it's physical or virtual.
    If ($ValidateVirtualStatus) {
      $VirtualStatus = Get-RemoteServerVirtualStatus ($Computer.Name)
      $IsVirtual = $VirtualStatus.IsVirtual
      $VirtualType = $VirtualStatus.VirtualType
      If ($IsVirtual -eq $True) {
        write-Host -ForegroundColor Green "- This is a virtual machine."
        write-Host -ForegroundColor Green "- Hypervisor: $VirtualType"
      } elseIf ($IsVirtual -eq $False) {
        write-Host -ForegroundColor Green "- This is a physical machine."
        $VirtualType = "N/A"
      } else {
        $IsVirtual = "unable to determine"
        $VirtualType = "unable to determine"
        write-Host -ForegroundColor Red "- $IsVirtual if it's virtual."
      }
      $output | Add-Member NoteProperty -Name "IsVirtual" $IsVirtual
      $output | Add-Member NoteProperty -Name "Hypervisor" $VirtualType
    }
  }

  If ($Computer.Enabled) {
    write-Host -ForegroundColor Green "- Account Enabled: $($Computer.Enabled)"
  } else {
    write-Host -ForegroundColor Red "- Account Enabled: $($Computer.Enabled)"
  }
  $output | Add-Member NoteProperty -Name "AccountEnabled" $Computer.Enabled

  If ($Workstations -eq $False) {
    $IsPingable = $False
    if (Test-Connection -Cn $Computer.Name -BufferSize 16 -Count 1 -ea 0 -quiet) {
      $IsPingable = $True
      write-Host -ForegroundColor Green "- Can be pinged: $IsPingable"
    } else {
      write-Host -ForegroundColor Red "- Can be pinged: $IsPingable"
    }
    $output | Add-Member NoteProperty -Name "IsPingable" $IsPingable
  }

  write-Host -ForegroundColor Green "- Password last set: $($Computer.passwordLastSet)"
  $PasswordTooOld = $False
  If ($Computer.passwordLastSet -le (Get-Date).AddDays(-90)) {
    write-Host -ForegroundColor Red "  it was set more than 90 days ago."
    $PasswordTooOld = $True
  } elseIf ($Computer.passwordLastSet -le (Get-Date).AddDays(-60)) {
    write-Host -ForegroundColor yellow "  it was set more than 60 days ago."
  } elseIf ($Computer.passwordLastSet -le (Get-Date).AddDays(-30)) {
    write-Host -ForegroundColor yellow "  it was set more than 30 days ago."
  }
  $output | Add-Member NoteProperty -Name "PasswordLastChanged" $Computer.passwordLastSet

  write-Host -ForegroundColor Green "- Last logon time stamp: $($Computer.LastLogonDate)"
  $HasNotRecentlyLoggedOn = $False
  If ($Computer.LastLogonDate -le (Get-Date).AddDays(-30)) {
    write-Host -ForegroundColor Red "  it has not logged in for more than 30 days."
    $HasNotRecentlyLoggedOn = $True
  }
  $output | Add-Member NoteProperty -Name "LastLogonDate" $Computer.LastLogonDate

  If ($ExactLastLogon) {
    write-Host -ForegroundColor Green "- Finding a more accurate last logon time..."
    $RealLastLogonDate = "01/01/1601 08:00:00"
    ForEach ($DomainController in $DomainControllers) {
      if ($DomainController.Domain -eq $ComputerDomain) {
        if (Test-Connection -Cn $DomainController.Name -BufferSize 16 -Count 1 -ea 0 -quiet) {
          write-Host -ForegroundColor Green "  - Checking against $($DomainController.Name)"
          Try {
            $Lastlogon = Get-ADComputer -Identity $Computer.Name -Properties LastLogon -Server $DomainController.Name
            if ($RealLastLogonDate -le [DateTime]::FromFileTime($Lastlogon.LastLogon)) {
              $RealLastLogonDate = [DateTime]::FromFileTime($Lastlogon.LastLogon)
              $LastusedDC = $DomainController.Name
           }
          }
          Catch {
            # There was an error connecting to $DomainController.Name
          }
        }
      }
    }
    if ($RealLastLogonDate -match "1/01/1601") {
      $RealLastLogonDate = "never logged on before"
      $LastusedDC = ""
    }
    If ($RealLastLogonDate -ne "never logged on before") {
      write-Host -ForegroundColor Green "  - It last logged on to $LastusedDC on $RealLastLogonDate"
    } else {
      write-Host -ForegroundColor Yellow "  - It has $RealLastLogonDate"
    }
    $output | Add-Member NoteProperty -Name "RealLastLogonDate" $RealLastLogonDate
    $output | Add-Member NoteProperty -Name "LastLogonDC" $LastusedDC
  }

  # Check if it's a stale account.
  $IsStale = $False
  If ($Workstations -eq $False) {
    If ($PasswordTooOld -eq $True -AND $HasNotRecentlyLoggedOn -eq $True -AND $IsPingable -eq $False) {
      $IsStale = $True
    }
  }  else {
    If ($PasswordTooOld -eq $True -AND $HasNotRecentlyLoggedOn -eq $True) {
      $IsStale = $True
    }
  }
  If ($IsStale) {
    write-Host -ForegroundColor Red "- It's highly probable that this is a stale account."
  }
  $output | Add-Member NoteProperty -Name "StaleAccount" $IsStale

  If ($Workstations -eq $False) {
    If ($CheckForService) {
      # Create new variable for each service in the $Services array
      ForEach ($Service in $Services) {
        $tempvar1 = $service+"Installed"
        new-variable -name $tempvar1 -value $False -Force
        $tempvar2 = $service+"Status"
        new-variable -name $tempvar2 -value "N/A" -Force
        $output | Add-Member NoteProperty -Name $tempvar1 (get-variable -name $tempvar1 -valueonly)
        $output | Add-Member NoteProperty -Name $tempvar2 (get-variable -name $tempvar2 -valueonly)
      }
      Try {
        $serviceObj = Get-Service -computername ($Computer.Name) -include $Services | Select-Object Name, Status
        If ($serviceObj -ne $NULL) {
          $InstalledServices = @()
          ForEach ($obj in $serviceObj) {
            $InstalledServices += $obj.Name
            If ($obj.Status -eq "Running") {
              write-host -ForegroundColor green "- $($obj.Name) service found in a $($obj.Status) state."
            } Else {
              write-host -ForegroundColor red "- $($obj.Name) service found in a $($obj.Status) state."
            }
            $tempvar1 = $obj.Name+"Installed"
            $tempvar2 = $obj.Name+"Status"
            set-variable -name $tempvar1 -value $True
            set-variable -name $tempvar2 -value $obj.Status
            $output.PSObject.Properties.Remove($tempvar1)
            $output | Add-Member NoteProperty -Name $tempvar1 (get-variable -name $tempvar1 -valueonly)
            $output.PSObject.Properties.Remove($tempvar2)
            $output | Add-Member NoteProperty -Name $tempvar2 (get-variable -name $tempvar2 -valueonly)
          }
          ForEach ($Service in $Services) {
            If ($InstalledServices -notcontains $Service) {
              write-host -ForegroundColor red "- $Service service not installed."
            }
          }
        } Else {
          write-host -ForegroundColor red "- None of the services are installed."
        }
      }
      Catch {
        $ComputerError = "$true"
        $ErrorDescription = "Error connecting using the Get-Service cmdlet."
        write-Host -ForegroundColor Red "- $ErrorDescription"
      }
    }
  }

  $array += $output

  $TotalProcessed ++
  $percent = "{0:P}" -f ($TotalProcessed/$Count)
  write-host -ForegroundColor green "Processed $TotalProcessed of $Count computer accounts = $percent complete."
  Write-host " "
  If ($ProgressBar) {
    Write-Progress -Activity 'Processing Users' -Status ("Username: {0}" -f $($Computer.Name)) -PercentComplete (($TotalProcessed/$Count)*100)
  }
}

# Write-Output $array | Format-Table
$array | export-csv -notype -path "$ReferenceFile"

# Remove the quotes
(get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii
