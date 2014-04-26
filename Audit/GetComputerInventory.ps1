<#
This script will perform auditing on domain computers using WMI, ADSI and Remote Registry
It helps gather whatever information is required for reporting.

Release 1.0
Written by Jeremy@jhouseconsulting.com 14th August 2013

#>

#-------------------------------------------------------------
# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$OnlineFile = $(&$ScriptPath) + "\inventory.csv"
$OfflineFile = $(&$ScriptPath) + "\inventory-offline.csv"

$Windows2008R2SP1Hotfixes = @("KB979744","KB979744","KB983440","KB979099","KB982867","KB977020")

#-------------------------------------------------------------

$datetime = Get-Date -Format "yyyyMMddhhmmss"

# If used for SPLA reporting, you need to report on previous month, so we get the first and
# last day of last month.
$date = Get-Date
$StartOfPreviousMonth = Get-Date $date.AddMonths(-1) -day 1 -hour 0 -minute 0 -second 0
$EndOfPreviousMonth = (($StartOfPreviousMonth).AddMonths(1).AddSeconds(-1))

# Despite what you read the whenCreated attribute is of Type System.DateTime, so you can just
# do a direct comparision. For Example: $objComputer.whencreated -lt $StartOfPreviousMonth

# However, some dates are expressed in UTC (GMT) format in Active Directory, so depending on
# which ones you are reading you may first need to convert the date to UTC format, and then
# to a string in the format expected by the GeneralizedTime syntax to correctly represent the
# date and time. This returns a string accurate to seconds. It can be modified by using:
$creationDateStr = $StartOfPreviousMonth.ToString("u") -Replace "-|:|\s"
$creationDateStr = $creationDateStr -Replace "Z", ".0Z"

$OnLineArray = @()
$OffLineArray = @()

$objDomain = New-Object System.DirectoryServices.DirectoryEntry
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.SearchRoot = $objDomain
$objSearcher.PageSize = 1000
$objSearcher.Filter = "(&(objectCategory=computer)(OperatingSystem=Window*Server*))"
$objSearcher.SearchScope = "Subtree"
# Set the attributes that you want to be returned from AD
$colProplist = @("name","description","whencreated")
foreach ($i in $colPropList) {
  $objSearcher.PropertiesToLoad.Add($i) | Out-Null
}
$colResults = $objSearcher.FindAll()
write-host "Found"$colResults.Count"results..."
foreach ($objResult in $colResults) {
  $ComputerNotFound = "$false"
  $objComputer = $objResult.Properties
  $computer = $objComputer.name[0]

  # If the computer object was created before the 1st day of the previous month, the SPLA
  # costs can be fully associated to the service, otherwise print the date the computer was
  # created so that the licensing specialist can manage pro rata charges.
  If($objComputer.whencreated -lt $StartOfPreviousMonth) {
    $DateCreated = "Before Start Of Previous Month"
  } Else {
    $DateCreated = $objComputer.whencreated[0].ToString()
  }
  #$DateCreated
  #$objComputer.description[0]

  $Computer
  if(Test-Connection -Cn $Computer -BufferSize 16 -Count 1 -ea 0 -quiet) {

#    $RegKey = "SOFTWARE\Telstra\BuildInfo"
#    $RegValue = "Cost Centre"
#    TRY {
#      # Create an instance of the Registry Object and open the HKLM base key
#      $regbasekey = [microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$computer)
#      $regsubkey = $regbasekey.OpenSubKey($RegKey)
#      $CostCentre = $regsubkey.GetValue($RegValue)
#      If($CostCentre -eq "" -OR $CostCentre -eq $NULL) {
#        $CostCentre = "Unknown"
#      }
#    }
#    CATCH {
#      $CostCentre = "Unknown"
#    }

    #### Win32 class short name assignment - add -credential $cred where needed
    TRY {
      $colItems = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computer
      if ($?)
      {
        foreach ($objItem in $colItems) 
        {
          $OSRunning = $objItem.caption + " " + $objItem.OSArchitecture + " SP " + $objItem.ServicePackMajorVersion
          $OSServicePack = $objItem.ServicePackMajorVersion
          $OSVersion = $objItem.Version
          $OSBuildNumber = $objItem.BuildNumber
          $TotalAvailMemory = $objItem.totalvisiblememorysize/1kb
          $TotalVirtualMemory = $objItem.totalvirtualmemorysize/1kb
          $TotalFreeMemory = $objItem.FreePhysicalMemory/1kb
          $TotalFreeVirtualMemory = $objItem.FreeVirtualMemory/1kb
          $TotalMem = "{0:N2}" -f $TotalAvailMemory
          $TotalVirt = "{0:N2}" -f $TotalVirtualMemory
          $FreeMem = "{0:N2}" -f $TotalFreeMemory
          $FreeVirtMem = "{0:N2}" -f $TotalFreeVirtualMemory
          $date = Get-Date
          $uptime = $objItem.ConvertToDateTime($objItem.lastbootuptime)
        }
      } else {
        throw $error[0].Exception
      }

      $colItems = Get-WmiObject -Class Win32_BIOS -ComputerName $computer
      if ($?)
      {
        foreach ($objItem in $colItems) 
        {
          $BiosVersion = $objItem.Manufacturer + " " + $objItem.SMBIOSBIOSVERSION + " " + $objItem.ConvertToDateTime($objItem.Releasedate)
          $BiosSerialNumber = $objItem.SerialNumber

        }
      } else {
        throw $error[0].Exception
        #$Error[0].Exception.GetType().FullName()
      }

      $colItems = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computer
      if ($?)
      {
        foreach ($objItem in $colItems) 
        {
          if (($objItem.DomainRole -eq "0") -or ($objItem.DomainRole -eq "1")) {
            $Role = "Workstation"
          }
            elseif (($objItem.DomainRole -eq "2") -or ($objItem.DomainRole -eq "3")) {
            $Role = "Member Server"
          }
            elseif (($objItem.DomainRole -eq "4") -or ($objItem.DomainRole -eq "5")) {
            $Role = "Domain Controller"
          }
            else {
            $Role = "Unknown"
          }
        }
      } else {
        throw $error[0].Exception
        #$Error[0].Exception.GetType().FullName()
      }

      $colItems = Get-WmiObject -Class Win32_Processor -ComputerName $computer
      if ($?)
      {
        foreach ($objItem in $colItems) 
        {
          $CPUInfo = $objItem.Name + " & has " + $objItem.NumberOfCores + " Cores & the FSB is " + $objItem.ExtClock + " Mhz"
          $CPULOAD = $objItem.LoadPercentage
        }
      } else {
        throw $error[0].Exception
        #$Error[0].Exception.GetType().FullName()
      }

      $colItems = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $computer | Where-Object {$_.DriveType -eq 3}
      if ($?)
      {
        foreach ($objItem in $colItems) 
        {

        }
      } else {
        throw $error[0].Exception
        #$Error[0].Exception.GetType().FullName()
      }


      If ($OSVersion -like '6.*') {
        $colItems = Get-WmiObject -Class SoftwareLicensingProduct -ComputerName $computer | Where-Object {$_.LicenseStatus -NotMatch "0"}
        if ($?)
        {
          foreach ($objItem in $colItems) {

            $status = switch($objItem.LicenseStatus) {
              0 {"Unlicensed"} 
              1 {"Licensed"} 
              2 {"Out-Of-Box Grace Period"} 
              3 {"Out-Of-Tolerance Grace Period"} 
              4 {"Non-Genuine Grace Period"} 
              5 {"Notification"} 
              6 {"Extended Grace"} 
              default {"Unknown value"}
            }
            $ActivationStatus = "Activation Status: {0}" -f $status
            $Licensed = "$True"
            $PartialProductKey = $objItem.PartialProductKey
            $LicenseProductName = $objItem.Name

            if ($objItem.Description -like "*VOLUME_KMSCLIENT*") {
              $LicenseType = "VOLUME_KMSCLIENT"
            }
              elseif ($objItem.Description -like "*VOLUME_KMS_R2_C*") {
              $LicenseType = "VOLUME_KMS_R2_C"
            }
              elseif ($objItem.Description -like "*VOLUME_KMS_R2_B*") {
              $LicenseType = "VOLUME_KMS_R2_B"
            }
              elseif ($objItem.Description -like "*VOLUME_MAK_B*") {
              $LicenseType = "VOLUME_MAK_B"
            }
              elseif ($objItem.Description -like "*OEM_SLP*") {
              $LicenseType = "OEM_SLP"
            }
              elseif ($objItem.Description -like "*OEM_COA_NSLP*") {
              $LicenseType = "OEM_COA_NSLP"
            }
              elseif ($objItem.Description -like "*RETAIL*") {
              $LicenseType = "RETAIL"
            }
              elseif ($objItem.Description -like "*TIMEBASED_EVAL*") {
              $LicenseType = "TIMEBASED_EVAL"
            }
              else {
              $LicenseType = "Unknown"
            }
          }
        } else {
          throw $error[0].Exception
          #$Error[0].Exception.GetType().FullName()
        }
      } Else {
        $LicenseType = "N/A"
      }

      $colItems = Get-WmiObject -Class win32_quickfixengineering -ComputerName $computer | Select-Object -Property "Description", "HotfixID", @{Name="InstalledOn"; Expression={([DateTime]($_.InstalledOn)).ToLocalTime()}}  
      if ($?)
      {
        $InstalledHotfixes = @()
        foreach ($objItem in $colItems) 
        {
          $HotfixID = $objItem.HotfixID
          $InstalledHotfixes += $HotfixID
        }
      } else {
        throw $error[0].Exception
        #$Error[0].Exception.GetType().FullName()
      }

      $output = New-Object PSObject
      $output | Add-Member NoteProperty ComputerName $Computer
      $output | Add-Member NoteProperty SPVer $OSServicePack
      $output | Add-Member NoteProperty BuildNo $OSBuildNumber
      $output | Add-Member NoteProperty BIOSSerial $BiosSerialNumber
      $output | Add-Member NoteProperty InstalledHotfixes $InstalledHotfixes
      $OnLineArray += $output

    }
    CATCH
    {
      $ComputerNotFound = "$true"
    }
  }
  Else
  {
    $ComputerNotFound = "$true"
  }
  if ($ComputerNotFound -eq $true) {
    $outputbad = New-Object PSObject
    $outputbad | Add-Member NoteProperty ComputerName $Computer
    $OffLineArray += $outputbad
  }
}

$OnLineArray | Export-CSV -notype "$OnlineFile"
$OffLineArray | Export-CSV -notype "$OfflineFile"

# Remove the quotes
(get-content "$OnlineFile") |% {$_ -replace '"',""} | out-file "$OnlineFile" -Fo -En ascii
(get-content "$OfflineFile") |% {$_ -replace '"',""} | out-file "$OfflineFile" -Fo -En ascii
