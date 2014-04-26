<#
  This script will get the TimeZone information for all servers in the
  domain and create a CSV file.

  Script Name: Get-TimeZoneInformation.ps1
  Release 1.1
  Modified by Jeremy@jhouseconsulting.com 24th February 2014
  Written by Jeremy@jhouseconsulting.com 20th November 2013

  To retrieve the timezone information can use one of four methods.

  1) The tzutil.exe utility
     tzutil does not have an option to connect to a remote computer, so we use the invoke-command cmdlet
     to connect via WinRM and execute the command.
       invoke-command -ComputerName PTHMSADDS02 {tzutil /g}
     Some have reported that this is not reliable and you should use a Windows PowerShell session
     (PSSession)
       $session = New-PSSession -ComputerName PTHMSADDS02
       $result = Invoke-Command -Session $session {tzutil /g}
       $result
       Remove-PSSession -Session $session

  2) The Win32_TimeZone WMI Class
  The DaylightBias property is a 32-bit integer that specifies the bias in minutes. It is a property of
  the SPTimeZoneInformation structure gets the bias in the number of minutes that daylight time for the
  time zone differs from Coordinated Universal Time (UTC).
  ie. the DaylightBias property gets the difference, in minutes, between UTC and local time (in daylight
  savings time). UTC = local time + bias.

  3) The System.TimeZone .Net Framework class
     This method can only retrieve local timezone information.

  4) The Registry
     HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\TimeZoneInformation

  There is a better chance that the WMI (RPC) ports are open to remote computers than WinRM, and sometimes
  WinRM isn't event enabled, so I choose to use the WMI method. We could also use the remote registry method.

  Note that setting the timezone information from a command line or script can only be done via tzutil.
  Therefore, if you were to script it using PowerShell it would need to be wrapped with the invoke-command
  cmdlet.

#>

#-------------------------------------------------------------
# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$OnlineFile = $(&$ScriptPath) + "\TimeZone-online.csv"
$OfflineFile = $(&$ScriptPath) + "\TimeZone-offline.csv"

# Set this value to true if you want to see the progress bar.
$ProgressBar = $True

#-------------------------------------------------------------

Import-Module ActiveDirectory 

# Get Domain Controllers
#$Computers = Get-ADDomainController -Filter * | Sort-Object Name

# Get All Servers filtering out Cluster Name Objects (CNOs) and Virtual computer Objects (VCOs) 
$Computers = Get-ADComputer -Filter * -Properties Name,Operatingsystem | Where-Object {($_.Operatingsystem -like '*server*') -AND !($_.serviceprincipalname -like '*MSClusterVirtualServer*')} | Sort-Object Name

#-------------------------------------------------------------

$Count = $Computers.Count
$TotalProcessed = 0
$onlinearray = @()
$offlinearray = @()

ForEach($Computer in $Computers){
  $ComputerError = "$false"
  if (Test-Connection -Cn $Computer.Name -BufferSize 16 -Count 1 -ea 0 -quiet) {
    write-Host -ForegroundColor Green "Checking time source and configuration of"($Computer.Name)
    Try {
      # Make all errors terminating
      $ErrorActionPreference = "Stop"
      $TimeZone = Get-WmiObject -Class Win32_TimeZone -ComputerName $($Computer.Name) | select DaylightBias, Caption, Bias, StandardBias, Description, DaylightName, StandardName
      $Caption = $TimeZone.Caption
      $Bias = $TimeZone.Bias
      $DaylightBias = $TimeZone.DaylightBias
      $output = New-Object PSObject
      $output | Add-Member NoteProperty -Name "ComputerName" $Computer.Name
      $output | Add-Member NoteProperty -Name "TimeZone" $Caption
      $output | Add-Member NoteProperty -Name "Bias" $Bias
      $output | Add-Member NoteProperty -Name "DaylightBias" $DaylightBias
      $onlinearray += $output
    }
    Catch {
      $ErrorDescription = $_.Exception.Message
      write-Host -ForegroundColor Red "Failed to access"$Computer.Name": "$_.Exception.Message
      $ComputerError = "$true"
    }
    Finally {
      # Reset the error action pref to default
      $ErrorActionPreference = "Continue"
    }
  } Else {
    $ComputerError = "$true"
    $ErrorDescription = "Unable to ping server"
    write-Host -ForegroundColor Red ($Computer.Name)"is offline"
  }
  if ($ComputerError -eq $true) {
    $outputbad = New-Object PSObject
    $outputbad | Add-Member NoteProperty ComputerName $Computer.Name
    $outputbad | Add-Member NoteProperty ErrorDescription $ErrorDescription
    $offlinearray += $outputbad
  }
  $TotalProcessed ++
  If ($ProgressBar) {
    Write-Progress -Activity 'Processing Users' -Status ("Username: {0}" -f $($Computer.Name)) -PercentComplete (($TotalProcessed/$Count)*100)
  }
}
$onlinearray | export-csv -notype "$OnlineFile" -Delimiter ';'
$offlinearray | export-CSV -notype "$OfflineFile"

# Remove the quotes
(get-content "$OnlineFile") |% {$_ -replace '"',""} | out-file "$OnlineFile" -Fo -En ascii
(get-content "$OfflineFile") |% {$_ -replace '"',""} | out-file "$OfflineFile" -Fo -En ascii
