<#
  This script will get all the missing subnets from the NETLOGON.LOG file from each
  Domain Controller in the Domain. It does this by copying all the NETLOGON.LOG files
  locally and then parsing them all to create a CSV output of unique IP Addresses.
  The CSV file is sorted by IP Address to make it easy to group them into subnets.

  Script Name: Find_missing_subnets_in_ActiveDirectory.ps1
  Release 1.2
  Modified by Jeremy@jhouseconsulting.com 23/01/2014
  Written by Jeremy@jhouseconsulting.com 02/01/2014

  Syntax examples:

  - To execute the script in the current Domain:
      Find_missing_subnets_in_ActiveDirectory.ps1

  - To execute the script in a trusted Domain:
      Find_missing_subnets_in_ActiveDirectory.ps1 -TrustedDomain mydemosthatrock.com

  This script was derived from the AD-Find_missing_subnets_in_ActiveDirectory.ps1
  script written by Francois-Xavier CAT.
   - Report the AD Missing Subnets from the NETLOGON.log
     http://www.lazywinadmin.com/2013/10/powershell-report-ad-missing-subnets.html

  Changes:
  - Stripped down the code to remove the e-mail functionality. This is a nice to
    have feature and can be added back in for a future release. I felt that it was
    more important to focus on ensuring the core functionality of the script was
    working correctly and efficiently.

  Improvements:
  - Reordered the Netlogon.log collection to make it more efficient.
  - Implemented a fix to deal with the changes to the fields in the Netlogon.log
    file from Windows 2012 and above:
    - http://www.jhouseconsulting.com/2013/12/13/a-change-to-the-fields-in-the-netlogon-log-file-from-windows-2012-and-above-1029
  - Tidied up the way it writes the CSV file.
  - Changed the write-verbose and write-warning messages to write-host to vary the
    message colors and improve screen output.
  - Added a "replay" feature so that you have the ability to re-create the CSV
    from collected log files.
#>
#-------------------------------------------------------------
param([String]$TrustedDomain)
#-------------------------------------------------------------

# Set this to the last number of lines to read from each NETLOGON.log file.
# This allows the report to contain the most recent and relevant errors.
[Int]$LogsLines = "200"

# Set this to $True to remove txt and log files from the output folder.
$Cleanup = $True

# Set this to $True if you have not removed the log files and want to replay
# them to create a CSV.
$ReplayLogFiles = $False

#-------------------------------------------------------------

# PATH Information 
$ScriptPath = (Split-Path -Path ((Get-Variable -Name MyInvocation).Value).MyCommand.Path)
$ScriptPathOutput = $ScriptPath + "\Output"

# Date and Time Information
$DateFormat = Get-Date -Format "yyyyMMdd_HHmmss"

$OutputFile = "$scriptPathOutput\$DateFormat-AD-Sites-MissingSubnets.csv"

$CombineAndProcess = $False

IF ($ReplayLogFiles -eq $False)
{
  IF (-not(Test-Path -Path $ScriptPathOutput))
  {
    Write-Host -ForegroundColor green "Creating the Output Folder: $ScriptPathOutput"
    New-Item -Path $ScriptPathOutput -ItemType Directory | Out-Null
  }

  if ([String]::IsNullOrEmpty($TrustedDomain)) {
    # Get the Current Domain Information
    $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
  } else {
    $context = new-object System.DirectoryServices.ActiveDirectory.DirectoryContext("domain",$TrustedDomain)
    Try {
      $domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($context)
    }
    Catch [exception] {
      write-host -ForegroundColor red $_.Exception.Message
      Exit
    }
  }

  Write-Host -ForegroundColor green "Domain: $domain"

  # Get the names of all the Domain Contollers in $domain
  Write-Host -ForegroundColor green "Getting all Domain Controllers from $domain ..."
  $DomainControllers = $domain | ForEach-Object -Process { $_.DomainControllers } | Select-Object -Property Name

  # Gathering the NETLOGON.LOG for each Domain Controller
  Write-Host -ForegroundColor green "Processing each Domain controller..."
  FOREACH ($dc in $DomainControllers)
  {
    $DCName = $($dc.Name)

    # Get the Current Domain Controller in the Loop
    Write-Host -ForegroundColor green "Gathering the log from $DCName..."

    IF (Test-Connection -Cn $DCName -BufferSize 16 -Count 1 -ea 0 -quiet) {

      # NETLOGON.LOG path for the current Domain Controller
      $path = "\\$DCName\admin`$\debug\netlogon.log"

      # Testing the $path
      IF ((Test-Path -Path $path) -and ((Get-Item -Path $path).Length -ne $null))
      {
        # Copy the NETLOGON.log locally for the current DC
        Write-Host -ForegroundColor green "- Copying the $path file..."
        $TotalTime = measure-command {Copy-Item -Path $path -Destination $ScriptPathOutput\$($dc.Name)-$DateFormat-netlogon.log}
        $TotalSeconds = $TotalTime.TotalSeconds
        Write-Host -ForegroundColor green "- Copy completed in $TotalSeconds seconds."

        IF ((Get-Content -Path $path | Measure-Object -Line).lines -gt 0)
        {
          # Export the $LogsLines last lines of the NETLOGON.log and send it to a file
          ((Get-Content -Path $ScriptPathOutput\$DCName-$DateFormat-netlogon.log -ErrorAction Continue)[-$LogsLines .. -1]) | 
            Foreach-Object {$_ -replace "\[\d{1,5}\] ", ""} |
            Out-File -FilePath "$ScriptPathOutput\$DCName.txt" -ErrorAction 'Continue' -ErrorVariable ErrorOutFileNetLogon
          Write-Host -ForegroundColor green "- Exported the last $LogsLines lines to $ScriptPathOutput\$DCName.txt."
        }#IF
        ELSE {Write-Host -ForegroundColor green "- File Empty."}

      } ELSE {Write-Host -ForegroundColor red "- $DCName is not reachable via the $path path."}

    } ELSE {Write-Host -ForegroundColor red "- $DCName is not reachable or offline."}

    $CombineAndProcess = $True

  }#FOREACH

} ELSE {

  Write-Host -ForegroundColor green "Replaying the log files..."
  IF (Test-Path -Path $ScriptPathOutput)
  {
    IF ((Get-ChildItem $scriptpathoutput\*.log | Measure-Object).Count -gt 0)
    {
      $LogFiles = Get-ChildItem $scriptpathoutput\*.log

      ForEach ($LogFile in $LogFiles)
      {
        $DCName = $LogFile.Name -Replace("-\d{7,8}_\d{6}-netlogon.log")
        Write-Host -ForegroundColor green "Processing the log from $DCName..."
        IF ((Get-Content -Path "$ScriptPathOutput\$($LogFile.Name)" | Measure-Object -Line).lines -gt 0)
        {
          # Export the $LogsLines last lines of the NETLOGON.log and send it to a file
          ((Get-Content -Path "$ScriptPathOutput\$($LogFile.Name)" -ErrorAction Continue)[-$LogsLines .. -1]) | 
                    Foreach-Object {$_ -replace "\[\d{1,5}\] ", ""} |
                    Out-File -FilePath "$ScriptPathOutput\$DCName.txt" -ErrorAction 'Continue' -ErrorVariable ErrorOutFileNetLogon
          Write-Host -ForegroundColor green "- Exported the last $LogsLines lines to $ScriptPathOutput\$DCName.txt."
        } ELSE {Write-Host -ForegroundColor green "- File Empty."}
        $CombineAndProcess = $True
      }#ForEach
    } ELSE {Write-Host -ForegroundColor red "There are no log files to process."}
  } ELSE {Write-Host -ForegroundColor red "The $ScriptpathOutput folder is missing."}
}#IF

IF ($CombineAndProcess)
{

  # Combine all the TXT file in one
  $FilesToCombine = Get-Content -Path "$ScriptPathOutput\*.txt" -Exclude "*All_Export.txt" -ErrorAction SilentlyContinue |
    Foreach-Object {$_ -replace "\[\d{1,5}\] ", ""}

  if ($FilesToCombine)
  {
    $FilesToCombine | Out-File -FilePath $ScriptPathOutput\$dateformat-All_Export.txt

    # Convert the TXT file to a CSV format
    Write-Host -ForegroundColor green "Importing exported data to a CSV format..."
    $importString = Import-Csv -Path $scriptpathOutput\$dateformat-All_Export.txt -Delimiter ' ' -Header Date,Time,Domain,Error,Name,IPAddress

    # Get Only the entries for the Missing Subnets
    $MissingSubnets = $importString | Where-Object {$_.Error -like "*NO_CLIENT_SITE*"}
    Write-Host -ForegroundColor green "Total of NO_CLIENT_SITE errors found within the last $LogsLines lines across all log files: $($MissingSubnets.count)"
    # Get the other errors from the log
    $OtherErrors = Get-Content $scriptpathOutput\$dateformat-All_Export.txt | Where-Object {$_ -notlike "*NO_CLIENT_SITE*"} | Sort-Object -Unique
    Write-Host -ForegroundColor green "Total of other Error(s) found within the last $LogsLines lines across all log files: $($OtherErrors.count)"

    # Export to a CSV File
    $UniqueIPAddresses = $importString | Select-Object -Property Date, Name, IPAddress, Domain, Error | 
    Sort-Object -Property IPAddress -Unique
    $UniqueIPAddresses | Export-Csv -notype -path "$OutputFile"
    # Remove the quotes
    (get-content "$OutputFile") |% {$_ -replace '"',""} | out-file "$OutputFile" -Fo -En ascii
    Write-Host -ForegroundColor green "$($UniqueIPAddresses.count) unique IP Addresses exported to $OutputFile."

  }#IF File to Combine
  ELSE {Write-Host -ForegroundColor red "No .txt files to process."}

  IF ($Cleanup)
  {
    Write-Host -ForegroundColor green "Removing the .txt and .log files..."
    Remove-item -Path $ScriptpathOutput\*.txt -force
    Remove-Item -Path $ScriptPathOutput\*.log -force
  }

}

Write-Host -ForegroundColor green "Script Completed."
