<#
    .SYSNOPSIS
        Generates a report showing all logon scripts being used in a GPO.

    .DESCRIPTION
        Generates a report showing all logon scripts being used in a GPO. Scans all of the GPOs in a domain.

    .NOTES
        Name: Get-GPOLogonScriptReport
        Author: Boe Prox
        Created: 05 Oct 2013

    .EXAMPLE
        .\Get-GPOLogonScriptReport.ps1 | Export-Csv -NotTypeInformation -Path 'GPOLogonScripts.csv'

        Description
        -----------
        Generates a report of all GPOs using logon scripts and then exports the data to a CSV file.


   Re-wrote to optimise the code for large environments to avoid inconsistent results and
   'System.OutOfMemoryException' errors.

   Merged the function written by Jason Yonder
   http://mctexpert.blogspot.com.au/2013/02/list-all-scripts-my-gpos-run.html

I've found that I sometimes get the following error if the GPMC is open when running the script.
WARNING: Operation is not valid due to the current state of the object.


#>
# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$ReferenceFile = $(&$ScriptPath) + "\GPOLogonScriptReport.csv"

Try {
    Import-Module GroupPolicy -ErrorAction Stop
    $ConsoleOutput = $True
    $i=0
    $AllScripts = @()
    $array = @()
    $count = (Get-GPO -All).Count
    Get-GPO -All |ForEach-Object {
        $DisplayName = $_.DisplayName
        $ID = $_.ID
        $GPOStatus = $_.GpoStatus
        $i++
        If ($ConsoleOutput) {
          Write-Progress -Activity 'GPO Scan' -Status ("GPO: {0}" -f $DisplayName) -PercentComplete (($i/$count)*100)
        }
        $xml = [xml]($_ | Get-GPOReport -ReportType XML)
        #User logon script
        $userScripts = @($xml.GPO.User.ExtensionData | Where {$_.Name -eq 'Scripts'})
        If ($userScripts.count -gt 0) {
            ForEach ($script in $userScripts.extension.Script) {
              $us = New-Object -TypeName PSObject
              $us | Add-Member -MemberType NoteProperty -Name "GPOName" -value $DisplayName
              $us | Add-Member -MemberType NoteProperty -Name "ID" -value $ID
              $us | Add-Member -MemberType NoteProperty -Name "GPOState" -value $GPOStatus
              $us | Add-Member -MemberType NoteProperty -Name "GPOType" -value 'User'
              $us | Add-Member -MemberType NoteProperty -Name "Type" -value $script.Type
              $us | Add-Member -MemberType NoteProperty -Name "Script" -value $script.command
              $us | Add-Member -MemberType NoteProperty -Name "ScriptType" -value ($script.command -replace '.*\.(.*)','$1')
              $UserScript += $uS
              $AllScripts += $US
              If ($ConsoleOutput) { write-output $uS }
            }
        } Else {
          $UserScript = @()
        }
        #Computer logon script
        $computerScripts = @($xml.GPO.Computer.ExtensionData | Where {$_.Name -eq 'Scripts'})
        If ($computerScripts.count -gt 0) {
            ForEach ($script in $computerScripts.extension.Script) {
              $cs = New-Object -TypeName PSObject
              $cs | Add-Member -MemberType NoteProperty -Name "GPOName" -value $DisplayName
              $cs | Add-Member -MemberType NoteProperty -Name "ID" -value $ID
              $cs | Add-Member -MemberType NoteProperty -Name "GPOState" -value $GPOStatus
              $cs | Add-Member -MemberType NoteProperty -Name "GPOType" -value 'Computer'
              $cs | Add-Member -MemberType NoteProperty -Name "Type" -value $script.Type
              $cs | Add-Member -MemberType NoteProperty -Name "Script" -value $script.command
              $cs | Add-Member -MemberType NoteProperty -Name "ScriptType" -value ($script.command -replace '.*\.(.*)','$1')
              $ComputerScript += $CS
              $AllScripts += $CS
              If ($ConsoleOutput) { write-output $CS }
            }
        } Else {
          $ComputerScript = @()
        }
        $Obj = New-Object -TypeName PSOBject
        $Obj | Add-Member -MemberType NoteProperty -Name "GPO" -Value $DisplayName
        $Obj | Add-Member -MemberType NoteProperty -Name "UserScript" -Value $UserScript
        $Obj | Add-Member -MemberType NoteProperty -Name "ComputerScript" -Value $ComputerScript
        $array += $obj
        #If ($ConsoleOutput) {Write-Output $obj}
    }

    #Write-Output $array | Format-Table
    $AllScripts | export-csv -notype -path "$ReferenceFile" -Delimiter ';'

    # Remove the quotes
    (get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii

} Catch {
    Write-Warning ("{0}" -f $_.exception.message)
}

