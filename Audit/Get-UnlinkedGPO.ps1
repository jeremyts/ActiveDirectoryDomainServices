import-module grouppolicy

function IsNotLinked($xmldata){
    If ($xmldata.GPO.LinksTo -eq $null) {
        Return $true
    }
    
    Return $false
}

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}

$ReferenceFile = $(&$ScriptPath) + "\UnlinkedGPO.csv"

$unlinkedGPOs = @()

Get-GPO -All | ForEach { $gpo = $_ ; $_ | Get-GPOReport -ReportType xml | ForEach { If(IsNotLinked([xml]$_)){$unlinkedGPOs += $gpo} }}

If ($unlinkedGPOs.Count -eq 0) {
    "No Unlinked GPO's Found"
}
Else{
    write-host $unlinkedGPOs.count"GPO's are unlinked:"
    $unlinkedGPOs | Select DisplayName,ID | ft
    $unlinkedGPOs | Select DisplayName,ID,CreationTime,ModificationTime,GpoStatus,Description | export-csv -notype "$ReferenceFile" -Delimiter ';'

    # Remove the quotes
    (get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii
}
