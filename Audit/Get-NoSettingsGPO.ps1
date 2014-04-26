import-module grouppolicy

function HasNoSettings{
    $cExtNodes = $xmldata.DocumentElement.SelectNodes($cQueryString, $XmlNameSpaceMgr)
  
    foreach ($cExtNode in $cExtNodes){
        If ($cExtNode.HasChildNodes){
            Return $false
        }
    }
    
    $uExtNodes = $xmldata.DocumentElement.SelectNodes($uQueryString, $XmlNameSpaceMgr)
    
    foreach ($uExtNode in $uExtNodes){
       If ($uExtNode.HasChildNodes){
            Return $false
        }
    }
    
    Return $true
}

function configNamespace{
    $script:xmlNameSpaceMgr = New-Object System.Xml.XmlNamespaceManager($xmldata.NameTable)

    $xmlNameSpaceMgr.AddNamespace("", $xmlnsGpSettings)
    $xmlNameSpaceMgr.AddNamespace("gp", $xmlnsGpSettings)
    $xmlNameSpaceMgr.AddNamespace("xsi", $xmlnsSchemaInstance)
    $xmlNameSpaceMgr.AddNamespace("xsd", $xmlnsSchema)
}

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}

$ReferenceFile = $(&$ScriptPath) + "\NoSettingsGPO.csv"

$noSettingsGPOs = @()

$xmlnsGpSettings = "http://www.microsoft.com/GroupPolicy/Settings"
$xmlnsSchemaInstance = "http://www.w3.org/2001/XMLSchema-instance"
$xmlnsSchema = "http://www.w3.org/2001/XMLSchema"

$cQueryString = "gp:Computer/gp:ExtensionData/gp:Extension"
$uQueryString = "gp:User/gp:ExtensionData/gp:Extension"

Get-GPO -All | ForEach { $gpo = $_ ; $_ | Get-GPOReport -ReportType xml | ForEach { $xmldata = [xml]$_ ; configNamespace ; If(HasNoSettings){$noSettingsGPOs += $gpo} }}

If ($noSettingsGPOs.Count -eq 0) {
    "No GPO's Without Settings Were Found"
}
Else{
    write-host $noSettingsGPOs.count"GPO's were found without settings:"
    $noSettingsGPOs | Select DisplayName,ID | ft
    $noSettingsGPOs | Select DisplayName,ID,CreationTime,ModificationTime,GpoStatus,Description | export-csv -notype "$ReferenceFile" -Delimiter ';'

    # Remove the quotes
    (get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii
}
