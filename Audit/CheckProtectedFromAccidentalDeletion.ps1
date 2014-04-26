# This script will check if all OU's and default top level Containers
# have been protected from accidental deletion.

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}

$ReferenceFile = $(&$ScriptPath) + "\CheckProtectedFromAccidentalDeletion.csv"

# Import the Modules
Import-Module ActiveDirectory
$defaultNamingContext = (get-adrootdse).defaultnamingcontext

$array = @()

# Check protection on all OU's
$OUs = Get-ADOrganizationalUnit -filter * -Properties ProtectedFromAccidentalDeletion | where {$_.ProtectedFromAccidentalDeletion -match "False"}
write-host -ForegroundColor green "The following"$OUs.Count"OUs have not been protected from accidental deletion..."
ForEach ($OU in $OUs) {
  $output = New-Object PSObject
  $output | Add-Member NoteProperty objectClass ($OU.objectClass)
  $output | Add-Member NoteProperty DistinguishedName ($OU.DistinguishedName)
  $output | Add-Member NoteProperty Description ($OU.Description)
  $array += $output
}

# Check protection on all default top level Containers
$Containers = Get-ADObject -filter 'objectClass -eq "container"' -SearchBase $defaultNamingContext -SearchScope OneLevel -Properties ProtectedFromAccidentalDeletion | where {$_.ProtectedFromAccidentalDeletion -match "False"}
write-host -ForegroundColor green "The following"$Containers.Count"top level containers have not been protected from accidental deletion..."
ForEach ($Container in $Containers) {
  $output = New-Object PSObject
  $output | Add-Member NoteProperty objectClass ($Container.objectClass)
  $output | Add-Member NoteProperty DistinguishedName ($Container.DistinguishedName)
  $output | Add-Member NoteProperty Description ($Container.Description)
  $array += $output
}

$array | export-csv -notype "$ReferenceFile" -Delimiter ";"

# Remove the quotes
(get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii
