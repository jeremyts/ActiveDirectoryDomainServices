<#
  This script will Export and Import the OU Structure

  Syntax examples:
    Export:
      ManageOUStructure.ps1 -Action Export -ReferenceFile OUExport.csv

    Export:
      ManageOUStructure.ps1 -Action Import -ReferenceFile OUExport.csv

  You could indeed use ldifde, but I find this method provides far more
  flexibility with the manipulation of the data in a simple format.

  Release 1.1
  Written by Jeremy@jhouseconsulting.com 13th September 2013
  Modified by Jeremy@jhouseconsulting.com 28th January 2014
#>

#-------------------------------------------------------------
param([String]$Action,[String]$ReferenceFile)

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}

if ([String]::IsNullOrEmpty($Action)) {
  write-host -ForeGroundColor Red "Action is a required parameter. Exiting Script.`n"
  exit
} else {
  switch ($Action)
  {
    "Import" {$Import = $true;$Export = $false}
    "Export" {$Import = $false;$Export = $true}
    default {$Import = $false;$Export = $false}
  }
  if ($Import -eq $false -AND $Export -eq $false) {
    write-host -ForeGroundColor Red "The Action parameter is invalid. Exiting Script.`n"
    exit
  }
}

if ([String]::IsNullOrEmpty($ReferenceFile)) {
  write-host -ForeGroundColor Red "ReferenceFile is a required parameter. Exiting Script.`n"
  exit
} else {
  $ReferenceFile = $(&$ScriptPath) + "\$ReferenceFile";
}

#-------------------------------------------------------------

# Import the Active Directory Module
Import-Module ActiveDirectory -WarningAction SilentlyContinue
if($Error.Count -eq 0) {
   #Write-Host "Successfully loaded Active Directory Powershell's module" -ForeGroundColor Green
}else{
   Write-Host "Error while loading Active Directory Powershell's module : $Error" -ForeGroundColor Red
   exit
}

#-------------------------------------------------------------

$defaultNamingContext = (get-adrootdse).defaultnamingcontext

#-------------------------------------------------------------
If ($Import -eq $true) {

  if ((Test-Path $ReferenceFile) -eq $False) {
    Write-Host -ForegroundColor Red "The $ReferenceFile file is missing.`n"
    exit
  }

  $AD_OU_LIST = Import-Csv -Path "$ReferenceFile" -Delimiter ';'

  #Sort the data using the OUPath column
  $AD_OU_LIST = $AD_OU_LIST | Sort Path

  foreach($OU_Object in $AD_OU_LIST) {

    #Clear the path, as it's rebuilt with every loop
    $OU_Path = ""

    #Split out the OU Path from the CSV File
    $tmpOUPath = $OU_Object.Path.Split("|")
    If($tmpOUPath.Count -eq 1)
    {
      $OU_Name= $tmpOUPath[0]
    }
    Else
    {
      # Reverse the Path
      [array]::Reverse($tmpOUPath)
      $OU_Name= $tmpOUPath[0]

      $i = 0
      ForEach ($subOU in $tmpOUPath)
      {
        $i = $i + 1
        If ($i -eq 1)
        {
          $OU_Name= $subOU
        }
        else
        {
          $OU_Path = $OU_Path + "OU=" + $subOU + ","
        }
      }
    }
    $OU_Description = $OU_Object.Description
    $OU_Protect = $OU_Object.Protect
    # Convert to boolean
    if($OU_Protect.ToLower() -eq "true")
    {
      $OU_Protect = [System.Convert]::ToBoolean("$True")
    }
    else
    {
      $OU_Protect = [System.Convert]::ToBoolean("$False")
    }
    $OU_Path = $OU_Path + $defaultNamingContext

    Try {
      # Check if the target OU exists. If not, create it.
      $ExistingOU = Get-ADOrganizationalUnit -Filter { name -eq $OU_Name } -SearchBase $OU_Path -SearchScope OneLevel
      }
    Catch {
      $ExistingOU = $NULL
      }

    if($ExistingOU -eq $null)
    {
      New-ADOrganizationalUnit -Name $OU_Name -Path $OU_Path -Description $OU_Description -ProtectedFromAccidentalDeletion $OU_Protect
    }
  }
}

#-------------------------------------------------------------
If ($Export -eq $true) {

  $array = @()

  $AD_OU_LIST = Get-ADOrganizationalUnit -Filter * -SearchBase $defaultNamingContext -Properties Description,ProtectedFromAccidentalDeletion

  ForEach ($OU in $AD_OU_LIST) {

    $OUPath = $OU.DistinguishedName -replace (",$defaultNamingContext","")
    $OUPath = $OUPath -replace ("OU=","")
    $OUPath = $OUPath -replace (",","|")
    ForEach ($item in $OUPath) {
      If ($Item -ne "Domain Controllers") {

        $tmpOUPath = $Item.Split("|")
        If($tmpOUPath.Count -eq 1)
        {
          $OU_Name = $tmpOUPath[0]
        }
        Else
        {
          # Reverse the Path
          [array]::Reverse($tmpOUPath)
          $OU_Name= $tmpOUPath[0]
          $i = 0
          ForEach ($subOU in $tmpOUPath)
          {
            $i = $i + 1
            If ($i -eq 1)
            {
              $OU_Name = $subOU
            }
            else
            {
              $OU_Name = $OU_Name + "|" + $subOU
            }
          }
        }

        $output = New-Object PSObject
        $output | Add-Member NoteProperty Path $OU_Name
        $output | Add-Member NoteProperty Description ($OU.Description)
        $output | Add-Member NoteProperty Protect ($OU.ProtectedFromAccidentalDeletion.ToString().ToLower())
        $array += $output

      }
    }
  }

  $array | export-csv -notype "$ReferenceFile" -Delimiter ";"

  # Remove the quotes
  (get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii

}
