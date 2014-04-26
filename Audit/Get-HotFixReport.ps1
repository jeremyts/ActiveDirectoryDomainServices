<# 
 
 NAME: Get-HotFixReport.ps1 
 
 AUTHOR: Jan Egil Ring 
 EMAIL: jan.egil.ring@powershell.no 
 
 COMMENT: Script to generate an installation-report for specified hotfixes on a set of computers retrieved from Active Directory. 
  
          The script leverages Microsoft`s PowerShell-module for Active Directory to retrieve computers. 
          Before running the script, customize the three variables under #Custom variables 
           
          More information: http://blog.powershell.no/2010/10/31/generate-an-installation-report-for-specific-hotfixes-using-windows-powershell 
 
 You have a royalty-free right to use, modify, reproduce, and 
 distribute this script file in any way you find useful, provided that 
 you agree that the creator, owner above has no warranty, obligations, 
 or liability for such use. 
 
 VERSION HISTORY: 
 1.0 31.10.2010 - Initial release 
 
#> 
 
#requires -version 2 
 
#Import Active Directory module 
Import-Module ActiveDirectory 

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}

#Custom variables 
#$Computers = Get-ADComputer -Filter * -Properties Name,Operatingsystem | Where-Object {$_.Operatingsystem -like "*server*"} | Select-Object Name 
$Computers = @()
$HotFix = "KB979744,KB979744,KB983440,KB979099,KB982867,KB977020" 
$CsvFilePath = $(&$ScriptPath) + "\Hotfixes.csv" 

# Get Computer Name
$ComputerName = (get-content env:COMPUTERNAME)

$Computers += $ComputerName

#Variable for writing progress information 
$TotalComputers = ($computers | Measure-Object).Count 
$CurrentComputer = 1 
  
#Create array to hold hotfix information 
$Export = @() 
 
#Splits the array if more than one hotfix are provided 
$Hotfixes = $HotFix.Split(",") 
 
#Loop through every computers   
foreach ($computer in $computers) { 
 
#Loop through every hotfix 
foreach ($hotfix in $hotfixes) { 
#Write progress information 
Write-Progress -Activity "Checking for hotfix $hotfix..." -Status "Current computer: $computer" -Id 1 -PercentComplete (($CurrentComputer/$TotalComputers) * 100) 
 
#Create a custom object for each hotfix 
$obj = New-Object -TypeName psobject 
$obj | Add-Member -Name Hotfix -Value $hotfix -MemberType NoteProperty 
$obj | Add-Member -Name Computer -Value $computer -MemberType NoteProperty 
 
#Check if hotfix are installed  
 try { 
 if (Test-Connection -Count 1 -ComputerName $computer -Quiet) { 
 Get-HotFix -Id $hotfix -ComputerName $computer -ea stop | Out-Null 
 $obj | Add-Member -Name HotfixInstalled -Value $true -MemberType NoteProperty 
 $obj | Add-Member -Name ErrorEncountered -Value "None" -MemberType NoteProperty 
 } 
 else { 
   $obj | Add-Member -Name HotfixInstalled -Value $false -MemberType NoteProperty 
   $obj | Add-Member -Name ErrorEncountered -Value $error[0].Exception.Message -MemberType NoteProperty 
 } 
 } 
  
 catch { 
   $obj | Add-Member -Name HotfixInstalled -Value $false -MemberType NoteProperty 
   $obj | Add-Member -Name ErrorEncountered -Value $error[0].Exception.Message -MemberType NoteProperty    
 } 
  
#Add the custom object to the array to be exported 
$Export += $obj 
 
} 
 
#Increase counter variable 
$CurrentComputer ++ 
 
 } 

#Export the array with hotfix-information to the user-specified path 
$Export | Export-Csv -Path $CsvFilePath -NoTypeInformation
