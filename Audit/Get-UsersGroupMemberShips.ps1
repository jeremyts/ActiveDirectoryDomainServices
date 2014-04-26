<#
  Get-UsersGroupMemberShips.ps1
  How Can I Generate a List of All Groups of Which a User Is a Member?
  - http://blogs.technet.com/b/heyscriptingguy/archive/2009/10/08/hey-scripting-guy-october-8-2009.aspx

#>

#-------------------------------------------------------------

Function New-Underline($Text)
{
  "`n$Text`n$(`"-`" * $Text.length)"
} #end New-UnderLine

Function Test-DotNetFrameWork35
{
 Test-path -path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5'
} #end Test-DotNetFrameWork35

Function Get-UserPrincipal($cName, $cContainer, $userName)
{
 $dsam = "System.DirectoryServices.AccountManagement" 
 $rtn = [reflection.assembly]::LoadWithPartialName($dsam)
 $cType = "domain" #context type
 $iType = "SamAccountName"
 $dsamUserPrincipal = "$dsam.userPrincipal" -as [type]
 $principalContext = new-object "$dsam.PrincipalContext"($cType,$cName,$cContainer)
 $dsamUserPrincipal::FindByIdentity($principalContext,$iType,$userName)
} # end Get-UserPrincipal


If(-not(Test-DotNetFrameWork35)) { “Requires .NET Framework 3.5” ; exit }

#-------------------------------------------------------------

[string]$userName = "admintbird"
[string]$cName = "fmg.local"
[string]$cContainer = "DC=FMG,DC=local"

#-------------------------------------------------------------

$userPrincipal = Get-UserPrincipal -userName $userName -cName $cName -cContainer $cContainer

New-UnderLine -Text "Direct Group MemberShip:"
$userPrincipal.getGroups() | foreach-object { $_.name }

New-UnderLine -Text "Indirect Group Membership:"
$userPrincipal.GetAuthorizationGroups()  | foreach-object { $_.name }
