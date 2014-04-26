<#

  This script retrieves an AD group with 2 additional properties: 
  - NestedGroupMembershipCount
  - MaxNestingLevel

  This helps us to understand the maximum nesting levels of groups and the recursive group membership count.

  Original script written by M Ali.
   - Token Bloat Troubleshooting by Analyzing Group Nesting in AD
     http://blogs.msdn.com/b/adpowershell/archive/2009/09/05/token-bloat-troubleshooting-by-analyzing-group-nesting-in-ad.aspx

  Usage:
    Get-ADGroupNesting.ps1 CarAnnounce
    Get-ADGroupNesting.ps1 CarAnnounce -showtree

    When used with the –ShowTree parameter, the script displays the recursive group membership tree along with emitting the ADGroup object.
    When using a group name that contains a space use single quotes around the name.

#>

#-------------------------------------------------------------

Param ( 
    [Parameter(Mandatory=$true, 
        Position=0, 
        ValueFromPipeline=$true, 
        HelpMessage="DN or ObjectGUID of the AD Group." 
    )] 
    [string]$groupIdentity, 
    [switch]$showTree 
    )

$global:numberOfRecursiveGroupMemberships = 0 
$lastGroupAtALevelFlags = @()

# Import the Active Directory Module
Import-Module ActiveDirectory -WarningAction SilentlyContinue
if($Error.Count -eq 0) {
   Write-Host "Successfully loaded Active Directory Powershell's module" -ForeGroundColor Green
}else{
   Write-Host "Error while loading Active Directory Powershell's module : $Error" -ForeGroundColor Red
   exit
}

function Get-GroupNesting ([string] $identity, [int] $level, [hashtable] $groupsVisitedBeforeThisOne, [bool] $lastGroupOfTheLevel) 
{ 
    $group = $null 
    $group = Get-ADGroup -Identity $identity -Properties "memberOf"    
    if($lastGroupAtALevelFlags.Count -le $level) 
    { 
        $lastGroupAtALevelFlags = $lastGroupAtALevelFlags + 0 
    } 
    if($group -ne $null) 
    { 
        if($showTree) 
        { 
            for($i = 0; $i -lt $level - 1; $i++) 
            { 
                if($lastGroupAtALevelFlags[$i] -ne 0) 
                { 
                    Write-Host -ForegroundColor Yellow -NoNewline "  " 
                } 
                else 
                { 
                    Write-Host -ForegroundColor Yellow -NoNewline "¦ " 
                } 
            } 
            if($level -ne 0) 
            { 
                if($lastGroupOfTheLevel) 
                { 
                    Write-Host -ForegroundColor Yellow -NoNewline "+-" 
                } 
                else 
                { 
                    Write-Host -ForegroundColor Yellow -NoNewline "+-" 
                } 
            } 
            Write-Host -ForegroundColor Yellow $group.Name 
        } 
        $groupsVisitedBeforeThisOne.Add($group.distinguishedName,$null) 
        $global:numberOfRecursiveGroupMemberships ++ 
        $groupMemberShipCount = $group.memberOf.Count 
        if ($groupMemberShipCount -gt 0) 
        { 
            $maxMemberGroupLevel = 0 
            $count = 0 
            foreach($groupDN in $group.memberOf) 
            { 
                $count++ 
                $lastGroupOfThisLevel = $false 
                if($count -eq $groupMemberShipCount){$lastGroupOfThisLevel = $true; $lastGroupAtALevelFlags[$level] = 1} 
                if(-not $groupsVisitedBeforeThisOne.Contains($groupDN)) #prevent cyclic dependancies 
                { 
                    $memberGroupLevel = Get-GroupNesting -Identity $groupDN -Level $($level+1) -GroupsVisitedBeforeThisOne $groupsVisitedBeforeThisOne -lastGroupOfTheLevel $lastGroupOfThisLevel 
                    if ($memberGroupLevel -gt $maxMemberGroupLevel){$maxMemberGroupLevel = $memberGroupLevel} 
                } 
            } 
            $level = $maxMemberGroupLevel 
        } 
        else #we've reached the top level group, return it's height 
        { 
            return $level 
        } 
        return $level 
    } 
} 
$global:numberOfRecursiveGroupMemberships = 0 
$groupObj = $null 
$groupObj = Get-ADGroup -Identity $groupIdentity 
if($groupObj) 
{ 
    [int]$maxNestingLevel = Get-GroupNesting -Identity $groupIdentity -Level 0 -GroupsVisitedBeforeThisOne @{} -lastGroupOfTheLevel $false 
    Add-Member -InputObject $groupObj -MemberType NoteProperty  -Name MaxNestingLevel -Value $maxNestingLevel -Force 
    Add-Member -InputObject $groupObj -MemberType NoteProperty  -Name NestedGroupMembershipCount -Value $($global:numberOfRecursiveGroupMemberships - 1) -Force 
    $groupObj 
}
