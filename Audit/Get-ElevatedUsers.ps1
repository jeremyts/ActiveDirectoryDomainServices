# PowerShell function to list users in Authoritative Groups in Active Directory
# http://jeffwouters.nl/index.php/2013/11/powershell-function-to-list-users-in-authorative-groups-in-active-directory/

# Import the Modules
Import-Module ActiveDirectory

function Get-ElevatedUsers { 
    $GroupType = '-2147483643'
    # A group type of -2147483643 specifies all groups with a domain local scope created by the system.
    # References:
    # - http://msdn.microsoft.com/en-us/library/windows/desktop/ms675935(v=vs.85).aspx
    # - http://blogs.technet.com/b/heyscriptingguy/archive/2004/12/21/how-can-i-tell-whether-a-group-is-a-security-group-or-a-distribution-group.aspx
    $ElevatedGroups = Get-ADGroup -Filter {grouptype -eq $GroupType} -Properties members 
    $Elevatedgroups = $ElevatedGroups | Where-Object {($_.Name -ne 'Guests') -and ($_.Name -ne 'Users')} 
    foreach ($ElevatedGroup in $ElevatedGroups) { 
        $Members = $ElevatedGroup | Select-Object -ExpandProperty members 
        foreach ($Member in $Members) { 
            $Status = $true
            try { 
                $MemberIsUser = Get-ADUser $Member -ErrorAction silentlycontinue 
            } catch { $Status = $false} 
            if ($Status -eq $true) { 
                $Object = New-Object -TypeName PSObject 
                $Object | Add-Member -MemberType noteproperty -Name 'Group' -Value $ElevatedGroup.Name 
                $Object | Add-Member -MemberType noteproperty -name 'User' -Value $MemberIsUser.Name 
                $Object
            } else { 
                $Status = $true
                try { 
                    $GroupMembers = Get-ADGroup $Member -ErrorAction silentlycontinue | Get-ADGroupMember -Recursive -ErrorAction silentlycontinue 
                } catch { $Status = $false } 
                if ($Status -eq $true) { 
                    foreach ($GroupMember in $GroupMembers) { 
                        $Object = New-Object -TypeName PSObject 
                        $Object | Add-Member -MemberType noteproperty -Name 'Group' -Value $ElevatedGroup.Name 
                        $Object | Add-Member -MemberType noteproperty -Name 'User' -Value $GroupMember.Name 
                        $Object
                    } 
                } 
            } 
        } 
    } 
}

Get-ElevatedUsers
