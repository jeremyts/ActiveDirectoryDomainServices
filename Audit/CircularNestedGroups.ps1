# CircularNestedGroups.ps1
# PowerShell program to find any instances of circular nested groups.
# Author: Richard Mueller
# PowerShell Version 1.0
# July 31, 2011

Function CheckNesting ($Group, $Parents)
{
    # Recursive function to enumerate group members of a group.
    # $Group is the group whose membership is being evaluated.
    # $Parents is an array of all parent groups of $Group.
    # $Count is the number of groups involved in circular nesting.
    # $GroupMembers is the hash table of all groups and their group members.
    # $Count and $GroupMembers must have script scope.
    # If any group member matches any of the parents, we have
    # detected an instance of circular nesting.

    # Enumerate all group members of $Group.
    ForEach ($Member In $Script:GroupMembers[$Group])
    {
        # Check if this group matches any parent group.
        ForEach ($Parent In $Parents)
        {
            If ($Member -eq $Parent)
            {
                Write-Host "Circular Nested Group: $Parent"
                $Script:Count = $Script:Count + 1
                # Avoid infinite loop.
                Return
            }
        }
        # Check all group members for group membership.
        If ($Script:GroupMembers.ContainsKey($Member))
        {
            # Add this member to array of parent groups.
            # However, this is not a parent for siblings.
            # Recursively call function to find nested groups.
            $Temp = $Parents
            CheckNesting $Member ($Temp += $Member)
        }
    }
}

# Hash table of groups and their direct group members.
$GroupMembers = @{}

# Search entire domain.
$Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$Root = $Domain.GetDirectoryEntry()
$Searcher = [System.DirectoryServices.DirectorySearcher]$Root

$Searcher.PageSize = 200
$Searcher.SearchScope = "subtree"
$Searcher.PropertiesToLoad.Add("distinguishedName") > $Null
$Searcher.PropertiesToLoad.Add("member") > $Null

# Filter on all group objects.
$Searcher.Filter = "(objectCategory=group)"
$Results = $Searcher.FindAll()

# Enumerate groups and populate Hash table. The key value will be
# the Distinguished Name of the group. The item value will be an array
# of the Distinguished Names of all members of the group that are groups.
# The item value starts out as an empty array, since we don't know yet
# which members are groups.
ForEach ($Group In $Results)
{
    $DN = [string]$Group.properties.Item("distinguishedName")
    $Script:GroupMembers.Add($DN, @())
}

# Enumerate the groups again to populate the item value arrays.
# Now we can check each member to see if it is a group.
ForEach ($Group In $Results)
{
    $DN = [string]$Group.properties.Item("distinguishedName")
    $Members = @($Group.properties.Item("member"))
    # Enumerate the members of the group.
    ForEach ($Member In $Members)
    {
        # Check if the member is a group.
        If ($Script:GroupMembers.ContainsKey($Member))
        {
            # Add the Distinguished Name of this member to the item value array.
            $Script:GroupMembers[$DN] += $Member
        }
    }
}

# Count the number of circular nested groups found.
$Script:Count = 0
# Retrieve array of all groups in the domain.
$Groups = $Script:GroupMembers.Keys
# Enumerate all groups and check group membership of each.
ForEach ($Group In $Groups)
{
    # Check group membership for circular nesting.
    CheckNesting $Group @($Group)
}

Write-Host "Number of circular nested groups found = $Script:Count"