#-------------------------------------------------------------
# The ChangeThePrimaryGroup Function. It's called from the CreateUsers Function.

function ChangeThePrimaryGroup{
  # Original script found here:
  # - http://www.indented.co.uk/index.php/2010/01/22/changing-the-primary-group-with-powershell/
  # Modified to be more flexible.
  param($Username,$NewPrimaryGroup,$RemoveOldGroup)

  # The Primary Group Token for Domain Users and Guests will always be
  # the same value (no matter the forest). Used as a demonstration of
  # how the value can be retrieved

  # The current Domain
  $ADRoot = ([System.DirectoryServices.DirectoryEntry]"LDAP://RootDSE")
  $DefaultNamingContext = $ADRoot.defaultNamingContext

  $ADScope = "SUBTREE"
  $ADPageSize = 1000
  $ADSearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($DefaultNamingContext)")

  # Find the New Group
  $ADFilter = "(&(ObjectCategory=group)(sAMAccountName=$NewPrimaryGroup))"
  $ADPropertyList = @("distinguishedname","samaccountname","primarygrouptoken")
  $ADSearcher = New-Object System.DirectoryServices.DirectorySearcher 
  $ADSearcher.SearchRoot = $ADSearchRoot
  $ADSearcher.PageSize = $ADPageSize 
  $ADSearcher.Filter = $ADFilter 
  $ADSearcher.SearchScope = $ADScope
  if ($ADPropertyList) {
    foreach ($ADProperty in $ADPropertyList) {
      [Void]$ADSearcher.PropertiesToLoad.Add($ADProperty)
    }
  }
  $results = $ADSearcher.FindAll()
  $Count = $results.Count
  if ($Count -ne 0) {
    foreach($result in $results) {
      $Group = $result.GetDirectoryEntry()
      $GroupDN = $result.Properties.distinguishedname[0]
      $objGroup = [ADSI]("LDAP://" + $GroupDN)
      $objGroup.GetInfoEx(@("primaryGroupToken"), 0)
      $GroupToken = $objGroup.Get("primaryGroupToken")
    }
  }
  $ADSearcher = $NULL

  # Find the account that will be effected by the change
  $ADFilter = "(&(objectClass=user)(objectCategory=person)(sAMAccountName=$Username))"
  $ADPropertyList = @("distinguishedname","samaccountname","objectsid","primarygroupid")
  $ADSearcher = New-Object System.DirectoryServices.DirectorySearcher 
  $ADSearcher.SearchRoot = $ADSearchRoot
  $ADSearcher.PageSize = $ADPageSize 
  $ADSearcher.Filter = $ADFilter 
  $ADSearcher.SearchScope = $ADScope
  if ($ADPropertyList) {
    foreach ($ADProperty in $ADPropertyList) {
      [Void]$ADSearcher.PropertiesToLoad.Add($ADProperty)
    }
  }
  $results = $ADSearcher.FindAll()
  $Count = $results.Count
  if ($Count -ne 0) {
    foreach($result in $results) {
      $User = $result.GetDirectoryEntry()
      $UserDN = $result.Properties.distinguishedname[0]

      # Get user SID
      $arruserSID = New-Object System.Security.Principal.SecurityIdentifier($user.Properties.objectsid[0], 0)
      $userSID = $arruserSID.Value

      # Get the SID of the Domain the account is in
      $AccountDomainSid = $arruserSID.AccountDomainSid.Value

      # Get Primary Group by binding to the user account
      $objUser = [ADSI]("LDAP://" + $UserDN)
      $ExistingprimarygroupID = $objUser.PrimaryGroupID
      # Primary group can be calculated by merging the account domain SID and primary group ID
      $ExistingprimarygroupSID = $AccountDomainSid + "-" + $ExistingprimarygroupID.ToString()
      $Existingprimarygroup = [adsi]("LDAP://<SID=$ExistingprimarygroupSID>")
      $Existingprimarygroupname = $Existingprimarygroup.name
      $objUser = $null

      If ($Existingprimarygroupname -ne $NewPrimaryGroup) {

        # The user must be a member of the group first
        Try {
          $Group.Add($User.AdsPath)
        }
        Catch {
          if ($error[0].exception -like ("*The object already exists.*")) {
            #The user is already a member of the group
          }
        }

        # Change the Primary Group
        $User.Put("primaryGroupId", $GroupToken)
        $User.SetInfo()

        If ($RemoveOldGroup) {
          # Remove the old group
          $Existingprimarygroup.Remove($User.AdsPath)
        }
      }
    }
  }
  $ADSearcher = $NULL
}

#-------------------------------------------------------------

$SamAccountName = "jesaunders"
$NewPrimaryGroup = "Domain Guests"
$RemoveOldGroup = $True

ChangeThePrimaryGroup $SamAccountName "$NewPrimaryGroup" $RemoveOldGroup

