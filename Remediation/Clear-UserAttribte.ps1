
function ClearUserAttribute{

  param($Username,$Attribute)

  [int]$ADS_PROPERTY_CLEAR = 1
  [int]$ADS_PROPERTY_UPDATE = 2
  [int]$ADS_PROPERTY_APPEND = 3
  [int]$ADS_PROPERTY_DELETE = 4

  # The current Domain
  $ADRoot = ([System.DirectoryServices.DirectoryEntry]"LDAP://RootDSE")
  $DefaultNamingContext = $ADRoot.defaultNamingContext

  $ADScope = "SUBTREE"
  $ADPageSize = 1000
  $ADSearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($DefaultNamingContext)")

  # Find the account that will be effected by the change
  $ADFilter = "(&(objectClass=user)(objectCategory=person)(sAMAccountName=$Username))"
  $ADPropertyList = @()
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
      $User.PutEx($ADS_PROPERTY_CLEAR, $Attribute, $null)
      $User.SetInfo()
      write-host "Cleared the $Attribute attribute from $Username account."
    }
  } else {
      write-host "$Username not found."
  }
}


ClearUserAttribute $sAMAccountName $Attribute
