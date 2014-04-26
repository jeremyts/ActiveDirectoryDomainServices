<#
  This script will Export and Import the Users

  Syntax examples:
    To export all users:
      ManageUsers.ps1 -Action Export -ReferenceFile UserExport.csv

    To export all users from a particular OU structure:
      ManageUsers.ps1 -Action Export -SearchBase "OU=Users,OU=Corp,DC=mydemosthatrock,DC=com" -ReferenceFile UserExport.csv

    You must use quotes around the SearchBase parameter otherwise the
    comma will be replaced with a space. This is because the comma is a
    special symbol in PowerShell.

    To import from CSV file:
      ManageUsers.ps1 -Action Import -ReferenceFile UserExport.csv

  You could indeed use ldifde, but I find this method provides far more
  flexibility with the manipulation of the data in a simple format.

  All Passwords will be randomly generated.

  There are two variables that can be set as part of the import process:
  1) $Update - Will update the user attributes of accounts that already
     exist with the information specified in the CSV file.
  2) $MoveOU - will move an existing user into the parent OU as specified
     in the CSV file.

  The user attributes we export/import:
  - Name
  - SamAccountName
  - FirstName
  - LastName
  - UserPrincipalName
  - EmailAddress
  - Description
  - DisplayName
  - OUPath
  - userAccountControl
  - CannotChangePassword
  - PasswordNeverExpires
  - PrimaryGroup
  - employeeID
  - employeeType
  - MemberOf

  Release 1.2
  Written by Jeremy@jhouseconsulting.com 13th September 2013
  Modified by Jeremy@jhouseconsulting.com 11th March 2014

#>

#-------------------------------------------------------------
param([String]$Action,[String]$SearchBase,[String]$ReferenceFile)

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

$UsedefaultNamingContext = $False
if ([String]::IsNullOrEmpty($SearchBase)) {
  $UsedefaultNamingContext = $True
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

# Import the Quest ActiveRoles Module
#Add-PSSnapin -Name Quest.ActiveRoles.ADManagement -ErrorAction SilentlyCOntinue -ErrorVariable err
#if ($err){
#    if($err[0].Exception.Message.Contains( 'because it is already added')){
#        Write-Host "Quest.ActiveRoles.ADManagement Snapin already added!" -ForegroundColor green
#    }else{
#        Write-Host "an error occurred:$($err[0])." -BackgroundColor white -ForegroundColor red
#        exit
#    }
#}else{
#    Write-Host "Quest.ActiveRoles.ADManagement Snapin installed" -ForegroundColor green
#}

#-------------------------------------------------------------

$defaultNamingContext = (get-adrootdse).defaultnamingcontext
$DistinguishedName = (Get-ADDomain).DistinguishedName
$DomainName = (Get-ADDomain).NetBIOSName
$DNSRoot = (Get-ADDomain).DNSRoot

If ($UsedefaultNamingContext -eq $True) {
  $SearchBase = $defaultNamingContext
} else {
  $TestSearchBase = Get-ADobject "$SearchBase"
  If ($TestSearchBase -eq $Null) {
    $SearchBase = $defaultNamingContext
  }
}

#-------------------------------------------------------------

function Test-ADPath ()
{
  param([string] $Path)
  <#
    This function was written by Written by VertigoRay found here:
    http://blog.vertigion.com/post/18075217070/powershell-checking-if-ou-exists
  #>
  try {
    if (!([adsi]::Exists("LDAP://$Path"))) {
		Throw('Supplied Path does not exist.')
                return $false
	} else {
		Write-Debug "Path Exists:  $Path"
                return $true
	}
  } catch {
	# If invalid format, error is thrown.
	Throw("Supplied Path is invalid.`n$_")
        return $false
  }
}

function Test-ADObject (
   $objectClass = $(throw "No AD object class specified."),
   $name        = $(throw "No AD object name specified.")
  )
{
  <#
    This function was written by Frank Peter Schultze found here:
    http://www.out-web.net/?p=221
  #>
    switch($objectClass)
    {
        "User"     {$filter = "(&(objectClass=User)(displayName=$name))"}
        "Group"    {$filter = "(&(objectClass=Group)(name=$name))"}
        "Computer" {$filter = "(&(objectClass=Computer)(name=$name))"}
        default    {throw "Unknown objectClass specified."}
    }
    $domainRoot = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().RootDomain.Name
    $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]"GC://$domainRoot")
    $searcher.Filter = $filter
    $result = $searcher.FindOne()
    if($result)
    {
        $result.GetDirectoryEntry()
    }
    else
    {
        $False
    }
}

function Test-XADObject() {
  [CmdletBinding(ConfirmImpact="Low")]
  Param (
    [Parameter(Mandatory=$true,
               Position=0,
               ValueFromPipeline=$true,
               HelpMessage="Identity of the AD object to verify if exists or not."
              )]
    [Object] $Identity
  )
  <#
    This function was written by Jairo Cadena found here:
    http://blogs.msdn.com/b/adpowershell/archive/2009/05/05/how-to-create-a-function-to-validate-the-existence-of-an-ad-object-test-xadobject.aspx
  #>
  trap [Exception] {
    return $false
  }
  $auxObject = Get-ADObject -Identity $Identity
  return $true
}

#-------------------------------------------------------------
# The GET-NewPassword Function. It's called from the CreateUsers process.
# It generates a random password 15 characters long that includes Ascii Upper and Lower case characters with Numbers.

function GET-NewPassword() {

<#
  This function was written by Sean Kearneyon July 3, 2010
  http://www.energizedtech.com/2010/07/powershell-generating-random-p.html

  Delare an array holding what I need.  Here is the format
  The first number is a the number of characters (Ie 26 for the alphabet)
  The Second Number is WHERE it resides in the Ascii Character set
  So 26,97 will pick a random number representing a letter in Asciii
  and add it to 97 to produce the ASCII Character
#>

[int32[]]$ArrayofAscii=26,97,26,65,10,48,15,33

# Complexity can be from 1 - 4 with the results being
# 1 - Pure lowercase Ascii
# 2 - Mix Uppercase and Lowercase Ascii
# 3 - Ascii Upper/Lower with Numbers
# 4 - Ascii Upper/Lower with Numbers and Punctuation
$Complexity=3

# Password Length can be from 1 to as Crazy as you want
$PasswordLength=15

# Nullify the Variable holding the password
$NewPassword=$NULL

# Here is our loop
Foreach ($counter in 1..$PasswordLength) {

# What we do here is pick a random pair (4 possible)
# in the array to generate out random letters / numbers

$pickSet=(GET-Random $complexity)*2

# Pick an Ascii Character and add it to the Password
# Here is the original line I was testing with 
# [char] (GET-RANDOM 26) +97 Which generates
# Random Lowercase ASCII Characters
# [char] (GET-RANDOM 26) +65 Which generates
# Random Uppercase ASCII Characters
# [char] (GET-RANDOM 10) +48 Which generates
# Random Numeric ASCII Characters
# [char] (GET-RANDOM 15) +33 Which generates
# Random Punctuation ASCII Characters

$NewPassword=$NewPassword+[char]((get-random $ArrayOfAscii[$pickset])+$ArrayOfAscii[$pickset+1])
}

# When we're done we Return the $NewPassword 
# BACK to the calling Party
Return $NewPassword

}

#-------------------------------------------------------------
# The GetThePrimaryGroup Function.

function GetThePrimaryGroup{
  param($Username)
  # The current Domain
  $DomainNC = ([ADSI]"LDAP://RootDSE").DefaultNamingContext
  $BaseOU = [ADSI]"LDAP://$DomainNC"
  $LdapFilter = "(&(objectClass=user)(objectCategory=person)(sAMAccountName=$Username))"

  # Find the user
  $Searcher = New-Object DirectoryServices.DirectorySearcher($BaseOU, $LdapFilter)
  $Searcher.PageSize = 1000
  $Searcher.FindAll() | %{
    $User = $_.GetDirectoryEntry()
    $groupID = $user.primaryGroupID
    $arrSID = $user.objectSid.Value
    $SID = New-Object System.Security.Principal.SecurityIdentifier ($arrSID,0)
    $groupSID = $SID.AccountDomainSid.Value + "-" + $user.primaryGroupID.ToString()
    $group = [adsi]("LDAP://<SID=$groupSID>")
    $group.name
  }
}

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
# The AddRemoveGroupMembers Function.

function AddRemoveGroupMembers{
  param($Member,$Groups,$Action)
  # It's important to understand that Users will not show up as a member of a Primary Group. It is only
  # possible to get the primaryGroupToken for the group and then query Active Directory for a list of
  # users with this token.
  $Groups | ForEach-Object {
    $object = Get-ADGroup -LDAPFilter "(sAMAccountName=$_)"
    if ($object -ne $null) {
      Write-Host -ForegroundColor Green "Getting members of the $_ group for verification..."
      # Note that the Get-ADGroupMember cmdlet fails with large groups when using either of the following command lines:
      #   "The size limit of this request was exceeded".
      # $CurrentMembers = Get-ADGroupMember -Identity "$_" | ForEach-Object {$_.samaccountname}
      #   OR
      # $CurrentMembers = Get-ADGroupMember -Identity "$_" | where { $_.samAccountName -eq $Member }
      # Therefore we need to use a differet method for testing group membership.
      # However, testing for group membership of large groups takes too much time, so we add the user the the group and
      # use Try Catch to keep it error free.
      # $GroupObject = [adsi]("LDAP://"+$object.DistinguishedName)
      # $CurrentMembers = $GroupObject.psbase.invoke("Members") | foreach {$_.GetType().InvokeMember("samaccountname",'GetProperty',$null,$_,$null)}
      # Use the following line to output the number of members:
      # $CurrentMembers.count
      If ($Action -eq "Add") {
      #  If ($CurrentMembers -notcontains $Member) {
      #    Write-Host -ForegroundColor Green "Adding $Member to the $_ group..."
      #    Add-ADGroupMember -Identity "$_" -Members "$Member"
      #  } Else {
      #    Write-Host -ForegroundColor Green "$Member is already a member of the $_ group"
      #  }
        try {
          Add-ADGroupMember -Identity "$_" -Members "$Member"
        }
        catch {
          Write-Host -ForegroundColor Green "There was a problem adding $Member to the $_ group. It may already be a member."
        }
      }
      If ($Action -eq "Remove") {
      #  If ($CurrentMembers -contains $Member) {
      #    Write-Host -ForegroundColor Green "Removing $Member from the $_ group..."
      #    Remove-ADGroupMember -Identity "$_" -Members "$Member" -Confirm:$false
      #  } Else {
      #    Write-Host -ForegroundColor Green "$Member is not a member of the $_ group"
      #  }
        try {
          Remove-ADGroupMember -Identity "$_" -Members "$Member" -Confirm:$false
        }
        catch {
          Write-Host -ForegroundColor Green "There was a problem removing $Member from the $_ group. It may not be a member to start with."
        }
      }
    } Else {
      Write-Host -ForegroundColor Green "The $_ group does not exist."
    }
  }
}

#-------------------------------------------------------------#-------------------------------------------------------------
# The CreateUsers Function.

function CreateUsers{
  param ([string]$Name,[string]$SamAccountName,[string]$FirstName,[string]$LastName,[string]$DisplayName,[string]$Description,[string]$OUPath,[string]$UserPrincipalName,[string]$EmailAddress,[int]$userAccountControl,[Boolean]$CannotChangePassword,[Boolean]$PasswordNeverExpires,[Boolean]$ResetPassword,[string]$Password,[Boolean]$ChangePasswordAtLogon,[string]$EmployeeID,[string]$EmployeeType,[string]$PrimaryGroup,[Array]$Memberof,[Boolean]$Update,[Boolean]$MoveOU)

  If ($ChangePasswordAtLogon -eq $True -AND $PasswordNeverExpires -eq $True) {
    # The New-ADUser cmdlet will fail if the PasswordNeverExpires and ChangePasswordAtLogon
    # parameters are both be set to True
    $ChangePasswordAtLogon = $False
  }

  # Set this to the number of seconds you would like to wait for Active Directory replication
  # to complete. If we don't wait, the Get-ADUser cmdlet may fail to find the new user account.
  $SleepTimer = 1

  # Refer to KB305144 for a list of UserAccountControl Flags
  $Enabled = $True
  Switch ($userAccountControl)
  {
    {($userAccountControl -bor 0x0002) -eq $userAccountControl} {
      $Enabled = $False
    }
  }

  If ($FirstName -eq "" -OR $FirstName -eq $NULL) {
    $FirstName = $SamAccountName
  }

  # Creating the user account
  If ( -not (Get-ADUser -LDAPFilter "(sAMAccountName=$SamAccountName)")) {
    Write-Host -ForegroundColor Green "Creating the $SamAccountName user account..."
    New-ADUser -Name $Name -SamAccountName $SamAccountName -Path "$OUPath" -UserPrincipalName $UserPrincipalName -GivenName "$FirstName" -SurName "$LastName" -DisplayName "$DisplayName" -Description "$Description" -EmailAddress $EmailAddress -AccountPassword (ConvertTo-SecureString -AsPlainText "$Password" -Force) -ChangePasswordAtLogon $ChangePasswordAtLogon -CannotChangePassword $CannotChangePassword -PasswordNeverExpires $PasswordNeverExpires -Enabled $Enabled
    $Update = $true
    # Pause to allow the new account to replicate between the domain controllers to avoid errors in the script
    Start-Sleep -s $SleepTimer
  }

  # Updating the user account
  If ($Update -eq $true) {
    Write-Host -ForegroundColor Green "Updating the properties of the $SamAccountName user account..."
    If ($LastName -ne "" -AND $DisplayName -ne "" -AND $Description -ne "") {
      Get-ADUser $SamAccountName | % { Set-ADUser $_ -GivenName "$FirstName" -SurName "$LastName" -DisplayName "$DisplayName" -Description "$Description" -UserPrincipalName $UserPrincipalName -EmailAddress $EmailAddress}
    } elseIf ($LastName -ne "" -AND $DisplayName -ne "" -AND $Description -eq "") {
      Get-ADUser $SamAccountName | % { Set-ADUser $_ -GivenName "$FirstName" -SurName "$LastName" -DisplayName "$DisplayName" -UserPrincipalName $UserPrincipalName -EmailAddress $EmailAddress}
    } elseIf ($LastName -ne "" -AND $DisplayName -eq "" -AND $Description -ne "") {
      Get-ADUser $SamAccountName | % { Set-ADUser $_ -GivenName "$FirstName" -SurName "$LastName" -Description "$Description" -UserPrincipalName $UserPrincipalName -EmailAddress $EmailAddress}
    } elseIf ($LastName -ne "" -AND $DisplayName -eq "" -AND $Description -eq "") {
      Get-ADUser $SamAccountName | % { Set-ADUser $_ -GivenName "$FirstName" -SurName "$LastName" -UserPrincipalName $UserPrincipalName -EmailAddress $EmailAddress}
    } elseIf ($LastName -eq "" -AND $DisplayName -ne "" -AND $Description -ne "") {
      #Get-ADUser $SamAccountName | % { Set-ADUser $_ -GivenName "$FirstName" -DisplayName "$DisplayName" -Description "$Description" -UserPrincipalName $UserPrincipalName -EmailAddress $EmailAddress}
    } elseIf ($LastName -eq "" -AND $DisplayName -ne "" -AND $Description -eq "") {
      Get-ADUser $SamAccountName | % { Set-ADUser $_ -GivenName "$FirstName" -DisplayName "$DisplayName" -UserPrincipalName $UserPrincipalName -EmailAddress $EmailAddress}
    } elseIf ($LastName -eq "" -AND $DisplayName -eq "" -AND $Description -ne "") {
      Get-ADUser $SamAccountName | % { Set-ADUser $_ -GivenName "$FirstName" -Description "$Description" -UserPrincipalName $UserPrincipalName -EmailAddress $EmailAddress}
    } else {
      Get-ADUser $SamAccountName | % { Set-ADUser $_ -GivenName "$FirstName" -UserPrincipalName $UserPrincipalName -EmailAddress $EmailAddress}
    }
    If ($EmployeeID -ne "" -AND $EmployeeType -ne "") {
      Get-ADUser $SamAccountName | % { Set-ADUser $_ -EmployeeID "$EmployeeID" -EmployeeType "$EmployeeType"}
    } elseIf ($EmployeeID -ne "") {
      Get-ADUser $SamAccountName | % { Set-ADUser $_ -EmployeeID "$EmployeeID"}
    } elseIf ($EmployeeType -ne "") {
      Get-ADUser $SamAccountName | % { Set-ADUser $_ -EmployeeID "$EmployeeType"}
    } else {
      # Nothing to set
    }
  }

  # Moving the user account to the correct OU
  If ($MoveOU -eq $true) {
    Write-Host -ForegroundColor Green "Moving the $SamAccountName user account into the '$OUPath' OU..."
    Get-ADUser $SamAccountName | Move-ADObject -TargetPath "$OUPath"
  }

  # Reset the password
  If ($ResetPassword -eq $true) {
    # Note that we need to set the Change Password At Logon after resetting the password.
    Write-Host -ForegroundColor Green "Setting the password for the $SamAccountName user account..."
    Get-ADUser $SamAccountName | % { Set-ADAccountPassword $_ -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$Password" -Force)
                               Set-ADUser $_ -ChangePasswordAtLogon $ChangePasswordAtLogon -CannotChangePassword $CannotChangePassword -PasswordNeverExpires $PasswordNeverExpires -Enabled $Enabled}
    # Note that the Set-ADAccountControl cmdlet can also be used here to help set the UserAccountControl Flags.
  }

  If ($Memberof -ne $NULL -OR $Memberof -ne "") {
    AddRemoveGroupMembers $SamAccountName $Memberof "Add"
  }

  # Changing the Primray Group
  If ($PrimaryGroup -ne $NULL -OR $PrimaryGroup -ne "") {
    $CurrentPrimaryGroup = GetThePrimaryGroup $SamAccountName
    If ($CurrentPrimaryGroup -ne $PrimaryGroup) {
      Write-Host -ForegroundColor Green "Changing the PrimaryGroup of $SamAccountName from $CurrentPrimaryGroup to $PrimaryGroup..."
      $RemoveOldGroup = $True
      ChangeThePrimaryGroup $SamAccountName "$PrimaryGroup" $RemoveOldGroup
    }
  }

}

#-------------------------------------------------------------
If ($Import -eq $true) {

  if ((Test-Path $ReferenceFile) -eq $False) {
    Write-Host -ForegroundColor Red "The $ReferenceFile file is missing.`n"
    exit
  }

  Import-Csv "$ReferenceFile" -Delimiter ';' | foreach-object {

    $Update = $False
    $MoveOU = $False

    write-host -ForegroundColor Green "Importing $($_.SamAccountName)"...

    $OUPath = $_.OUPath -replace ('\|',',')
    $OUPath = $OUPath + "," + $defaultNamingContext

    $UserPrincipalName = $_.UserPrincipalName
    If ($UserPrincipalName -ne $NULL -AND $UserPrincipalName -ne "") {
      $UserPrincipalName = $_.UserPrincipalName+"@"+$DNSRoot
    } else {
      $UserPrincipalName = $_.SamAccountName+"@"+$DNSRoot
    }

    $EmailAddress = $_.EmailAddress
    If ($EmailAddress -ne $NULL -AND $EmailAddress -ne "") {
      $EmailAddress= $_.EmailAddress+"@"+$DNSRoot
    } else {
      $EmailAddress= $_.SamAccountName+"@"+$DNSRoot
    }

    $CannotChangePassword = $_.CannotChangePassword
    if($CannotChangePassword -eq "False") {
      $CannotChangePassword = [System.Convert]::ToBoolean("$False")
    } else {
      $CannotChangePassword = [System.Convert]::ToBoolean("$True")
    }

    $PasswordNeverExpires = $_.PasswordNeverExpires
    if($PasswordNeverExpires -eq "False") {
      $PasswordNeverExpires = [System.Convert]::ToBoolean("$False")
    } else {
      $PasswordNeverExpires = [System.Convert]::ToBoolean("$True")
    }

    $userAccountControl = [int]$_.userAccountControl

    $ChangePasswordAtLogon = $True

    # Generate a random password
    $Password = GET-NewPassword

    # Force a password reset for accounts that are already created.
    $ResetPassword = $False

    If ($_.Memberof -ne "") {
      $Memberof = $_.Memberof.Split("|")
    } Else {
      $Memberof = ""
    }

    # Check if the parent/target exists, and if it's an organizationalunit or container.
    $ParentClass = Get-ADObject -Filter {distinguishedName -eq $OUPath} | % {$_.ObjectClass}
    if($ParentClass -ne $null) {
      # Create the users.
      CreateUsers $_.Name $_.SamAccountName $_.FirstName $_.LastName $_.DisplayName $_.Description $OUPath $UserPrincipalName $EmailAddress $userAccountControl $CannotChangePassword $PasswordNeverExpires $ResetPassword $Password $ChangePasswordAtLogon $_.EmployeeID $_.EmployeeType $_.PrimaryGroup $Memberof $Update $MoveOU
    } Else {
      write-host -ForegroundColor Red "`nThe $OUPath path does not exist.`nThe" $_.GroupName "group cannot be created."
    }
  }
}

#-------------------------------------------------------------
If ($Export -eq $true) {

  function Get-ADParent ([string] $dn) {
    $parts = $dn -split '(?<![\\]),'
    $parts[1..$($parts.Count-1)] -join ','
  }

  $parent = @{Name='Parent'; Expression={ Get-ADParent $_.DistinguishedName } }

  $array = @()

  $UserExclusions = @("")
  $ParentExclusions = @("")
  $IncludeDisabledUsers = $True
  $IncludeUsersWithEmptyEmployeeID = $True
  $EmployeeIDLength = 6

  If ($IncludeDisabledUsers -eq $False -AND $IncludeUsersWithEmptyEmployeeID -eq $False) {
    $filter = "(&(!useraccountcontrol:1.2.840.113556.1.4.803:=2)(employeeID=*))"
  } elseif ($IncludeDisabledUsers -eq $False -AND $IncludeUsersWithEmptyEmployeeID -eq $True) {
    $filter = "(!useraccountcontrol:1.2.840.113556.1.4.803:=2)"
  } elseif ($IncludeDisabledUsers -eq $True -AND $IncludeUsersWithEmptyEmployeeID -eq $False) {
    $filter = "(employeeID=*)"
  } else {
    $filter = ""
  }

  If ($filter -eq "") {
    $Users = Get-ADUser -Filter * -SearchBase $SearchBase -Properties * | Where-Object {!($_.IsCriticalSystemObject) } | select-object Name,SamAccountName,DistinguishedName,$parent,GivenName,Initials,SurName,EmailAddress,UserPrincipalName,Description,DisplayName,Manager,EmployeeID,EmployeeType,CannotChangePassword,PasswordNeverExpires,userAccountControl,MemberOf
  } else {
    $Users = Get-ADUser -LDAPFilter $filter -SearchBase $SearchBase -Properties * | Where-Object {!($_.IsCriticalSystemObject) } | select-object Name,SamAccountName,DistinguishedName,$parent,GivenName,Initials,SurName,EmailAddress,UserPrincipalName,Description,DisplayName,Manager,EmployeeID,EmployeeType,CannotChangePassword,PasswordNeverExpires,userAccountControl,MemberOf
  }

  # Note how we are using Select-Object cmdlet to add the Parent property to the existing "User" object.
  # We could also use the Add-Member cmdlet. But for what we need the Select-Object cmdlet is simpler.

  # Filtering out user that have their IsCriticalSystemObject property set will remove the following users:
  # - Administrator
  # - Guest
  # - krbtgt

  ForEach ($User in $Users) {

    write-host -ForegroundColor Green "Exporting $($User.SamAccountName)"...

    If ($($User.Name).Contains("CNF:") -eq $False) {

      # Refer to KB305144 for a list of UserAccountControl Flags
      $Enabled = $True
      Switch ($User.userAccountControl)
      {
        {($User.userAccountControl -bor 0x0002) -eq $User.userAccountControl} {
          $Enabled = $False
        }
      }

      $IsValidEmployeeIDChars = $False
      If (($User.EmployeeID | Measure-Object -Character).Characters -eq $EmployeeIDLength) {
        $IsValidEmployeeIDChars = $True
      }

      $OUPath = $User.Parent -replace (",$defaultNamingContext","")
      $OUPath = $OUPath -replace (",","|")

      $Memberof = $User.MemberOf 
      $Members = ""
      # Not that you need to use quotes around the identity that you pass to the get-adgroup cmdlet,
      # especially if using a sAMAccountName, as the cmdlet searches the default naming context or
      # partition to find the object. If two or more objects are found, the cmdlet returns a
      # non-terminating error. By using quotes we want are ensuring an exact match.
      $Memberof | %{get-adgroup "$_" |  % {$_.Name}} | ForEach {
      $Member = $_
        If ($Member.Contains("CNF:") -eq $False) {
          If ($Members -ne "" ) {
            $Members += "|" + $Member
          } else {
            $Members += $Member
          }
        } else {
          # Skipping this group as this is a duplication created by a replication collision
        }
      }

      $PrimaryGroup = GetThePrimaryGroup $User.SamAccountName

      Try {
        #$UserPrincipalName = $User.UserPrincipalName -replace ("@$DNSRoot","")
        $UserPrincipalName = $User.UserPrincipalName.Split("@")
        $UserPrincipalName = $UserPrincipalName[0]
        }
      Catch {
        $UserPrincipalName = $User.SamAccountName
        }

      Try {
        #$EmailAddress = $User.EmailAddress -replace ("@$DNSRoot","")
        $EmailAddress = $User.EmailAddress.Split("@")
        $EmailAddress = $EmailAddress[0]
        }
      Catch {
        $EmailAddress = ""
        }

      If ($UserExclusions -notcontains $User.Name -AND $ParentExclusions -notcontains $OUPath) {
        $output = New-Object PSObject
        $output | Add-Member NoteProperty Name $User.Name
        $output | Add-Member NoteProperty SamAccountName $User.SamAccountName
        $output | Add-Member NoteProperty FirstName $User.GivenName
        $output | Add-Member NoteProperty LastName $User.SurName
        $output | Add-Member NoteProperty UserPrincipalName $UserPrincipalName
        $output | Add-Member NoteProperty EmailAddress  $EmailAddress
        $output | Add-Member NoteProperty Description $User.Description
        $output | Add-Member NoteProperty DisplayName $User.DisplayName
        $output | Add-Member NoteProperty OUPath $OUPath
        $output | Add-Member NoteProperty userAccountControl $User.userAccountControl
        $output | Add-Member NoteProperty CannotChangePassword $User.CannotChangePassword
        $output | Add-Member NoteProperty PasswordNeverExpires $User.PasswordNeverExpires
        $output | Add-Member NoteProperty EmployeeID $User.EmployeeID
        $output | Add-Member NoteProperty EmployeeType $User.EmployeeType
        $output | Add-Member NoteProperty PrimaryGroup $PrimaryGroup
        $output | Add-Member NoteProperty MemberOf $Members
        $array += $output
      }

    } else {
      write-host -ForegroundColor Red "- Skipping as user account as this is a duplication created by a replication collision."
    }
  }

  $array | export-csv -notype "$ReferenceFile" -Delimiter ';'

  # Remove the quotes
  (get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii

}
