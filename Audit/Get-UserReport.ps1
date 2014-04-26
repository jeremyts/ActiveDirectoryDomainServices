<#
  This script will enumerate all user accounts in a Domain, and report
  on the following attributes:
  - samAccountName
  - givenName
  - sn
  - initials
  - mail
  - telephoneNumber
  - mobile
  - displayName
  - description
  - title
  - company
  - physicalDeliveryOfficeName
  - employeeID
  - employeeType
  - msexchextensioncustomattribute1
  - primaryGroupID
  - userAccountControl
  - objectsid
  - accountExpires
  - lastlogontimestamp
  - whencreated
  - memberOf

  We derive the Enabled and PasswordExpired boolean value from the
  userAccountControl value.

  The IsStale logic is…
  - If it was created more than 90 days ago…
    - Mark it as a stale account if it's never been logged on before.
    - Mark it as a stale account if it hasn't been logged on in 90 days.
  - If it expired more than 30 days ago, mark it as a stale account.

  We also check to see if the samAccountName is lowercase and the number
  of characters it contains, as the account length can be a maximum of
  12 characters for SAP.

  We also check to see if the Surname (sn) contains a non-alpha character.

  IMPORTANT: I am using the -Append parameter of the Export-Csv cmdlet,
             which is ONLY support from PowerShell v3 and above. If using
             v2, you'll need to download and add the following function to
             your profile to make Export-CSV cmdlet handle -Append parameter
             http://dmitrysotnikov.wordpress.com/2010/01/19/export-csv-append/

  Script Name: Get-UserReport.ps1
  Release 1.4
  Written by Jeremy@jhouseconsulting.com 27/12/2013
  Modified by Jeremy@jhouseconsulting.com 15/04/2014

#>

#-------------------------------------------------------------

# Set this value to true if you want to see the progress bar.
$ProgressBar = $True

# Set this to true to process disabled user accounts.
$ProcessDisabledUsers = $False

# Set this to true to include extra user attributes such as
# displayName, description, telephoneNumber, mobile, title,
# company, physicalDeliveryOfficeName
$ExtendedDetails = $False

# Set this to true to include the user's direct group membership
$IncludeMemberOf = $False
#-------------------------------------------------------------

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$ReferenceFile = $(&$ScriptPath) + "\UserReport.csv"

if (Test-Path -path $ReferenceFile) {
  remove-item $ReferenceFile -force -confirm:$false
}

$TotalUsersProcessed = 0
$UserCount = 0

$ADRoot = ([System.DirectoryServices.DirectoryEntry]"LDAP://RootDSE")
$DefaultNamingContext = $ADRoot.defaultNamingContext

# Derive FQDN Domain Name
$TempDefaultNamingContext = $DefaultNamingContext.ToString().ToUpper()
$DomainName = $TempDefaultNamingContext.Replace(",DC=",".")
$DomainName = $DomainName.Replace("DC=","")

If ($ProcessDisabledUsers -eq $False) {
  # Create an LDAP search for all enabled users not marked as criticalsystemobjects to avoid system accounts
  $ADFilter = "(&(objectClass=user)(objectcategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2)(!(isCriticalSystemObject=TRUE))(!name=IUSR*)(!name=IWAM*)(!name=ASPNET))"
} else {
  # Create an LDAP search for all users not marked as criticalsystemobjects to avoid system accounts
  $ADFilter = "(&(objectClass=user)(objectcategory=person)(!(isCriticalSystemObject=TRUE))(!name=IUSR*)(!name=IWAM*)(!name=ASPNET))"
}
# There is a known bug in PowerShell requiring the DirectorySearcher
# properties to be in lower case for reliability.
$ADPropertyList = @("distinguishedname","samaccountname","givenname","sn","initials","mail","telephonenumber","mobile","displayname","description","title","company","physicaldeliveryofficename","employeeid","employeetype","useraccountcontrol","objectsid","primarygroupid","lastlogontimestamp","whencreated","accountexpires","msexchextensioncustomattribute1","memberof")
$ADScope = "SUBTREE"
$ADPageSize = 1000
$ADSearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($DefaultNamingContext)") 
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
$colResults = $ADSearcher.Findall()
$UserCount = $colResults.Count

if ($UserCount -ne 0) {
  foreach($objResult in $colResults) {
    $lastLogonTimeStamp = ""
    $lastLogon = ""
    $IsStale = $False
    $UserDN = $objResult.Properties.distinguishedname[0]
    $samAccountName = $objResult.Properties.samaccountname[0]
    if ($samAccountName -cmatch "^[^A-Z]*$") {
      $IssamAccountNameLowerCase = $True
    } else {
      $IssamAccountNameLowerCase = $False
    }
    $samAccountNameLength = ($samAccountName | Measure-Object -Character).Characters
    If (($objResult.Properties.givenname | Measure-Object).Count -gt 0) {
      $Firstname = $objResult.Properties.givenname[0]
    } else {
      $Firstname = ""
    }
    If (($objResult.Properties.sn | Measure-Object).Count -gt 0) {
      $Surname = $objResult.Properties.sn[0]
      $nonalphacharsinsurname = $Surname -match '[^a-zA-Z]'
    } else {
      $Surname = ""
      $nonalphacharsinsurname = $False
    }
    # Get user SID
    $arruserSID = New-Object System.Security.Principal.SecurityIdentifier($objResult.Properties.objectsid[0], 0)
    $userSID = $arruserSID.Value
    If (($objResult.Properties.initials | Measure-Object).Count -gt 0) {
      $Initials = $objResult.Properties.initials[0]
    } else {
      $Initials = ""
    }
    If (($objResult.Properties.employeeid | Measure-Object).Count -gt 0) {
      $EmployeeID = $objResult.Properties.employeeid[0]
    } else {
      $EmployeeID = ""
    }
    If (($objResult.Properties.employeetype | Measure-Object).Count -gt 0) {
      $EmployeeType = $objResult.Properties.employeetype[0]
    } else {
      $EmployeeType = ""
    }
    If (($objResult.Properties.mail | Measure-Object).Count -gt 0) {
      $EMail = $objResult.Properties.mail[0]
    } else {
      $EMail = ""
    }
    If ($ExtendedDetails) {
      If (($objResult.Properties.displayname | Measure-Object).Count -gt 0) {
        $DisplayName = $objResult.Properties.displayname[0]
      } else {
        $DisplayName = ""
      }
      If (($objResult.Properties.description | Measure-Object).Count -gt 0) {
        $Description = $objResult.Properties.description[0]
      } else {
        $Description = ""
      }
      If (($objResult.Properties.telephonenumber | Measure-Object).Count -gt 0) {
        $TelephoneNumber = $objResult.Properties.telephonenumber[0]
      } else {
        $TelephoneNumber = ""
      }
      If (($objResult.Properties.mobile | Measure-Object).Count -gt 0) {
        $Mobile = $objResult.Properties.mobile[0]
      } else {
        $Mobile = ""
      }
      If (($objResult.Properties.title | Measure-Object).Count -gt 0) {
        $Title = $objResult.Properties.title[0]
      } else {
        $Title = ""
      }
      If (($objResult.Properties.company | Measure-Object).Count -gt 0) {
        $Company = $objResult.Properties.company[0]
      } else {
        $Company = ""
      }
      If (($objResult.Properties.physicaldeliveryofficename | Measure-Object).Count -gt 0) {
        $Office = $objResult.Properties.physicaldeliveryofficename[0]
      } else {
        $Office = ""
      }
    }
    If (($objResult.Properties.msexchextensioncustomattribute1| Measure-Object).Count -gt 0) {
      $msExchExtensionCustomAttribute1 = $objResult.Properties.msexchextensioncustomattribute1[0]
    } else {
      $msExchExtensionCustomAttribute1 = ""
    }
    If (($objResult.Properties.lastlogontimestamp | Measure-Object).Count -gt 0) {
      $lastLogonTimeStamp = $objResult.Properties.lastlogontimestamp[0]
      $lastLogon = [System.DateTime]::FromFileTime($lastLogonTimeStamp)
      if ($lastLogon -match "1/01/1601") {$lastLogon = "Never logged on before"}
    } else {
      $lastLogon = "Never logged on before"
    }

    $whencreated = $objResult.Properties.whencreated[0]

    # If it was created more than 90 days ago...
    # - Mark it as a stale account if it's never been logged on before.
    # - Mark it as a stale account if it hasn't been logged on in 90 days.
    If ($whencreated -le (Get-Date).AddDays(-90)) {
      If ($lastLogon -eq "Never logged on before") {
        $IsStale = $True
      } elseif ($lastLogon -le (Get-Date).AddDays(-90)) {
        $IsStale = $True
      }
    }

    $AE = $objResult.Properties.accountexpires
    If (($AE.Item(0) -eq 0) -or ($AE.Item(0) -gt [DateTime]::MaxValue.Ticks)) {
      $accountExpires = "Never"
    } else {
      $AEDate = [DateTime]$AE.Item(0)
      $accountExpires = $AEDate.AddYears(1600).ToLocalTime()
      # Mark it as a stale account if it expired more than 30 days ago.
      If ($accountExpires -le (Get-Date).AddDays(-30)) {
        $IsStale = $True
      }
    }

    # Get user SID
    $arruserSID = New-Object System.Security.Principal.SecurityIdentifier($objResult.Properties.objectsid[0], 0)
    $userSID = $arruserSID.Value

    # Get the SID of the Domain the account is in
    $AccountDomainSid = $arruserSID.AccountDomainSid.Value

    # Get User Account Control & Primary Group by binding to the user account
    $objUser = [ADSI]("LDAP://" + $UserDN)
    If (($objUser.useraccountcontrol | Measure-Object).Count -gt 0) {
      $UACValue = $objUser.useraccountcontrol[0]
    } else {
      $UACValue = ""
    }
    $primarygroupID = $objUser.primarygroupid
    If ($primarygroupID -ne $NULL) {
      # Primary group can be calculated by merging the account domain SID and primary group ID
      $primarygroupSID = $AccountDomainSid + "-" + $primarygroupID.ToString()
      $primarygroup = [adsi]("LDAP://<SID=$primarygroupSID>")
      $primarygroupname = $primarygroup.name[0]
      $objUser = $null
    } else {
      $primarygroupname = "NULL"
    }

    $Enabled = $True
    $PasswordExpired = $False
    switch ($UACValue)
    {
      {($UACValue -bor 0x0002) -eq $UACValue} {
        $Enabled = $False
      }
      {($UACValue -bor 0x800000) -eq $UACValue} {
        $PasswordExpired = $True
      }
    }

    If ($IncludeMemberOf) {
      $Members = ""
      $groups = $objResult.Properties.memberof | ForEach {
        $groupDN = $_
        $objGroup = [ADSI]("LDAP://" + $groupDN)
        $Member = $objGroup.samaccountname
          If ($Members -ne "" ) {
            $Members += "|" + $Member
          } else {
            $Members += $Member
          }
          $objGroup = $null
        }
    }

    $obj = New-Object -TypeName PSObject
    $obj | Add-Member -MemberType NoteProperty -Name "Username" -value $SamAccountName
    $obj | Add-Member -MemberType NoteProperty -Name "IsNameLowerCase" -value $IssamAccountNameLowerCase
    $obj | Add-Member -MemberType NoteProperty -Name "LengthOfName" -value $samAccountNameLength
    $obj | Add-Member -MemberType NoteProperty -Name "Firstname" -value $Firstname
    $obj | Add-Member -MemberType NoteProperty -Name "Surname" -value $Surname
    $obj | Add-Member -MemberType NoteProperty -Name "NonAlphaCharsInSurname" -value $nonalphacharsinsurname
    $obj | Add-Member -MemberType NoteProperty -Name "Initials" -value $Initials
    $obj | Add-Member -MemberType NoteProperty -Name "EmployeeID" -value $EmployeeID
    $obj | Add-Member -MemberType NoteProperty -Name "EmployeeType" -value $EmployeeType
    $obj | Add-Member -MemberType NoteProperty -Name "EMail" -value $EMail
    If ($ExtendedDetails) {
      $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -value $DisplayName
      $obj | Add-Member -MemberType NoteProperty -Name "Description" -value $Description
      $obj | Add-Member -MemberType NoteProperty -Name "TelephoneNumber" -value $TelephoneNumber
      $obj | Add-Member -MemberType NoteProperty -Name "Mobile" -value $Mobile
      $obj | Add-Member -MemberType NoteProperty -Name "Title" -value $Title
      $obj | Add-Member -MemberType NoteProperty -Name "Company" -value $Company
      $obj | Add-Member -MemberType NoteProperty -Name "Office" -value $Office
    }
    $obj | Add-Member -MemberType NoteProperty -Name "msExchExtensionCustomAttribute1" -value $msExchExtensionCustomAttribute1
    $obj | Add-Member -MemberType NoteProperty -Name "PrimaryGroup" -value $primarygroupname
    $obj | Add-Member -MemberType NoteProperty -Name "Enabled" -value $Enabled
    $obj | Add-Member -MemberType NoteProperty -Name "PasswordExpired" -value $PasswordExpired
    $obj | Add-Member -MemberType NoteProperty -Name "IsStale" -value $IsStale
    $obj | Add-Member -MemberType NoteProperty -Name "Expires" -value $accountExpires
    $obj | Add-Member -MemberType NoteProperty -Name "LastLogon" -value $lastLogon
    $obj | Add-Member -MemberType NoteProperty -Name "Created" -value $whencreated
    $obj | Add-Member -MemberType NoteProperty -Name "ObjectSID" -value $userSID
    If ($IncludeMemberOf) {
      $obj | Add-Member -MemberType NoteProperty -Name "MemberOf" -value $Members
    }

    # Write-Output $array | Format-Table
    $obj | Export-Csv -Path "$ReferenceFile" -Append -Delimiter ';' -NoTypeInformation -Encoding ASCII

    $TotalUsersProcessed ++
    If ($ProgressBar) {
      Write-Progress -Activity 'Processing Users' -Status ("Username: {0}" -f $samAccountName) -PercentComplete (($TotalUsersProcessed/$UserCount)*100)
    }

  }

  # Remove the quotes from the output file.
  (get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii
}
