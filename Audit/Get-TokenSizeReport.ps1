<#
  This script will enumerate all user accounts in a Domain, calculate their estimated Token Size and create
  a report of the top x users in CSV format.

  Script Name: Get-TokenSizeReport.ps1
  Release 1.8
  Modified by Jeremy@jhouseconsulting.com 31/12/2013
  Written by Jeremy@jhouseconsulting.com 02/12/2013

  Original script was derived from the CheckMaxTokenSize.ps1 written by Tim Springston [MS] on 7/19/2013
  http://gallery.technet.microsoft.com/scriptcenter/Check-for-MaxTokenSize-520e51e5

  Re-wrote the script to be more efficient and provide a report for all users in the
  Domain.

  References:
  - Microsoft KB327825: Problems with Kerberos authentication when a user belongs to many groups
    http://support.microsoft.com/kb/327825
  - Microsoft KB243330: Well-known security identifiers in Windows operating systems
    http://support.microsoft.com/kb/243330
  - Microsoft KB328889: Users who are members of more than 1,015 groups may fail logon authentication
    http://support.microsoft.com/kb/328889
  - Microsoft KB938118: How to use Group Policy to add the MaxTokenSize registry entry to multiple computers
    http://support.microsoft.com/kb/938118
  - Microsoft Blog: Managing Token Bloat:
    http://blogs.dirteam.com/blogs/sanderberkouwer/archive/2013/05/22/common-challenges-when-managing-active-directory-domain-services-part-2-unnecessary-complexity-and-token-bloat.aspx

  Although it's not documented in KB327825 or any other Microsoft references, I also add the number of
  global groups outside the user's account domain that the user is a member of to the "d" calculation of
  the TokenSize. Whilst the Microsoft methodology is to add universal groups from other domains, it is
  possible to add global groups too. Therefore it's important to capture this and correctly include it in
  the calculation.

  For users with large tokens consider reducing direct and transitive (nested) group memberships.
  Larger environments that have evolved over time also have a tendancy to suffer from Circular Group
  Nesting and sIDHistory.

  On the odd ocasion I was receiving the following error:
  - Exception calling "FindByIdentity" with "2" argument(s): "Multiple principals contain
    a matching Identity."
  - There seemed to be a known bug in .NET 4.x when passing two arguments to the FindByIdentity() method.
  - The solution was to either use a machine with .NET 3.5 or re-write the script to pass
    three arguments as per the Get-UserPrincipal function provided in the following Scripting Guy article:
    - http://blogs.technet.com/b/heyscriptingguy/archive/2009/10/08/hey-scripting-guy-october-8-2009.aspx
    This function passes the Context Type, FQDN Domain Name and Parent OU/Container.
  - Other references:
    - http://richardspowershellblog.wordpress.com/2008/05/27/account-management-member-of/
    - http://www.powergui.org/thread.jspa?threadID=20194

  I have also seen the following error:
  - Exception calling "GetAuthorizationGroups" with "0" argument(s): "An error (1301) occurred while
    enumerating the groups. The group's SID could not be resolved."
  - Other references:
    - http://richardspowershellblog.wordpress.com/2008/05/27/account-management-member-of/
    - https://groups.google.com/forum/#!topic/microsoft.public.adsi.general/jX3wGd0JPOo
    - http://lucidcode.com/2013/02/18/foreign-security-groups-in-active-directory/

  Added the tokenGroups attribute to get all nested groups as I could not achieve 100% reliability using
  the GetAuthorizationGroups() method. Could not afford for it to start failing after running for hours
  in large environments.
  - References:
    - http://www.msxfaq.de/code/tokengroup.htm
    - http://www.msxfaq.de/tools/dumpticketsize.htm

  There are important differences between using the GetAuthorizationGroups() method versus the tokenGroups
  attribute that need to be understood. Aside from the unreliability of GetAuthorizationGroups(), when push
  comes to shove you get different results depending on which method you use, and what you want to achieve.
    - The tokenGroups attribute only contains the actual "Active Directory" principals, which are groups and
      siDHistory.
    - However, tokenGroups does not reveal cross-forest/domain group memberships. The tokenGroups attribute
      is constructed by Active Directory on request, and this depends on the availability of a Global Catalog
      server: http://msdn.microsoft.com/en-us/library/ms680275(VS.85).aspx
    - The GetAuthorizationGroups() method also returns the well-known security identifiers of the local
      system (LSALogonUser) for the user running the script, which will include groups such as:
      - Everyone (S-1-1-0)
      - Authenticated Users (S-1-5-11)
      - This Organization (S-1-5-15)
      - Low Mandatory Level (S-1-16-4096)
      This will vary depending on where you're running the script from and in what user context. The result
      is still consistent, as it adds the same overhead to each user. But this is misleading.
    - GetAuthorizationGroups() will return cross-forest/domain group memberships, but cannot resolve them
      because they contain a ForeignSecurityPrincipal. It therefore fails as documented above.
    - GetAuthorizationGroups() does not contain siDHistory.

  In my view you would use the tokenGroups attribute to collate a consistent and accurate user report across
  the environment, whereas the GetAuthorizationGroups() method could be used in a logon script to calucate
  the token of the user together with the system they are logging on to. The actual calculation of the token
  size adds the estimated value for ticket overhead anyway, hence the reason why using the tokenGroups
  attribute provides a consistent result for all users.
  If you wanted an accurate token size per user per system and GetAuthorizationGroups() method continues to
  prove to be unreliable, you could use the tokenGroups attribute together with the addition of the output
  from the "whoami /groups" command to get all the well-known groups and label needed to calculate the
  complete local token.

  Microsoft also has a tool called Tokensz.exe that could also be used in a logon script. It can be downloded
  from here: http://www.microsoft.com/download/en/details.aspx?id=1448

  To be completed:
  - Some further research and testing needs to be completed with the code that retrieves the tokenGroups
    attribute to validate performance between the GetInfoEx method or RefreshCache method.
  - Work out how to report on cross-forest/domain group memberships as neither tokenGroups or
    GetAuthorizationGroups() can achieve this.
#>

#-------------------------------------------------------------

# Set this value to the number of users with large tokens that
# you want to report on.
$TopUsers = 200

# Set this to the size in bytes that you want to capture the user
# information for the report.
$TokensSizeThreshold = 6000
# Note that if we don't set a threshold high enough, then in large
# environments the $array will grow too large and may then create
# memory issues.

# Set this value to true if you want to see the progress bar.
$ProgressBar = $True

# Set this value to true if you want to output to the console
$ConsoleOutput = $True

# Set this value to true if you want a summary output to the
# console when the script has completed.
$OutputSummary = $True

# Set this value to true to use the tokenGroups attribute
$UseTokenGroups = $True

# Set this value to true to use the GetAuthorizationGroups() method
$UseGetAuthorizationGroups = $False

#-------------------------------------------------------------

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}
$ReferenceFile = $(&$ScriptPath) + "\KerberosTokenSizeReport.csv"

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

Function Test-DotNetFrameWork35
{
 Test-path -path 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5'
} #end Test-DotNetFrameWork35

If(-not(Test-DotNetFrameWork35)) { "Requires .NET Framework 3.5" ; exit }

$array = @()
$TotalUsersProcessed = 0
$UserCount = 0
$GroupCount = 0
$LargestTokenSize = 0
$TotalGoodTokens = 0
$TotalTokensBetween8and12K = 0
$TotalLargeTokens = 0
$TotalVeryLargeTokens = 0

$ADRoot = ([System.DirectoryServices.DirectoryEntry]"LDAP://RootDSE")
$DefaultNamingContext = $ADRoot.defaultNamingContext

# Derive FQDN Domain Name
$TempDefaultNamingContext = $DefaultNamingContext.ToString().ToUpper()
$DomainName = $TempDefaultNamingContext.Replace(",DC=",".")
$DomainName = $DomainName.Replace("DC=","")

# Create an LDAP search for all enabled users not marked as criticalsystemobjects to avoid system accounts
$ADFilter = "(&(objectClass=user)(objectcategory=person)(!userAccountControl:1.2.840.113556.1.4.803:=2)(!(isCriticalSystemObject=TRUE))(!name=IUSR*)(!name=IWAM*)(!name=ASPNET))"
# There is a known bug in PowerShell requiring the DirectorySearcher
# properties to be in lower case for reliability.
$ADPropertyList = @("distinguishedname","samaccountname","useraccountcontrol","objectsid","sidhistory","primarygroupid","lastlogontimestamp","memberof")
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
$Users = $ADSearcher.Findall()
$UserCount = $users.Count

if ($UserCount -ne 0) {
  foreach($user in $users) {
    #$user.Properties
    #$user.Properties.propertynames
    $lastLogonTimeStamp = ""
    $lastLogon = ""
    $UserDN = $user.Properties.distinguishedname[0]
    $samAccountName = $user.Properties.samaccountname[0]
    If (($user.Properties.lastlogontimestamp | Measure-Object).Count -gt 0) {
      $lastLogonTimeStamp = $user.Properties.lastlogontimestamp[0]
      $lastLogon = [System.DateTime]::FromFileTime($lastLogonTimeStamp)
      if ($lastLogon -match "1/01/1601") {$lastLogon = "Never logged on before"}
    } else {
      $lastLogon = "Never logged on before"
    }
    $OU = $user.GetDirectoryEntry().Parent
    $OU = $OU -replace ("LDAP:\/\/","")

    # Get user SID
    $arruserSID = New-Object System.Security.Principal.SecurityIdentifier($user.Properties.objectsid[0], 0)
    $userSID = $arruserSID.Value

    # Get the SID of the Domain the account is in
    $AccountDomainSid = $arruserSID.AccountDomainSid.Value

    # Get User Account Control & Primary Group by binding to the user account
    $objUser = [ADSI]("LDAP://" + $UserDN)
    $UACValue = $objUser.useraccountcontrol[0]
    $primarygroupID = $objUser.PrimaryGroupID
    # Primary group can be calculated by merging the account domain SID and primary group ID
    $primarygroupSID = $AccountDomainSid + "-" + $primarygroupID.ToString()
    $primarygroup = [adsi]("LDAP://<SID=$primarygroupSID>")
    $primarygroupname = $primarygroup.name
    $objUser = $null

    # Get SID history
    $SIDCounter = 0
    if ($user.Properties.sidhistory -ne $null) {
      foreach ($sidhistory in $user.Properties.sidhistory) {
        $SIDHistObj = New-Object System.Security.Principal.SecurityIdentifier($sidhistory, 0)
        #Write-Host -ForegroundColor green $SIDHistObj.Value "is in the SIDHistory."
        $SIDCounter++
      }
    }
    $SIDHistObj = $null

    $TotalUsersProcessed ++
    If ($ProgressBar) {
      Write-Progress -Activity 'Processing Users' -Status ("Username: {0}" -f $samAccountName) -PercentComplete (($TotalUsersProcessed/$UserCount)*100)
    }

    # Use TokenGroups Attribute
    If ($UseTokenGroups) {
      $UserAccount = [ADSI]"$($User.Path)"
      $UserAccount.GetInfoEx(@("tokenGroups"),0) | Out-Null
      $ErrorActionPreference = "continue"
      $error.Clear()
      $groups = $UserAccount.GetEx("tokengroups")
      if ($Error) {
        Write-Warning "  Tokengroups not readable"
        $Groups=@()   #empty enumeration
      }
      $GroupCount = 0
      # Note that the tokengroups includes all principals, which includes siDHistory, so we need
      # to subtract the sIDHistory count to correctly report on the number of groups in the token.
      $GroupCount = $groups.count - $SIDCounter

      $SecurityDomainLocalScope = 0
      $SecurityGlobalInternalScope = 0
      $SecurityGlobalExternalScope = 0
      $SecurityUniversalInternalScope = 0
      $SecurityUniversalExternalScope = 0

      foreach($token in $groups) {
        $principal = New-Object System.Security.Principal.SecurityIdentifier($token,0)
        $GroupSid = $principal.value
        #$group = $principal.Translate([System.Security.Principal.NTAccount])
        #$group.value
        $grp = [ADSI]"LDAP://<SID=$GroupSid>"
        if ($grp.Path -ne $null) {
          $grpdn = $grp.distinguishedName.tostring().ToLower()
          $grouptype = $grp.groupType.psbase.value

          switch -exact ($GroupType) {
            "-2147483646"   { 
                            # Global security scope 
                            if ($GroupSid -match $DomainSID) 
                            {
                              $SecurityGlobalInternalScope++
                            } else { 
                              # Global groups from others.
                              $SecurityGlobalExternalScope++
                            } 
                            } 
            "-2147483644"   { 
                            # Domain Local scope 
                            $SecurityDomainLocalScope++
                            } 
            "-2147483643"   { 
                            # Domain Local BuildIn scope
                            $SecurityDomainLocalScope++
                            }
            "-2147483640"   { 
                            # Universal security scope 
                            if ($GroupSid -match $AccountDomainSid) 
                            { 
                              $SecurityUniversalInternalScope++ 
                            } else { 
                              # Universal groups from others.
                              $SecurityUniversalExternalScope++ 
                            } 
                            } 
          }
        }
      } 
    }

    # Use GetAuthorizationGroups() Method
    If ($UseGetAuthorizationGroups) {

      $userPrincipal = Get-UserPrincipal -userName $SamAccountName -cName $DomainName -cContainer "$OU"

      $GroupCount = 0
      $SecurityDomainLocalScope = 0
      $SecurityGlobalInternalScope = 0
      $SecurityGlobalExternalScope = 0
      $SecurityUniversalInternalScope = 0
      $SecurityUniversalExternalScope = 0

      # Use GetAuthorizationGroups() for Indirect Group MemberShip, which includes all Nested groups and the Primary group
      Try {
        $groups = $userPrincipal.GetAuthorizationGroups() | select SamAccountName, GroupScope, SID
        
        $GroupCount = $groups.count

        foreach ($group in $groups) {
          $GroupSid = $group.SID.value
          #$group

          switch ($group.GroupScope)
            {
              "Local" {
                # Domain Local & Domain Local BuildIn scope
                $SecurityDomainLocalScope++
                }
              "Global" {
                # Global security scope 
                if ($GroupSid -match $DomainSID) {
                  $SecurityGlobalInternalScope++
                } else { 
                  # Global groups from others.
                  $SecurityGlobalExternalScope++
                }
              }
                "Universal" {
                # Universal security scope 
                if ($GroupSid -match $AccountDomainSid) {
                  $SecurityUniversalInternalScope++
                } else {
                  # Universal groups from others.
                  $SecurityUniversalExternalScope++
                }
              }
            }
        }
      }
      Catch {
        write-host "Error with the GetAuthorizationGroups() method: $($_.Exception.Message)" -ForegroundColor Red
      }
    }

    If ($ConsoleOutput) {
      Write-Host -ForegroundColor green "Checking the token of user $SamAccountName in domain $DomainName"
      Write-Host -ForegroundColor green "There are $GroupCount groups in the token."
      Write-Host -ForegroundColor green "- $SecurityDomainLocalScope are domain local security groups."
      Write-Host -ForegroundColor green "- $SecurityGlobalInternalScope are domain global scope security groups inside the users domain."
      Write-Host -ForegroundColor green "- $SecurityGlobalExternalScope are domain global scope security groups outside the users domain."
      Write-Host -ForegroundColor green "- $SecurityUniversalInternalScope are universal security groups inside the users domain."
      Write-Host -ForegroundColor green "- $SecurityUniversalExternalScope are universal security groups outside the users domain."
      Write-host -ForegroundColor green "The primary group is $primarygroupname."
      Write-host -ForegroundColor green "There are $SIDCounter SIDs in the users SIDHistory."
      Write-Host -ForegroundColor green "The current userAccountControl value is $UACValue."
    }

    $TrustedforDelegation = $false
    if ((($UACValue -bor 0x80000) -eq $UACValue) -OR (($UACValue -bor 0x1000000) -eq $UACValue)) {
      $TrustedforDelegation = $true
    }

    # Calculate the current token size, taking into account whether or not the account is trusted for delegation or not.
    $TokenSize = 1200 + (40 * ($SecurityDomainLocalScope + $SecurityGlobalExternalScope + $SecurityUniversalExternalScope + $SIDCounter)) + (8 * ($SecurityGlobalInternalScope  + $SecurityUniversalInternalScope))
    if ($TrustedforDelegation -eq $false) {
      If ($ConsoleOutput) {
        Write-Host -ForegroundColor green "Token size is $Tokensize and the user is not trusted for delegation."
      }
    } else {
      $TokenSize = 2 * $TokenSize
      If ($ConsoleOutput) {
        Write-Host -ForegroundColor green "Token size is $Tokensize and the user is trusted for delegation."
      }
    }

    If ($TokenSize -le 12000) {
      $TotalGoodTokens ++
      If ($TokenSize -gt 8192) {
        $TotalTokensBetween8and12K ++
      }
    } elseIf ($TokenSize -le 48000) {
      $TotalLargeTokens ++
    } else {
      $TotalVeryLargeTokens ++
    }

    If ($TokenSize -gt $LargestTokenSize) {
      $LargestTokenSize = $TokenSize
      $LargestTokenUser = $SamAccountName
    }

    If ($TokenSize -ge $TokensSizeThreshold) {
      $obj = New-Object -TypeName PSObject
      $obj | Add-Member -MemberType NoteProperty -Name "Domain" -value $DomainName
      $obj | Add-Member -MemberType NoteProperty -Name "SamAccountName" -value $SamAccountName
      $obj | Add-Member -MemberType NoteProperty -Name "TokenSize" -value $TokenSize
      $obj | Add-Member -MemberType NoteProperty -Name "Memberships" -value $GroupCount
      $obj | Add-Member -MemberType NoteProperty -Name "DomainLocal" -value $SecurityDomainLocalScope
      $obj | Add-Member -MemberType NoteProperty -Name "GlobalInternal" -value $SecurityGlobalInternalScope
      $obj | Add-Member -MemberType NoteProperty -Name "GlobalExternal" -value $SecurityGlobalExternalScope
      $obj | Add-Member -MemberType NoteProperty -Name "UniversalInternal" -value $SecurityUniversalInternalScope
      $obj | Add-Member -MemberType NoteProperty -Name "UniversalExternal" -value $SecurityUniversalExternalScope
      $obj | Add-Member -MemberType NoteProperty -Name "SIDHistory" -value $SIDCounter
      $obj | Add-Member -MemberType NoteProperty -Name "UACValue" -value $UACValue
      $obj | Add-Member -MemberType NoteProperty -Name "TrustedforDelegation" -value $TrustedforDelegation
      $obj | Add-Member -MemberType NoteProperty -Name "LastLogon" -value $lastLogon
      $array += $obj
    }

    If ($ConsoleOutput) {
      $percent = "{0:P}" -f ($TotalUsersProcessed/$UserCount)
      write-host -ForegroundColor green "Processed $TotalUsersProcessed of $UserCount user accounts = $percent complete."
      Write-host " "
    }
  }

  If ($OutputSummary) {
    Write-Host -ForegroundColor green "Summary:"
    Write-Host -ForegroundColor green "- Processed $UserCount user accounts."
    Write-Host -ForegroundColor green "- $TotalGoodTokens have a calculated token size of less than or equal to 12000 bytes."
    If ($TotalGoodTokens -gt 0) {
      Write-Host -ForegroundColor green "  - These users are good."
    }
    If ($TotalTokensBetween8and12K -gt 0) {
      Write-Host -ForegroundColor green "  - Although $TotalTokensBetween8and12K of these user accounts have tokens above 8K and should therefore be reviewed."
    }
    Write-Host -ForegroundColor green "- $TotalLargeTokens have a calculated token size larger than 12000 bytes."
    If ($TotalLargeTokens -gt 0) {
      Write-Host -ForegroundColor green "  - These users will be okay if you have increased the MaxTokenSize to 48000 bytes.`n  - Consider reducing direct and transitive (nested) group memberships."
    }
    Write-Host -ForegroundColor red "- $TotalVeryLargeTokens have a calculated token size larger than 48000 bytes."
    If ($TotalVeryLargeTokens -gt 0) {
      Write-Host -ForegroundColor red "  - These users will have problems. Do NOT increase the MaxTokenSize beyond 48000 bytes.`n  - Reduce the direct and transitive (nested) group memberships."
    }
    Write-Host -ForegroundColor green "- $LargestTokenUser has the largest calculated token size of $LargestTokenSize bytes in the $DomainName domain."
  }

  # Write-Output $array | Format-Table
  $array | Sort-Object TokenSize -descending | select-object -first $TopUsers | export-csv -notype -path "$ReferenceFile" -Delimiter ';'

  # Remove the quotes
  (get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii
}
