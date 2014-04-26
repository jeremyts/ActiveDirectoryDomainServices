<#
  This script will Export and Import the Groups

  Syntax examples:
    Export:
      ManageGroups.ps1 -Action Export -ReferenceFile GroupExport.csv

    Import:
      ManageGroups.ps1 -Action Import -ReferenceFile GroupExport.csv

  You could indeed use ldifde, but I find this method provides far more
  flexibility with the manipulation of the data in a simple format.

  IMPORTANT: You may need to run this script twice to ensure that all
             groups are first created and then added to other groups as
             members.

  Release 1.2
  Written by Jeremy@jhouseconsulting.com 13th September 2013
  Modified by Jeremy@jhouseconsulting.com 4th February 2014
#>

#-------------------------------------------------------------
param([String]$Action,[String]$ReferenceFile)

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

#-------------------------------------------------------------
# The AddMembers Function.

function AddMembers{
  param($Group,$Members)
  $object = Get-ADGroup -LDAPFilter "(sAMAccountName=$Group)"
  if ($object -ne $null -AND $object -ne "") {
    # Note that the Get-ADGroupMember cmdlet fails with large groups: "The size limit of this request was exceeded".
    # $CurrentMembers = Get-ADGroupMember -Identity "$_" | ForEach-Object {$_.samaccountname}
    # Therefore we need to use a differet method for testing group membership.
    $GroupObject = [adsi]("LDAP://"+$object.DistinguishedName)
    $CurrentMembers = $GroupObject.psbase.invoke("Members") | foreach {$_.GetType().InvokeMember("samaccountname",'GetProperty',$null,$_,$null)}
    # Use the following line to output the number of members:
    # $CurrentMembers.count
    $Members | ForEach-Object {
      if ($_ -ne $null -AND $_ -ne "") {
        If (Get-ADObject -LDAPFilter "(sAMAccountName=$_)") {
          If($CurrentMembers -notcontains "$_") {
            Write-Host -ForegroundColor Green "Adding $_ to the $Group group..."
            Add-ADGroupMember -Identity "$Group" -Members "$_"
          } else {
            Write-Host -ForegroundColor Green "$_ is already a member of the $Group group"
          }
        } else {
          Write-Host -ForegroundColor Red "The $_ object does not exist."
        }
      }
    }
  }
}


#-------------------------------------------------------------
# The CreateGroup Function.

function CreateGroup{
  param($GroupName,$DisplayName,$Description,$Scope,$Category,$OUPath,$ManagedBy,$Members)
  $EnableManagerCanUpdateMembershipList = $True
  Write-Host -ForegroundColor Green "`nChecking for the $GroupName group"
  $object = Get-ADGroup -LDAPFilter "(sAMAccountName=$GroupName)"
    if ($object -eq $null) {
      Write-Host -ForegroundColor Green "Creating the group"
      #If ($ManagedBy -eq $NULL -OR $ManagedBy -eq "") {
        New-ADGroup -name "$GroupName" -path "$OUPath" -DisplayName "$DisplayName" -Description "$Description" -groupScope $Scope -groupCategory $Category
      #} else {
      #  New-ADGroup -name "$GroupName" -path "$OUPath" -DisplayName "$DisplayName" -Description "$Description" -groupScope $Scope -groupCategory $Category -managedby "$ManagedBy"
      #}
    }
    else {
      Write-Host -ForegroundColor Green "The group already exists"
      Write-Host -ForegroundColor Green "Updating the properties of the $GroupName group..."
      If ($DisplayName -ne $NULL -AND $DisplayName -ne "" ) {
        Get-ADGroup $GroupName | % { Set-ADGroup $_ -DisplayName "$DisplayName" }
      }
      If ($Description -ne $NULL -AND $Description -ne "" ) {
        Get-ADGroup $GroupName | % { Set-ADGroup $_ -Description "$Description" }
      }
      If ($ManagedBy -ne $NULL -AND $ManagedBy -ne "") {
        $ManagedBy = Get-ADObject -LDAPFilter "(sAMAccountName=$ManagedBy)" | % {$_.Name}
        If ($ManagedBy -ne $NULL) {
          Write-Host -ForegroundColor Green "Updating the ManagedBy attribute."
          Get-ADGroup $GroupName | % { Set-ADGroup $_ -managedby "$ManagedBy" }
          If ($EnableManagerCanUpdateMembershipList -eq $true) {
            # Set the "Manager can update membership list" check box
            #Get-QADGroup -Identity "$GroupName" | Add-QADPermission -Account "$ManagedBy" -Rights WriteProperty -Property ('member') -ApplyTo ThisObjectOnly | Out-Null
          }
        }
      }
    }
  if ($Members -ne "" -AND $Members -ne $null) {
    AddMembers "$GroupName" $Members
  }
}

#-------------------------------------------------------------
If ($Import -eq $true) {

  $HoldingOU = "OU=Temp,OU=Security Groups,OU=Corporate" + "," + $defaultNamingContext

  if ((Test-Path $ReferenceFile) -eq $False) {
    Write-Host -ForegroundColor Red "The $ReferenceFile file is missing.`n"
    exit
  }

  Import-Csv "$ReferenceFile" -Delimiter ';' | foreach-object {

    $OUPath = $_.OUPath -replace ('\|',',')
    $OUPath = $OUPath + "," + $defaultNamingContext

    If ($_.Members -ne "") {
      $Members = $_.Members.Split("|")
    } Else {
      $Members = ""
    }

    # Check if the parent/target exists, and if it's an organizationalunit or container.
    $ParentClass = Get-ADObject -Filter {distinguishedName -eq $OUPath} | % {$_.ObjectClass}

    if($ParentClass -eq $null) {
      $OUPath = $HoldingOU
      Try {
        # Check if the target OU exists. If not, create it.
        $ExistingOU = Get-ADObject -Filter {distinguishedName -eq $OUPath}
        }
      Catch {
        $ExistingOU = $NULL
        }
    }
 
    if($ParentClass -ne $null -OR $ExistingOU -ne $null) {
      # Create the group.
      CreateGroup $_.GroupName $_.DisplayName $_.Description $_.Scope $_.Category $OUPath $_.ManagedBy $Members
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

  $GroupExclusions = @("DnsAdmins","DnsUpdateProxy")
  $ParentExclusions = @("CN=Builtin")

  $Groups = Get-ADGroup -Filter * -SearchBase $defaultNamingContext -Properties * | Where-Object {!($_.IsCriticalSystemObject) } | select-object Name,SamAccountName,DistinguishedName,$parent,Description,DisplayName,GroupCategory,GroupScope,ManagedBy
  # Note how we are using Select-Object cmdlet to add the Parent property to the existing "Group" object.
  # We could also use the Add-Member cmdlet. But for what we need the Select-Object cmdlet is simpler.

  # Filtering out groups that have their IsCriticalSystemObject property set will remove the following groups:
  # - Administrators
  # - Users
  # - Guests
  # - Print Operators
  # - Backup Operators
  # - Replicator
  # - Remote Desktop Users
  # - Network Configuration Operators
  # - Performance Monitor Users
  # - Performance Log Users
  # - Distributed COM Users
  # - IIS_IUSRS
  # - Cryptographic Operators
  # - Event Log Readers
  # - Certificate Service DCOM Access
  # - Domain Computers
  # - Domain Controllers
  # - Schema Admins
  # - Enterprise Admins
  # - Cert Publishers
  # - Domain Admins
  # - Domain Users
  # - Domain Guests
  # - Group Policy Creator Owners
  # - RAS and IAS Servers
  # - Server Operators
  # - Account Operators
  # - Pre-Windows 2000 Compatible Access
  # - Incoming Forest Trust Builders
  # - Windows Authorization Access Group
  # - Terminal Server License Servers
  # - Allowed RODC Password Replication Group
  # - Denied RODC Password Replication Group
  # - Read-only Domain Controllers
  # - Enterprise Read-only Domain Controllerss
  # This is typically all groups from under the Builtin container, and some groups from the Users container
  # with the exclusion of the DnsAdmins and DnsUpdateProxy groups.

  ForEach ($Group in $Groups) {

    write-host -ForegroundColor Green "Exporting $($Group.SamAccountName)"...

    If ($($Group.Name).Contains("CNF:") -eq $False) {

      $OUPath = $Group.Parent -replace (",$defaultNamingContext","")
      $OUPath = $OUPath -replace (",","|")

      If ($GroupExclusions -notcontains $Group.Name -AND $ParentExclusions -notcontains $OUPath) {

        If ($Group.ManagedBy -ne $NULL -AND $Group.ManagedBy -ne "") {
          $ManagedBy = $Group.ManagedBy
          $ManagedBy = Get-ADObject -Filter {distinguishedName -eq $ManagedBy} | % {$_.Name}
        }

        # Get Members
        $Members = ""
        # When using the Get-ADGroupMember cmdlet to get group members it will
        # fail to get the group members if the group contains a member that is
        # a foreign security principal (i.e. members from another Domain) where
        # the SID cannot be resolved. To work around this we use the member
        # property of the Get-ADGroup cmdlet instead, which will ignore any
        # foreign security principals.
        $GetMembers = (Get-ADGroup -identity $Group.DistinguishedName -properties member).member | Get-ADObject -Properties Name,sAMAccountName | Select Name,SamAccountName | ForEach {
        #$GetMembers = Get-ADGroupMember -identity $Group.DistinguishedName |Select Name,SamAccountName | ForEach {
          $Member = $_.Name
          If ($Member.Contains("CNF:") -eq $False) {
            $Member = $_.SamAccountName 
            If ($Members -ne "" ) {
              If ($Member -ne "" -OR $Member -ne $NULL) {
                $Members += "|" + $Member
              }
            } else {
              $Members += $_.SamAccountName
            }
          } else {
            # Skipping this group as this is a duplication created by a replication collision
          }
        }

        $output = New-Object PSObject
        $output | Add-Member NoteProperty GroupName $Group.Name
        $output | Add-Member NoteProperty Description $Group.Description
        $output | Add-Member NoteProperty DisplayName $Group.DisplayName
        $output | Add-Member NoteProperty Scope $Group.GroupScope
        $output | Add-Member NoteProperty Category $Group.GroupCategory
        $output | Add-Member NoteProperty OUPath $OUPath
        $output | Add-Member NoteProperty ManagedBy $ManagedBy
        $output | Add-Member NoteProperty Members $Members
        $array += $output

      }

    } else {
      write-host -ForegroundColor Red "- Skipping as this is a duplication created by a replication collision."
    }
  }

  $array | export-csv -notype "$ReferenceFile" -Delimiter ';'

  # Remove the quotes
  (get-content "$ReferenceFile") |% {$_ -replace '"',""} | out-file "$ReferenceFile" -Fo -En ascii

}
