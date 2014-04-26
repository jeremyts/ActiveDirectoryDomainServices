<#
  This script will set the homeDrive and homeDirectory user
  attributes and create the home drive.

  Syntax:
    .\Provision_Home_Folder.ps1 -Username sAMAccountName -Homedrive H: -HomeDirectory "\\server\share\sAMAccountName"

  Example:
    .\Provision_Home_Folder.ps1 -Username jsaunders -Homedrive H: -HomeDirectory "\\mydemosthatrock.com\userhome\jsaunders"

  Script Name: Provision_Home_Folder.ps1
  Release: 1.2
  Written by Jeremy@jhouseconsulting.com 12th December 2011
  Modified by Jeremy@jhouseconsulting.com 24th April 2014

#>

#-------------------------------------------------------------
param([String]$Username,[String]$HomeDrive,[String]$HomeDirectory)

if ([String]::IsNullOrEmpty($Username)) {
  write-host -ForeGroundColor Red "Username is a required parameter. Exiting Script.`n"
  exit
}

if ([String]::IsNullOrEmpty($HomeDrive)) {
  write-host -ForeGroundColor Red "HomeDrive is a required parameter. Exiting Script.`n"
  exit
}

if ([String]::IsNullOrEmpty($HomeDirectory)) {
  write-host -ForeGroundColor Red "HomeDirectory is a required parameter. Exiting Script.`n"
  exit
}

#-------------------------------------------------------------
# Get the Current Domain Information
$CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
$DNSDomainName = $CurrentDomain.Name
$DomainDistinguishedName = $CurrentDomain.GetDirectoryEntry() | Select-Object -ExpandProperty DistinguishedName

# Find the NetBios Domain Name
$Forest = [system.directoryservices.activedirectory.Forest]::GetCurrentForest() 
$RootDSE = [System.DirectoryServices.DirectoryEntry]([ADSI]"LDAP://RootDSE") 
$ConfigNC = $RootDSE.Get("configurationNamingContext") 
$Searcher = New-Object System.DirectoryServices.DirectorySearcher 
$Searcher.SearchScope = "subtree" 
$Searcher.PropertiesToLoad.Add("nETBIOSName") > $Null 
# Base of search is Partitions container in the configuration container. 
$Searcher.SearchRoot = "LDAP://cn=Partitions,$ConfigNC"
ForEach ($objDomain In $Forest.Domains) {
  $Searcher.Filter = "(nCName=$DomainDistinguishedName)" 
  $NetBIOSDomainName = ($Searcher.FindOne()).Properties.Item("nETBIOSName")[0]
} 

#-------------------------------------------------------------
# The Set-Owner Function. It's called from the CreateHomeDirectory Function.

function Set-Owner {
  <#
  Setting the owner on an ACL in Powershell
  - http://cosmoskey.blogspot.com.au/2010/07/setting-owner-on-acl-in-powershell.html

  Set Owner with PowerShell: "The security identifier is not allowed to be the owner of this object"
  - http://fixingit.wordpress.com/2011/07/08/set-owner-with-powershell-%e2%80%9cthe-security-identifier-is-not-allowed-to-be-the-owner-of-this-object%e2%80%9d/
  #>
 param(
  [System.Security.Principal.IdentityReference]$Principal=$(throw "Mandatory parameter -Principal missing."),
  $File=$(throw "Mandatory parameter -File missing.")
 )
 if(-not (Test-Path $file)){
  throw "File $file is missing."
 }
 if($Principal -eq $null){
  throw "Principal is NULL"
 }

 $code = @"
using System;
using System.Runtime.InteropServices;

namespace CosmosKey.Utils
{
 public class TokenManipulator
 {

  [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
  internal static extern bool AdjustTokenPrivileges(IntPtr htok, bool disall, ref TokPriv1Luid newst, int len, IntPtr prev, IntPtr relen);

  [DllImport("kernel32.dll", ExactSpelling = true)]
  internal static extern IntPtr GetCurrentProcess();

  [DllImport("advapi32.dll", ExactSpelling = true, SetLastError = true)]
  internal static extern bool OpenProcessToken(IntPtr h, int acc, ref IntPtr phtok);

  [DllImport("advapi32.dll", SetLastError = true)]
  internal static extern bool LookupPrivilegeValue(string host, string name, ref long pluid);

  [StructLayout(LayoutKind.Sequential, Pack = 1)]
  internal struct TokPriv1Luid
  {
   public int Count;
   public long Luid;
   public int Attr;
  }

  internal const int SE_PRIVILEGE_DISABLED = 0x00000000;
  internal const int SE_PRIVILEGE_ENABLED = 0x00000002;
  internal const int TOKEN_QUERY = 0x00000008;
  internal const int TOKEN_ADJUST_PRIVILEGES = 0x00000020;

  public const string SE_ASSIGNPRIMARYTOKEN_NAME = "SeAssignPrimaryTokenPrivilege";
  public const string SE_AUDIT_NAME = "SeAuditPrivilege";
  public const string SE_BACKUP_NAME = "SeBackupPrivilege";
  public const string SE_CHANGE_NOTIFY_NAME = "SeChangeNotifyPrivilege";
  public const string SE_CREATE_GLOBAL_NAME = "SeCreateGlobalPrivilege";
  public const string SE_CREATE_PAGEFILE_NAME = "SeCreatePagefilePrivilege";
  public const string SE_CREATE_PERMANENT_NAME = "SeCreatePermanentPrivilege";
  public const string SE_CREATE_SYMBOLIC_LINK_NAME = "SeCreateSymbolicLinkPrivilege";
  public const string SE_CREATE_TOKEN_NAME = "SeCreateTokenPrivilege";
  public const string SE_DEBUG_NAME = "SeDebugPrivilege";
  public const string SE_ENABLE_DELEGATION_NAME = "SeEnableDelegationPrivilege";
  public const string SE_IMPERSONATE_NAME = "SeImpersonatePrivilege";
  public const string SE_INC_BASE_PRIORITY_NAME = "SeIncreaseBasePriorityPrivilege";
  public const string SE_INCREASE_QUOTA_NAME = "SeIncreaseQuotaPrivilege";
  public const string SE_INC_WORKING_SET_NAME = "SeIncreaseWorkingSetPrivilege";
  public const string SE_LOAD_DRIVER_NAME = "SeLoadDriverPrivilege";
  public const string SE_LOCK_MEMORY_NAME = "SeLockMemoryPrivilege";
  public const string SE_MACHINE_ACCOUNT_NAME = "SeMachineAccountPrivilege";
  public const string SE_MANAGE_VOLUME_NAME = "SeManageVolumePrivilege";
  public const string SE_PROF_SINGLE_PROCESS_NAME = "SeProfileSingleProcessPrivilege";
  public const string SE_RELABEL_NAME = "SeRelabelPrivilege";
  public const string SE_REMOTE_SHUTDOWN_NAME = "SeRemoteShutdownPrivilege";
  public const string SE_RESTORE_NAME = "SeRestorePrivilege";
  public const string SE_SECURITY_NAME = "SeSecurityPrivilege";
  public const string SE_SHUTDOWN_NAME = "SeShutdownPrivilege";
  public const string SE_SYNC_AGENT_NAME = "SeSyncAgentPrivilege";
  public const string SE_SYSTEM_ENVIRONMENT_NAME = "SeSystemEnvironmentPrivilege";
  public const string SE_SYSTEM_PROFILE_NAME = "SeSystemProfilePrivilege";
  public const string SE_SYSTEMTIME_NAME = "SeSystemtimePrivilege";
  public const string SE_TAKE_OWNERSHIP_NAME = "SeTakeOwnershipPrivilege";
  public const string SE_TCB_NAME = "SeTcbPrivilege";
  public const string SE_TIME_ZONE_NAME = "SeTimeZonePrivilege";
  public const string SE_TRUSTED_CREDMAN_ACCESS_NAME = "SeTrustedCredManAccessPrivilege";
  public const string SE_UNDOCK_NAME = "SeUndockPrivilege";
  public const string SE_UNSOLICITED_INPUT_NAME = "SeUnsolicitedInputPrivilege";        

  public static bool AddPrivilege(string privilege)
  {
   try
   {
    bool retVal;
    TokPriv1Luid tp;
    IntPtr hproc = GetCurrentProcess();
    IntPtr htok = IntPtr.Zero;
    retVal = OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
    tp.Count = 1;
    tp.Luid = 0;
    tp.Attr = SE_PRIVILEGE_ENABLED;
    retVal = LookupPrivilegeValue(null, privilege, ref tp.Luid);
    retVal = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
    return retVal;
   }
   catch (Exception ex)
   {
    throw ex;
   }

  }
  public static bool RemovePrivilege(string privilege)
  {
   try
   {
    bool retVal;
    TokPriv1Luid tp;
    IntPtr hproc = GetCurrentProcess();
    IntPtr htok = IntPtr.Zero;
    retVal = OpenProcessToken(hproc, TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref htok);
    tp.Count = 1;
    tp.Luid = 0;
    tp.Attr = SE_PRIVILEGE_DISABLED;
    retVal = LookupPrivilegeValue(null, privilege, ref tp.Luid);
    retVal = AdjustTokenPrivileges(htok, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
    return retVal;
   }
   catch (Exception ex)
   {
    throw ex;
   }

  }
 }
}
"@

 $errPref = $ErrorActionPreference
 $ErrorActionPreference= "silentlycontinue"
 $type = [CosmosKey.Utils.TokenManipulator]
 $ErrorActionPreference = $errPref
 if($type -eq $null){
  add-type $code
 }
 $acl = Get-Acl $File
 $acl.psbase.SetOwner($principal)
 [void][CosmosKey.Utils.TokenManipulator]::AddPrivilege([CosmosKey.Utils.TokenManipulator]::SE_RESTORE_NAME)
 set-acl -Path $File -AclObject $acl -passthru
 [void][CosmosKey.Utils.TokenManipulator]::RemovePrivilege([CosmosKey.Utils.TokenManipulator]::SE_RESTORE_NAME)
}

#-------------------------------------------------------------
# The Test-UNC Function. It's called from the CreateHomeDirectory Function.

function Test-UNC {
  <#
  QuickTip: How to validate a UNC path
  - http://blogs.microsoft.co.il/blogs/scriptfanatic/archive/2010/05/27/quicktip-how-to-validate-a-unc-path.aspx
  #>
  param( 
    [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)] 
    [Alias('FullName')] 
    [System.String[]]$Path 
  ) 
  process 
  { 
    foreach($p in $Path) 
    { 
      [bool]([System.Uri]$p).IsUnc    
    } 
  } 
}

#-------------------------------------------------------------
# The CreateHomeDirectory Function.

function CreateHomeDirectory{
  param($Domain,$Username,$HomeDirectory)
    if ((Test-Path -path $HomeDirectory) -eq $False) {
      new-item -ItemType Directory -Path $HomeDirectory
    }

  $NewACL = Get-acl $HomeDirectory

  # Builtin Administrators - Allow - Full Control (Apply onto: This folder, subfolders and files)
  # Local System - Allow - Full Control (Apply onto: This folder, subfolders and files)
  #$NewACL.SetSecurityDescriptorSddlForm("O:BAG:DUD:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)")

  # Domain Admins - Allow - Full Control (Apply onto: This folder, subfolders and files)
  #$Access = $Domain + "\" + "Domain Admins"
  #$objAccess = new-object system.security.principal.NtAccount($Access)
  #$InheritanceFlag = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
  #$PropagationFlag = [system.security.accesscontrol.PropagationFlags]"None"
  #$FolderRights = [System.Security.AccessControl.FileSystemRights]"FullControl"
  #$AccessType = [System.Security.AccessControl.AccessControlType]"Allow"
  #$AccessRule = New-Object system.security.AccessControl.FileSystemAccessRule($objAccess, $FolderRights, $InheritanceFlag, $PropagationFlag, $AccessType)
  #$NewACL.AddAccessRule($AccessRule)

  # File Server Admins - Allow - Full Control (Apply onto: This folder, subfolders and files)
  #$Access = $Domain + "\" + "IS-File-Services-Admins"
  #$objAccess = new-object system.security.principal.NtAccount($Access)
  #$InheritanceFlag = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
  #$PropagationFlag = [system.security.accesscontrol.PropagationFlags]"None"
  #$FolderRights = [System.Security.AccessControl.FileSystemRights]"FullControl"
  #$AccessType = [System.Security.AccessControl.AccessControlType]"Allow"
  #$AccessRule = New-Object system.security.AccessControl.FileSystemAccessRule($objAccess, $FolderRights, $InheritanceFlag, $PropagationFlag, $AccessType)
  #$NewACL.AddAccessRule($AccessRule)

  # Backup Operators - Allow - Full Control (Apply onto: This folder, subfolders and files)
  #$Access = $Domain + "\" + "IS-Backup-Operators"
  #$objAccess = new-object system.security.principal.NtAccount($Access)
  #$InheritanceFlag = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
  #$PropagationFlag = [system.security.accesscontrol.PropagationFlags]"None"
  #$FolderRights = [System.Security.AccessControl.FileSystemRights]"FullControl"
  #$AccessType = [System.Security.AccessControl.AccessControlType]"Allow"
  #$AccessRule = New-Object system.security.AccessControl.FileSystemAccessRule($objAccess, $FolderRights, $InheritanceFlag, $PropagationFlag, $AccessType)
  #$NewACL.AddAccessRule($AccessRule)

  # Home Folder Admins - Allow - Modify (Apply onto: This folder, subfolders and files)
  #$Access = $Domain + "\" + "All-Home-Folder-Admins"
  #$objAccess = new-object system.security.principal.NtAccount($Access)
  #$InheritanceFlag = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
  #$PropagationFlag = [system.security.accesscontrol.PropagationFlags]"None"
  #$FolderRights = [System.Security.AccessControl.FileSystemRights]"Modify"
  #$AccessType = [System.Security.AccessControl.AccessControlType]"Allow"
  #$AccessRule = New-Object system.security.AccessControl.FileSystemAccessRule($objAccess, $FolderRights, $InheritanceFlag, $PropagationFlag, $AccessType)
  #$NewACL.AddAccessRule($AccessRule)

  # Home Folder Admins - Allow - Delete subfolders and files, Change permission, Take Ownership (Apply onto: Subfolders and files only)
  #$Access = $Domain + "\" + "All-Home-Folder-Admins"
  #$objAccess = new-object system.security.principal.NtAccount($Access)
  #$InheritanceFlag = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
  #$PropagationFlag = [system.security.accesscontrol.PropagationFlags]"InheritOnly"
  #$FolderRights = [System.Security.AccessControl.FileSystemRights]"DeleteSubdirectoriesAndFiles, ChangePermissions, TakeOwnership"
  #$AccessType = [System.Security.AccessControl.AccessControlType]"Allow"
  #$AccessRule = New-Object system.security.AccessControl.FileSystemAccessRule($objAccess, $FolderRights, $InheritanceFlag, $PropagationFlag, $AccessType)
  #$NewACL.AddAccessRule($AccessRule)

  Try {
    $IsValidAccount = $True
    $Access = $Domain + "\" + "$Username"
    $objAccess = new-object System.Security.Principal.NTAccount($Access)
    # Validate the user account by retieving its SID
    $SID = $objAccess.Translate([System.Security.Principal.SecurityIdentifier])
  }
  Catch {
    $IsValidAccount = $False
    Write-Host -ForegroundColor Red "The $Username account was not found."
  }

  If ($IsValidAccount -eq $True) {

    # The user account - Allow - Full Control (Apply onto: This folder, subfolders and files)
    $InheritanceFlag = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit"
    $PropagationFlag = [System.Security.AccessControl.PropagationFlags]"None"
    $FolderRights = [System.Security.AccessControl.FileSystemRights]"FullControl"
    $AccessType = [System.Security.AccessControl.AccessControlType]"Allow"
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule ($objAccess, $FolderRights, $InheritanceFlag, $PropagationFlag, $AccessType) 
    $NewACL.AddAccessRule($AccessRule)

    # The user account - Deny - Change permission, Take Ownership (Apply onto: This folder only)
    # This is handy to prevent users from changing permissions and ownership of their own home folder.
    #$InheritanceFlag = [System.Security.AccessControl.InheritanceFlags]"None"
    #$PropagationFlag = [System.Security.AccessControl.PropagationFlags]"None"
    #$FolderRights = [System.Security.AccessControl.FileSystemRights]"ChangePermissions, TakeOwnership"
    #$AccessType = [System.Security.AccessControl.AccessControlType]"Deny"
    #$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule ($objAccess, $FolderRights, $InheritanceFlag, $PropagationFlag, $AccessType) 
    #$NewACL.AddAccessRule($AccessRule)

  } Else {
    
    $Access = "BUILTIN\Administrators"
    $objAccess = new-object System.Security.Principal.NTAccount($Access)
    Write-Host -ForegroundColor DarkYellow "Setting the BUILTIN\Administrators group as owner."
    Write-Host -ForegroundColor DarkYellow "Consider archiving this folder."

  }

  $CurrentACL = Get-acl $HomeDirectory
  $ACLDifferences = compare-object $($CurrentACL.access) $($NewACL.access) -property FileSystemRights,AccessControlType,IdentityReference,InheritanceFlags,PropagationFlags

  # Set the ACL and owner if different
  If ($ACLDifferences -ne $NULL) {
    Write-Host -ForegroundColor Red "The ACL on $HomeDirectory is incorrect."
    if((Test-UNC -Path $HomeDirectory) -eq $True) {$NewACL.SetOwner($objAccess)}
    $Success = $True
    Try {
      Set-Acl -path $HomeDirectory -aclobject $NewACL
    }
    Catch {
      $Success = $False
    }
    if ($Success -eq $True) {
      Write-Host -ForegroundColor Green "Successfully applied permissions and owner to the home folder."
    } else {
      Write-Host -ForegroundColor Red "Error applying permissions and owner to the home folder."
    }
    if((Test-UNC -Path $HomeDirectory) -eq $False) {Set-Owner $objAccess $HomeDirectory}
  } Else {
    Write-Host -ForegroundColor Green "The ACL on $HomeDirectory is correct."
    # Set the owner if different
    If ($CurrentACL.Owner -ne $objAccess.value) {
      Write-Host -ForegroundColor Red "The Owner of $HomeDirectory is incorrect."
      $Success = $False
      if((Test-UNC -Path $HomeDirectory) -eq $True) {
        $CurrentACL.SetOwner($objAccess)
        $Success = $True
        Try {
          Set-Acl -path $HomeDirectory -aclobject $CurrentACL
        }
        Catch {
          $Success = $False
        }
      }
      if((Test-UNC -Path $HomeDirectory) -eq $False) {
        Set-Owner $objAccess $HomeDirectory
        $Success = $True
      }
       if ($Success -eq $True) {
        Write-Host -ForegroundColor Green "Successfully changed the owner on the home folder."
      } else {
        Write-Host -ForegroundColor Red "Error changing the owner on the home folder."
      }
    } else {
      Write-Host -ForegroundColor Green "The owner of $HomeDirectory is correct."
    }
  }
}

#-------------------------------------------------------------

function SetUserAttribute{

  param($Username,$Attribute,$Value)

  # Get the Current Domain Information
  $Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
  $DomainDistinguishedName = $Domain.GetDirectoryEntry() | Select-Object -ExpandProperty DistinguishedName

  $ADScope = "SUBTREE"
  $ADPageSize = 1000
  $ADSearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($DomainDistinguishedName)")

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
      $User.Put($Attribute, $Value)
      $User.SetInfo()
      return $True
      write-host "Updated the $Attribute attribute from $Username account to $Value."
    }
  } else {
      return $False
      write-host "$Username not found."
  }
}

#-------------------------------------------------------------

Write-Host -ForegroundColor Green "Setting the home drive and home directory attributes..."
$Success = SetUserAttribute $Username "homeDrive" $HomeDrive
$Success = SetUserAttribute $Username "homeDirectory" "$HomeDirectory"

If ($Success -eq $True) {
  Write-Host -ForegroundColor Green "Creating the Home Drive..."
  CreateHomeDirectory $NetBIOSDomainName $Username $HomeDirectory
} else {
  Write-Host -ForegroundColor Red "$Username does not exist."
}
