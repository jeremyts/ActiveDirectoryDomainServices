<#

  This script will get the Active Directory Tombstone Period.

  Notes:
  - The tombstonelifetime period will default to 60 days if missing.
  - From Windows Server 2003 with Service Pack 1 (SP1), the default tombstonelifetime value is set
    to 180 days to increase shelf-life of backups and allow longer disconnection times.
  - The tombstonelifetime attribute and value is only set on a new forest build and not an upgrade,
    so if it is missing, the forest probably started its life as a Windows 2000 or Windows 2003
    pre-SP1 domain. However, there was also a known bug the CD2 of the Windows 2003 R2 media which
    was missing the tombstonelifetime attribute, even though the CD1 media is Windows 2003 SP1. As
    this is only set for a new build, upgrades to the forest/domain over the years does not fix this.
 -  Microsoft increased this back in Q1 2005 to help Companies avoid issues with management and
    backups of Active Directory. You don't necessarily need to set it to 180 days. It is preferable
    to adjust the value to what is appropriate for the Company's backup and recovery strategy.
    However, the value should at least be present to remove any confusion for future assessments and
    any functionality that may leverage the attribute in the future.

  References:
  - http://blog.joeware.net/2006/07/21/476/
  - http://blog.joeware.net/2006/07/23/484/

#>

Import-Module ActiveDirectory

$defaultNamingContext = (get-adrootdse).defaultnamingcontext

$Days = (get-adobject "cn=Directory Service,cn=Windows NT,cn=Services,cn=Configuration,$defaultNamingContext" -properties "tombstonelifetime").tombstonelifetime

If ($Days -eq "" -OR $Days -eq $NULL)
{
  $Days = "missing so will default to 60"
}

write-host -ForegroundColor green "The Active Directory Tombstone Period is $Days days."
