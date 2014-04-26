
#How to Query Individual Properties of the "userAccountControl" Active Directory User property using LDAP
#http://blogs.msdn.com/b/muaddib/archive/2008/10/08/query-individual-properties-of-the-useraccountcontrol-active-directory-user-property.aspx

# The default UserAccountControl value for a typical user is 512

# Some of the more interesting settings are:
# - TRUSTED_FOR_DELEGATION - When this flag is set, the service account (the user or computer account) under which a service runs is trusted for Kerberos delegation. Any such service can impersonate a client requesting the service. To enable a service for Kerberos delegation, you must set this flag on the userAccountControl property of the service account.
# - NOT_DELEGATED - When this flag is set, the security context of the user is not delegated to a service even if the service account is set as trusted for Kerberos delegation.
# - USE_DES_KEY_ONLY - (Windows 2000/Windows Server 2003) Restrict this principal to use only Data Encryption Standard (DES) encryption types for keys.
# - DONT_REQUIRE_PREAUTH - (Windows 2000/Windows Server 2003) This account does not require Kerberos pre-authentication for logging on.
# - PASSWORD_EXPIRED - (Windows 2000/Windows Server 2003) The user's password has expired. 
# - TRUSTED_TO_AUTH_FOR_DELEGATION - (Windows 2000/Windows Server 2003) The account is enabled for delegation. This is a security-sensitive setting. Accounts that have this option enabled should be tightly controlled. This setting lets a service that runs under the account assume a client's identity and authenticate as that user to other remote servers on the network. 

# References:
# - http://support.microsoft.com/kb/305144
# - http://support.microsoft.com/kb/269181

Import-Module ActiveDirectory

# Use DES encryption types for this account
#Get-ADUser -LdapFilter "(&(objectclass=user)(objectcategory=user)(useraccountcontrol:1.2.840.113556.1.4.803:=2097152))" | Format-Table Name, DistinguishedName


# Do not require Kerberos preauthentication
#Get-ADUser -LdapFilter "(&(objectclass=user)(objectcategory=user)(useraccountcontrol:1.2.840.113556.1.4.803:=4194304))" | Format-Table Name, DistinguishedName


# Password never expires
#Get-ADUser -LdapFilter "(&(objectclass=user)(objectcategory=user)(userAccountControl:1.2.840.113556.1.4.803:=65536)" | Format-Table Name, DistinguishedName


# Account Expires


# Account is sensitive and cannot be delegated


# Account is disabled
#$DisabledAccounts = Get-ADUser -LdapFilter "(&(objectclass=user)(objectcategory=user)(userAccountControl:1.2.840.113556.1.4.803:=2)
#$DisabledAccounts.Count
