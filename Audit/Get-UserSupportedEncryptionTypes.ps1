
Import-Module ActiveDirectory

# Windows Server 2008 and above
# This account supports Kerberos AES 128 bit encryption
# This account supports Kerberos AES 256 bit encryption
# msDS-SupportedEncryptionTypes
# CRC (KERB_ENCTYPE_DES_CBC_CRC, 0x00000001): Supports CRC32
# MD5 (KERB_ENCTYPE_DES_CBC_MD5, 0x00000002): Supports RSA-MD5
# RC4 (KERB_ENCTYPE_RC4_HMAC_MD5, 0x00000004): Supports RC4-HMAC-MD5
# A128 (KERB_ENCTYPE_AES128_CTS_HMAC_SHA1_96, 0x00000008): Supports HMAC-SHA1-96-AES128
# A256 (KERB_ENCTYPE_AES256_CTS_HMAC_SHA1_96, 0x00000010): Supports HMAC-SHA1-96-AES256


function resolve-EncryptionTypes {            
 param (            
  [int]$key            
 )
  switch ($key)
    {
      "1" {$SupportedEncryptionTypes = @("DES_CRC")}
      "2" {$SupportedEncryptionTypes = @("DES_MD5")}
      "3" {$SupportedEncryptionTypes = @("DES_CRC","DES_MD5")}
      "4" {$SupportedEncryptionTypes = @("RC4")}
      "8" {$SupportedEncryptionTypes = @("AES128")}
      "16" {$SupportedEncryptionTypes = @("AES256")}
      "24" {$SupportedEncryptionTypes = @("AES128","AES256")}
      "28" {$SupportedEncryptionTypes = @("RC4","AES128","AES256")}
      "31" {$SupportedEncryptionTypes = @("DES_CRC","DES_MD5","RC4","AES128","AES256")}
      default {$SupportedEncryptionTypes = @("Undefined value of $key")}
    }
  $SupportedEncryptionTypes            
}
$Users = Get-ADUser -Properties * -LdapFilter "(&(objectclass=user)(objectcategory=user)(msDS-SupportedEncryptionTypes=*)(!msDS-SupportedEncryptionTypes=0))" | Select-Object Name, @{N="EncryptionTypes"; E={resolve-EncryptionTypes $($_."msDS-SupportedEncryptionTypes")}}
ForEach ($User in $Users) {
  $User.Name
  ForEach ($EncryptionType in $User.EncryptionTypes) {
    $EncryptionType
  }
}
