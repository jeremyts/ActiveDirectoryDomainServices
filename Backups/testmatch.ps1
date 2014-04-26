#http://ss64.com/ps/syntax-regex.html

$String1 = "JEAN_CLAUDE"
$String2 = "VAN DAMME"

#Write-host "Match any word character in '$String1'"
# \w This is equivalent to [a-zA-Z_0-9]
#$String1 -match '\w+'

Write-host "Match any non-word character in '$String1'"
# \W This is equivalent to [^a-zA-Z_0-9]
$String1 -match '\W+'

Write-host "Match any non-letter in '$String1'"
$String1 -match '[^a-zA-Z]'

Write-host "Match any white-space in '$String1'"
# \s This is equivalent to [ \f\n\r\t\v]
$String1 -match '\s+'

#Write-host "Match any word character in '$String2'"
# \w This is equivalent to [a-zA-Z_0-9]
#$String2 -match '\w+'

Write-host "Match any non-word character in '$String2'"
# \W This is equivalent to [^a-zA-Z_0-9]
$String2 -match '\W+'

Write-host "Match any non-letter in '$String2'"
$String2 -match '[^a-zA-Z]'

Write-host "Match any white-space in '$String2'"
# \s This is equivalent to [ \f\n\r\t\v]
$String2 -match '\s+'
