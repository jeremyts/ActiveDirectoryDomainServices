@echo off
cls
Echo Running various Active Directory reports...
Echo.

SetLocal

Echo.
Echo Active Directory Sites and Subnets Report
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ActiveDirectorySitesandSubnetsReport.ps1"'"

Echo.
Echo Check Active Directory Sites
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0CheckActiveDirectorySites.ps1"'"

Echo.
Echo Find Orphaned GPOs
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0FindOrphanedGPOs.ps1"'"

Echo.
Echo Generate OU Permission Report
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0OU_Permissions.ps1"'"

Echo.
Echo Circular Nested Groups Report
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0CircularNestedGroups.ps1"'"

::Echo.
::Echo Get Hotfixes Installed Report
::powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0Get-HotFixReport.ps1"'"

Echo.
Echo Generate BPA Reports
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0GenerateBPAReports.ps1"'"

:Finish
EndLocal
Exit /b 0
