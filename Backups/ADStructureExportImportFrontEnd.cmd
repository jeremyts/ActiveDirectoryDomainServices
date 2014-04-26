@echo off

SetLocal

If /I "%1"=="BackupTasks" GOTO BackupTasks
If /I "%1"=="Backup" GOTO BackupTasks
If /I "%1"=="RestoreTasks" GOTO RestoreTasks
If /I "%1"=="Restore" GOTO RestoreTasks

:Selection
cls
ECHO.
Echo.
Echo ------------------------------------------
Echo ** AD Structure Export/Import Front-End **
Echo ------------------------------------------
Echo.
ECHO You are currently logged into the %userdnsdomain% Domain.
Echo.
ECHO Which set of tasks do you want to run?
ECHO.
ECHO 	1. Backup/Export Tasks
Echo         - Export OU Structure
Echo         - Export Groups
Echo         - Export Users
Echo.
ECHO 	2. Restore/Import Tasks
Echo         - Import OU Structure
Echo         - Import Groups
Echo         - Import Users
Echo         - Import Groups (We run this script twice to ensure
Echo                          that all groups are created and
Echo                          added to other groups as members.)
ECHO.
Echo.
ECHO 	3. Exit
ECHO.
ECHO    Enter (1, 2 or 3)
ECHO.

SET /P SELECTION=

IF %SELECTION% EQU 1 (
GOTO BackupTasks
)
IF %SELECTION% EQU 2 (
GOTO RestoreTasks
)
IF %SELECTION% EQU 3 (
GOTO Finish
)
GOTO Selection

:BackupTasks
::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 1) Export OU Structure

Set ReferenceFile=OUStructureExport.csv
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageOUStructure.ps1"' -Action Export -ReferenceFile %ReferenceFile%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 2) Export Groups

Set ReferenceFile=GroupExport.csv
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageGroups.ps1"' -Action Export -ReferenceFile %ReferenceFile%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 3) Export Users

Set ReferenceFile=UserExport1.csv
::Set SearchBase=OU=Sites,OU=FMG,DC=FMG,DC=local
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageUsers.ps1"' -Action Export -ReferenceFile %ReferenceFile%"
::powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageUsers.ps1"' -Action Export -SearchBase '%SearchBase%' -ReferenceFile %ReferenceFile%"

::Set ReferenceFile=UserExport2.csv
::Set SearchBase=OU=Corporate,DC=FMG,DC=local
::powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageUsers.ps1"' -Action Export -ReferenceFile %ReferenceFile%"
::powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageUsers.ps1"' -Action Export -SearchBase '%SearchBase%' -ReferenceFile %ReferenceFile%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo The backup process is complete.

GOTO Finish

:RestoreTasks
::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 1) Import OU Structure

Set ReferenceFile=OUStructureNew.csv
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageOUStructure.ps1"' -Action Import -ReferenceFile %ReferenceFile%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 2) Import Groups
::Set ReferenceFile=GroupExport.csv
::powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageGroups.ps1"' -Action Import -ReferenceFile %ReferenceFile%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 3) Import Users

::Set ReferenceFile=UserExport1.csv
::powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageUsers.ps1"' -Action Import -ReferenceFile %ReferenceFile%"

::Set ReferenceFile=UserExport2.csv
::powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageUsers.ps1"' -Action Import -ReferenceFile %ReferenceFile%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 4) Import Groups

:: This script is run twice to ensure that all groups are created and added to other groups as members.
:: It also ensures that any groups used for "Primary Groups" and "Managed By" groups are available and correctly applied.
Set ReferenceFile=GroupExport.csv
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageGroups.ps1"' -Action Import -ReferenceFile %ReferenceFile%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo The restore process is complete.

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

:Finish
EndLocal
Exit /b 0
