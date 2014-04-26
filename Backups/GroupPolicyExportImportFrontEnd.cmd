@echo off

:: Using PowerShell to Back Up Group Policy Objects:
:: - http://blogs.technet.com/b/heyscriptingguy/archive/2014/01/04/using-powershell-to-back-up-group-policy-objects.aspx

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
Echo ** Group Policy Export/Import Front-End **
Echo ------------------------------------------
Echo.
ECHO You are currently logged into the %userdnsdomain% Domain.
Echo.
ECHO Which set of tasks do you want to run?
ECHO.
ECHO 	1. Backup/Export Tasks
Echo         - Backup ADMX Central Store
Echo         - Export WMI Filters
Echo         - Backup GPOs
Echo         - Export GPO Links
Echo         - Create and Modify Migration Table
Echo.
ECHO 	2. Restore/Import Tasks
Echo         - Restore ADMX Central Store
Echo         - Import WMI Filters
Echo         - Create objects from Migration table
Echo         - Import GPOs
Echo         - Link WMI Filters
Echo         - Import GPO Links
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
Echo 1) Backup ADMX Central Store
Set BackupLocation=%~dp0ADMXCentralStore
Set PoliciesLocation=\\%userdnsdomain%\SYSVOL\%userdnsdomain%\Policies

IF NOT EXIST "%BackupLocation%" MD "%BackupLocation%"
xcopy /e/s/v/y "%PoliciesLocation%\PolicyDefinitions\*.*" "%BackupLocation%\"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 2) Export WMI Filters
Set ReferenceFile=WMIFiltersExport.csv
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageWMIFilters.ps1"' -Action Export -ReferenceFile %ReferenceFile%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 3) Backup GPOs
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0GPOBackup.ps1"'"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 4) Export GPO Links
Set ReferenceFile=GPOLinksExport.csv
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageGPOLinks.ps1"' -Action Export -ReferenceFile %ReferenceFile%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Echo.
Echo 5) Create and Modify Migration Table
Set MigrationTable=AllGPO.migtable
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0CreateMigrationTable.ps1"' -Action Create -MigrationTable %MigrationTable%"
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0CreateMigrationTable.ps1"' -Action Modify -MigrationTable %MigrationTable%"

:: Note: You can also use the CreateMigrationTable.wsf script that is part of the Group Policy
:: Sample Scripts downloadable from Microsoft. This also requires the Lib_CommonGPMCFunctions.js
:: script, which contains the common functions. You must enter a folder location for the migtable
:: file, otherwise it defaults to %SystemRoot%\System32.
::cscript CreateMigrationTable.wsf %~dp0source.migtable /OverWrite /MapByName /AllGPOs

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo The backup process is complete.

GOTO Finish

:RestoreTasks
::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 1) Restore ADMX Central Store
Set BackupLocation=%~dp0ADMXCentralStore
Set PoliciesLocation=\\%userdnsdomain%\sysvol\%userdnsdomain%\policies

IF NOT EXIST "%PoliciesLocation%\PolicyDefinitions" MD "%PoliciesLocation%\PolicyDefinitions"
xcopy /e/s/v/y "%BackupLocation%\*.*" "%PoliciesLocation%\PolicyDefinitions\"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 2) Import WMI Filters
Set ReferenceFile=WMIFiltersExport.csv
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageWMIFilters.ps1"' -Action Import -ReferenceFile %ReferenceFile%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 3) Create objects from Migration table
Set MigrationTable=AllGPO.migtable
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0CreateObjectsFromMigrationTable.ps1"' -MigrationTable %MigrationTable%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 4) Import GPOs
Set MigrationTable=AllGPO.migtable
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ImportGPOsFromBackup.ps1"' -MigrationTable %MigrationTable%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 5) Link WMI Filters
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0LinkWMIFilters.ps1"'"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 6) Import GPO Links
Set ReferenceFile=GPOLinksExport.csv
powershell.exe -ExecutionPolicy Bypass -Command "& '"%~dp0ManageGPOLinks.ps1"' -Action Import -ReferenceFile %ReferenceFile%"

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo 7) Go through GPOs and change references, such as...
Echo    - DNS Client Search Order List – MUST BE DONE
Echo    - WSUS Patch Management
Echo    - Firewall Rules
Echo    - There are also a couple of Computer Startup scripts that run from the group policies that help manage the SNMP, KMS Licensing and Antivirus Deployment Settings. You'll need to modify these.

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Echo.
Echo The restore process is complete.

::xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

:Finish
EndLocal
Exit /b 0
