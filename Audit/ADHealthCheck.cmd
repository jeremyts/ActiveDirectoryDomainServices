@echo off
Echo Active Directory Health Check

:: Original Script written by pd techguru found here:
:: http://pdtechguru.wordpress.com/2012/10/04/active-directory-health-check/

:: This needs to be run from a Domain Controller where DSQuery, Dcdiag and Repadmin are available.

SetLocal

Set OutputFile=%~dp0ADHealth.txt

echo Started: %time:~-11,2%:%time:~-8,2%:%time:~-5,2% %date% > %OutputFile%

echo ================================ >> %OutputFile%
echo Domain Controllers In the Domain
echo Domain Controllers In the Domain >> %OutputFile%
echo List all the Domain Controllers in Active Directory... >> %OutputFile%
echo ================================ >> %OutputFile%
%SystemRoot%\System32\DSQUERY Server -o rdn >> %OutputFile%

GOTO replsummary
echo ====================== >> %OutputFile%
echo Repadmin - Syncall - e
echo Repadmin - Syncall - e >> %OutputFile%
echo Synchronize a specified domain controller with all replication partners, and report if the sync was successful or not... >> %OutputFile%
echo ====================== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /syncall /e >> %OutputFile%

echo ====================== >> %OutputFile%
echo Repadmin - Syncall - a
echo Repadmin - Syncall - a >> %OutputFile%
echo ====================== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /syncall /A >> %OutputFile%

echo ====================== >> %OutputFile%
echo Repadmin - Syncall - d
echo Repadmin - Syncall - d >> %OutputFile%
echo ====================== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /syncall /d >> %OutputFile%

:replsummary
echo ====================== >> %OutputFile%
echo Repadmin - Replsum
echo Repadmin - Replsum >> %OutputFile%
Echo All domain controllers should show 0 in column "Fails", and "Deltas" longer (indicating the time since the last synchronization) must be less than or at most equal to the time of replication used in the Site-Link domain Controller (30 minutes)... >> %OutputFile%
echo ===================== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /replsum * /bysrc /bydest /sort:delta >> %OutputFile%

echo ====================== >> %OutputFile%
echo Repadmin - Replsummary
echo Repadmin - Replsummary >> %OutputFile%
echo Replsummary operation quickly and concisely summarizes the replication state and relative health of a forest... >> %OutputFile%
Echo It identifies domain controllers that are failing inbound replication or outbound replication, and summarizes the results in a report... >> %OutputFile%
echo ====================== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /replsummary * >> %OutputFile%

GOTO showbackup
echo ============== >> %OutputFile%
echo Repadmin - KCC
echo Repadmin - KCC >> %OutputFile%
echo Force the KCC on targeted domain controller(s) to immediately recalculate its inbound replication topology... >> %OutputFile%
echo ============== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /kcc * >> %OutputFile%

:showbackup
echo ===================== >> %OutputFile%
echo Repadmin - showbackup
echo Repadmin - showbackup >> %OutputFile%
echo Find the last time the DCs were backed up, by reading the DSASignature attribute from all servers... >> %OutputFile%
echo ===================== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /showbackup * >> %OutputFile%

echo =================== >> %OutputFile%
echo Repadmin - Showrepl >> %OutputFile%
echo Output all replication summary information from all DCs... >> %OutputFile%
echo =================== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /showrepl *  >> %OutputFile%

echo ================ >> %OutputFile%
echo Repadmin - Queue
echo Repadmin - Queue >> %OutputFile%
echo Display inbound replication requests that the domain controller has to issue to become consistent with its source replication partners... >> %OutputFile%
echo ================ >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /queue *  >> %OutputFile%

echo ====================== >> %OutputFile%
echo Repadmin - Bridgeheads
echo Repadmin - Bridgeheads >> %OutputFile%
echo Lists the Topology information of all the bridgehead servers... >> %OutputFile%
echo ====================== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /bridgeheads * /verbose >> %OutputFile%

echo =============== >> %OutputFile%
echo Repadmin - ISTG
echo Repadmin - ISTG >> %OutputFile%
echo Inter Site Topology Generator Report... >> %OutputFile%
echo =============== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /istg * /verbose >> %OutputFile%

echo ======================= >> %OutputFile%
echo Repadmin - Showoutcalls
echo Repadmin - Showoutcalls >> %OutputFile%
echo Display calls that have not yet been answered, made by the specified server to other servers... >> %OutputFile%
echo ======================= >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /showoutcalls * >> %OutputFile%

echo ==================== >> %OutputFile%
echo Repadmin - Failcache
echo Repadmin - Failcache >> %OutputFile%
echo Display a list of failed replication events detected by the Knowledge Consistency Checker (KCC)... >> %OutputFile%
echo ==================== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /failcache * >> %OutputFile%

echo ==================== >> %OutputFile%
echo Repadmin - Showtrust
echo Repadmin - Showtrust >> %OutputFile%
echo List all domains trusted by a specified domain... >> %OutputFile%
echo ==================== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /showtrust * >> %OutputFile%

echo =============== >> %OutputFile%
echo Repadmin - Bind
echo Repadmin - Bind >> %OutputFile%
echo Display the replication features for a directory partition on a domain controller... >> %OutputFile%
echo =============== >> %OutputFile%
%SystemRoot%\System32\repadmin.exe /bind * >> %OutputFile%

echo ====== >> %OutputFile%
echo Dcdiag - DNS Test
echo Dcdiag - DNS Test >> %OutputFile%
echo Determining DNS Health... >> %OutputFile%
echo ====== >> %OutputFile%
%SystemRoot%\System32\DCdiag /Test:DNS /e /v >> %OutputFile%

echo ====== >> %OutputFile%
echo Dcdiag - Comprehensive, runs all tests
echo Dcdiag - Comprehensive, runs all tests >> %OutputFile%
echo Dcdiag analyzes the state of domain controllers in a forest or enterprise and reports any problems to help in troubleshooting... >> %OutputFile%
echo ====== >> %OutputFile%
%SystemRoot%\System32\dcdiag /c /e /v >> %OutputFile%

echo ================================ >> %OutputFile%
echo Finished: %time:~-11,2%:%time:~-8,2%:%time:~-5,2% %date% >> %OutputFile%

EndLocal
Exit /b 0
