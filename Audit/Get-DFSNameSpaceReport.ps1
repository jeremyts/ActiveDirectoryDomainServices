<#
  This script will find all Domain DFS Namespaces (Roots) and export them
  to two separate CSV files. One is a high level report; the other it a full
  report. Along the way it also formally exports each Namespace using the
  DFSUtil tool and addresses a known bug with that export process.

  The script creates 4 hash tables that can be linked to provide full
  reporting in whichever way is needed:
  - $namespacehashtable
  - $namespacetargethashtable
  - $linkhashtable
  - $linktargethashtable

  The hierarchy has been tied together with parent objects to allow you to
  create some extensive reports:
    -- $namespacehashtable
      |-- $namespacetargethashtable
      |-- $linkhashtable
         |-- $linktargethashtable

  Script Name: Get-DFSNamespaceReport.ps1
  Release 1.1
  Modified by Jeremy@jhouseconsulting.com 23/1/2014
  Written by Jeremy@jhouseconsulting.com 22/1/2014

  The DFS configuration Data queried by DFSUtil is stored within the
  following location within Active Directory: 
    CN=Dfs-Configuration,CN=System,DC=<domain DN> 

  In Windows Server 2003, each Domain DFS Root/Namespace is stored
  within an "fTDfs" object which contains an attribute "pKT" containing
  the configuration data (namespace settings, namespace servers, folder
  targets, etc).  For instance, the "DATA" namespace listed in the
  dfsutil.exe output above is located with an fTDfs object at this
  location:  CN=DATA,CN=Dfs-Configuration,CN=System,DC=<domain DN>.
  No parts of this object should ever be modified directly.  
    CN=Dfs-Configuration,CN=System,DC=<domain DN>
           |_fTDfs 

  In Windows Server 2008, Domain DFS Roots/Namespaces may be configured
  in "Windows Server 2008 mode".  In this mode, configuration data is
  stored under an "msDFS-NamespaceAnchor" class object. An object of
  class "msDFS-Namespacev2" represents each root, and each root contains
  an "msDFS-Linkv2" object representing each hosted link. 
    CN=Dfs-Configuration,CN=System,DC=<domain DN>
           |_msDFS-NamespaceAnchor
                   |_msDFS-Namespacev2
                           |_msDFS-Linkv2

  As some of the DFS information is stored in "blobs" and not consistent
  between versions, DFSUtil is used to export each namespace to an XML
  file providing consistency for reading and further reporting.

  References:
  - http://support.microsoft.com/kb/969382
  - http://pinchii.com/home/2013/07/combining-multiple-shares-into-one-dfs-folder-with-powershell/
  - http://tompaps.blogspot.com.au/2013/07/compare-two-roots-dfs.html
  - http://www.coldtail.com/wiki/index.php?title=PowerShell:_DFS_Path_to_Remote_Share_Translation_Table
  - http://technet.microsoft.com/en-us/library/cc753875.aspx

  If you change this script so that you're adding multiple values per key,
  you'll need to use arrays instead to avoid the issues documented here:
  - http://gallery.technet.microsoft.com/scriptcenter/An-Array-of-PowerShell-069c30aa

  You'll note throughout the script we strongly define all collectiona as
  a collection so that the += operator will apply to the collection rather
  than the object, otherwise we will end up receiving an error stating:
  "Method invocation failed because [System.Management.Automation.PSObject]
  doesn't contain a method named 'op_Addition'. 

  Known Dfsutil bug:
    The "&" (ampersand) character is strictly illegal in XML. If any DFS
    targets (share names) within the DFS Links contain an ampersand
    character, whilst the Dfsutil root export command works correctly,
    the import will fail. Furthermore, when reading the XML file as a true
    XmlDocument type from a scripting point of view, it will also fail.

    The PowerShell Get-Content cmdlet will give you errors such as...
    Cannot convert value "System.Object[]" to type "System.Xml.XmlDocument"

    Using the ampersand character in share names is quite acceptable, as
    long as you understand the ramifications when using these shares as link
    targets in a DFS structure. For this reason Microsoft has always
    suggested that it's not a good practice.

    The solution is to open the XML file in a text editor and replacing the
    "&" (ampersand) symbols with the Entity Encoding &amp;

    The irony is that it's a bug with the tool and can be easily addressed
    by Microsoft.

#>

#-------------------------------------------------------------

# Import the Active Directory Module
Import-Module ActiveDirectory -WarningAction SilentlyContinue
if ($Error.Count -eq 0) {
  #Write-Host "Successfully loaded Active Directory Powershell's module`n" -ForeGroundColor Green
} else {
  Write-Host "Error while loading Active Directory Powershell's module : $Error`n" -ForeGroundColor Red
  exit
}

#-------------------------------------------------------------

$ReExport = $True

$DNSRoot = (get-addomain).DNSRoot

# Get the script path
$ScriptPath = {Split-Path $MyInvocation.ScriptName}

$HighLevelReportName = $(&$ScriptPath) + "\DFSNamespaceHighLevelReport.csv"
$FullReportName = $(&$ScriptPath) + "\DFSNamespaceFullReport.csv"
$OutputPath = $(&$ScriptPath) + "\DFSNExport"

#-------------------------------------------------------------

Function Get-DFSNamespaces($FQDN,$Namespaces,$Export,$ExportPath)
{
  $resultarray = @()
  ForEach($Namespace in $Namespaces) {
    $Name = $Namespace.Name
    $objectClass = $Namespace.objectClass
    switch ($objectClass)
      {
        "msDFS-Namespacev2" {$Version = "v2";break}
        "fTDfs" {$Version = "v1";break}
        default {$Version = "unknown"}
      }
    If (-not(Test-Path -Path $ExportPath)) {
      New-Item -Path $ExportPath -ItemType Directory | out-null
    }
    write-host -ForegroundColor green "- \\$FQDN\$Name"
    # The command variable needs to be correctly constructed to ensure you are passing
    # type a 'System.String' object to the invoke-expression cmdlet.
    $command = "cmd /c dfsutil.exe root export ""\\$FQDN\$Name"" ""$ExportPath\$Name.XML"""
    $err = invoke-expression $command

    $obj = New-Object -TypeName PSObject
    $obj | Add-Member -MemberType NoteProperty -Name "Name" -value $Name
    $obj | Add-Member -MemberType NoteProperty -Name "objectClass" -value $objectClass
    $obj | Add-Member -MemberType NoteProperty -Name "Version" -value $Version
    $resultarray += $obj
    $obj = $NULL
  }
  $resultarray | write-output
}

#-------------------------------------------------------------

If ($ReExport) {

  $namespaces = @()

  If ((Get-ADObject -Filter 'objectClass -eq "fTDfs"' -Properties "name" | Measure-Object).count -gt 0 ) {
    write-host -ForegroundColor green "`nExporting the v1 Namespaces..."
    $DfsNamespaces = Get-ADObject -Filter 'objectClass -eq "fTDfs"' -Properties "name"  | Sort-object Name
    $namespaces += Get-DFSNamespaces $DNSRoot $DfsNamespaces $ExportDFSNamespace $OutputPath
  }


  If ((Get-ADObject -Filter 'objectClass -eq "msDFS-Namespacev2"' -Properties "name" | Measure-Object).count -gt 0 ) {
    write-host -ForegroundColor green "`nExporting the v2 Namespaces..."
    $DfsNamespaces = Get-ADObject -Filter 'objectClass -eq "msDFS-Namespacev2"' -Properties "name"  | Sort-object Name
    $namespaces += Get-DFSNamespaces $DNSRoot $DfsNamespaces $ExportDFSNamespace $OutputPath
  }

  $NamespaceCount = ($namespaces | Measure-Object).count

  If ($NamespaceCount -eq 0) {
    Write-Host -ForegroundColor red "No Active Directory DFS Namespaces were found."
    Exit
  }

}

IF (Test-Path -Path $OutputPath) {

  [PSObject[]] $HighLevelReport = @()
  $namespacehashtable = $null
  $namespacehashtable = @{}
  $namespacetargethashtable = $null
  $namespacetargethashtable = @{}
  $linkhashtable = $null
  $linkhashtable = @{}
  $linktargethashtable = $null
  $linktargethashtable = @{}

  IF ((Get-ChildItem $OutputPath\*.xml | Measure-Object).Count -gt 0) {
    $XMLFiles = Get-ChildItem $OutputPath\*.xml | Sort-object Name
    $XMLFileCount = ($XMLFiles | Measure-Object).Count
    If ($ReExport -eq $False) {
      $NamespaceCount = $XMLFileCount
    }
    If ($NamespaceCount -eq $XMLFileCount) {
      write-host -ForegroundColor green "`nProcessing the XML files..."
      [PSObject[]] $rootobj = @()
      ForEach ($XMLFile in $XMLFiles) {
        Write-Host -ForegroundColor green "- Reading the $($XMLFile.Name) file"

        # Replace any "&" (ampersand) characters with the &amp; Entity Encoding
        (Get-Content "$OutputPath\$($XMLFile.Name)") | ForEach-Object {
          $_ -replace '&amp;','&' -replace '&','&amp;'
          } | out-file "$OutputPath\$($XMLFile.Name)" -Fo -En ascii

        # Specify XML type which will cast the text returned from get-content
        # to an XmlDocument object.
        [xml]$xml = Get-Content "$OutputPath\$($XMLFile.Name)"

        $NamespaceName = ($xml.Root.Name).Split("\\")[3]
        $rootobj = New-Object -TypeName PSObject
        If ($ReExport) {
          ForEach ($namespace in $namespaces) {
            If ($namespace.Name -eq $NamespaceName) {
              $RootName = $namespace.Name
              $RootObjectClass = $namespace.objectClass
              $RootVersion = $namespace.Version
            }
          }
          $rootobj | Add-Member -MemberType NoteProperty -Name "Name" -value $RootName
          $rootobj | Add-Member -MemberType NoteProperty -Name "ObjectClass" -value $RootObjectClass
          $rootobj | Add-Member -MemberType NoteProperty -Name "Version" -value $RootVersion
        } Else {
          $RootName = $NamespaceName
          $rootobj | Add-Member -MemberType NoteProperty -Name "Name" -value $RootName
        }
        $rootobj | Add-Member -MemberType NoteProperty -Name "Path" -value $xml.Root.Name
        $rootobj | Add-Member -MemberType NoteProperty -Name "State" -value $xml.Root.State
        $rootobj | Add-Member -MemberType NoteProperty -Name "Timeout" -value $xml.Root.Timeout
        $rootobj | Add-Member -MemberType NoteProperty -Name "SiteCosting" -value $xml.Root.SITECOSTING
        $rootobj | Add-Member -MemberType NoteProperty -Name "RootScalability" -value $xml.Root.ROOTSCALABILITY
        $rootobj | Add-Member -MemberType NoteProperty -Name "ABDE" -value $xml.Root.ABDE
        #$rootobj | Add-Member -MemberType NoteProperty -Name "Insite" -value $xml.Root.Insite
        #$rootobj | Add-Member -MemberType NoteProperty -Name "TargetfailBack" -value $xml.Root.TargetfailBack

        # Namespace Servers
        [PSObject[]] $targetobj = @()
        [PSObject[]] $Targets = @()
        foreach ($target in $xml.Root.Target) {
          $targetobj = New-Object -TypeName PSObject
          $targetobj | Add-Member -MemberType NoteProperty -Name "TargetPath" -value $Target.'#Text'
          $targetobj | Add-Member -MemberType NoteProperty -Name "TargetParent" -value $RootName
          $targetobj | Add-Member -MemberType NoteProperty -Name "TargetState" -value $Target.State
          $Targets += $targetobj
        }
        $rootobj | Add-Member -MemberType NoteProperty -Name "Targets" -value $targets
        $HighLevelReport += $rootobj

        $namespacetargethashtable = $namespacetargethashtable + @{($RootName) = $Targets}

        # Folders and Targets
        [PSObject[]] $linkobj = @()
        foreach ($link in $xml.Root.Link) {
          $linkobj = New-Object -TypeName PSObject
          $linkobj | Add-Member -MemberType NoteProperty -Name "LinkName" -value $Link.Name
          $linkobj | Add-Member -MemberType NoteProperty -Name "LinkParent" -value $xml.Root.Name
          $linkobj | Add-Member -MemberType NoteProperty -Name "LinkState" -value $Link.State
          $linkobj | Add-Member -MemberType NoteProperty -Name "LinkTimeout" -value $Link.Timeout

          [PSObject[]] $linktargetobj = @()
          [PSObject[]] $LinkTargets = @()
          ForEach ($LinkTarget in $Link.Target) {
            $linktargetobj = New-Object -TypeName PSObject
            $linktargetobj | Add-Member -MemberType NoteProperty -Name "LinkTargetName" -value $LinkTarget.'#Text'
            $linktargetobj | Add-Member -MemberType NoteProperty -Name "LinkTargetParent" -value $Link.Name
            $linktargetobj | Add-Member -MemberType NoteProperty -Name "LinkTargetState" -value $LinkTarget.State
            $LinkTargets += $linktargetobj
            # Build the $linktargethashtable hash table
            # The shares (AKA link targets) may be used across different DFS links. Therefore
            # it may not be unique. If it isn't, we add "append" the value. It therefore becomes
            # a multidimentional collection due to the one-to-many relationships.
            If (!($linktargethashtable.ContainsKey($LinkTarget.'#Text'))) {
              $linktargethashtable = $linktargethashtable + @{($LinkTarget.'#Text') = $linktargetobj}
              #$linktargethashtable.add(($LinkTarget.'#Text'),$linktargetobj)
            } else {
              $existingvalue = $linktargethashtable.Get_Item($LinkTarget.'#Text')
              $newvalue = @($existingvalue,$linktargetobj)
              $linktargethashtable.Set_Item(($LinkTarget.'#Text'),$newvalue)
            }

          }
          $linkobj | Add-Member -MemberType NoteProperty -Name "LinkTargets" -value  $LinkTargets

          # Build the $linkhashtable hash table
          # The Link Names may not be unique across all DFS namespaces, so we
          # must prepend the root path to give the key uniquness.
          $linkhashtable = $linkhashtable + @{($xml.Root.Name+"\"+$Link.Name) = $linkobj}

        }
        # Build the $namespacehashtable hash table
        $namespacehashtable = $namespacehashtable + @{$RootName = $rootobj}

      }#ForEach
      write-host -ForegroundColor green "`nCreating the high level report..."
      $HighLevelReport | export-csv -notype -path "$HighLevelReportName"
      # Remove the quotes
      (get-content "$HighLevelReportName") |% {$_ -replace '"',""} | out-file "$HighLevelReportName" -Fo -En ascii

    } ELSE {
      Write-Host -ForegroundColor red "There is a count mismatch`n - XML file count: $XMLFileCount`n - Namespace count in Active Directory: $NamespaceCount"
      Exit
    }
  } ELSE {
    Write-Host -ForegroundColor red "There are no XML files to process."
    Exit
  }
} ELSE {
  Write-Host -ForegroundColor red "The $OutputPath folder is missing."
  Exit
}

write-host -ForegroundColor green "`nCreating the full report..."

######################### Roots/Namespaces ############################
# These commands are samples of how you can extract data for reporting purposes

# Output the hashtable, sorting by the key name
#$namespacehashtable.GetEnumerator() | Sort-Object Name | Format-Table -AutoSize

# Output the keys, sorted
#$namespacehashtable.keys | sort

# Output the values, sorted by the name value
$namespacehashtable.values | ForEach {$_ } | ForEach {$_ } | Sort-Object Name | Format-Table -AutoSize

# Output a Namespace
#$namespacehashtable."F_Drive"

# Output the Targets of the Namespace
#$namespacehashtable."F_Drive".Targets | Format-Table -AutoSize

###################### Root/Namespace Targets #########################
# These commands are samples of how you can extract data for reporting purposes

# Output the hashtable, sorting by the key name
#$namespacetargethashtable.GetEnumerator() | Sort-Object Name | Format-Table -AutoSize

# Output the keys, sorted
#$namespacetargethashtable.keys | sort

# Output the values, sorted by the TargetParent value
#$namespacetargethashtable.values | ForEach {$_ } | ForEach {$_ } | Sort-Object TargetParent | Format-Table -AutoSize

# Output the value of a key
#$namespacetargethashtable."F_Drive" | Format-Table -AutoSize

############################### Links #################################
# These commands are samples of how you can extract data for reporting purposes

# Output the hashtable, sorting by the key name
#$linkhashtable.GetEnumerator() | Sort-Object Name | Format-Table -AutoSize

# Output the keys, sorted
#$linkhashtable.keys | sort

# Output the values, sorted by the LinkName value
#$linkhashtable.values | ForEach {$_ } | ForEach {$_ } | Sort-Object LinkName | Format-Table -AutoSize

# Output the value of a key
#$linkhashtable."\\FMG\F_Drive\temp doc ctrl"

# Output the Targets of the Links
#$linkhashtable."\\FMG\F_Drive\temp doc ctrl".LinkTargets | Format-Table -AutoSize
#$linkhashtable."\\FMG\PERTH\T.GIS\GIS_Data_Warehouse".LinkTargets | Format-Table -AutoSize

########################### Link Targets ##############################
# These commands are samples of how you can extract data for reporting purposes

# Output the hashtable, sorting by the key name
#$linktargethashtable.GetEnumerator() | Sort-Object Name | Format-Table -AutoSize

# Output the keys, sorted
#$linktargethashtable.keys | sort

# Outputs a sorted list of all Targets (shares) used in DFS links, along with the state and parent (link)
#$linktargethashtable.values | ForEach {$_ } | ForEach {$_ } | Sort-Object LinkTargetName | Format-Table -AutoSize

# Output the value of a key
#$linktargethashtable."\\archive\Perth.DocControlDump$" | Format-Table -AutoSize
