Get-ADDomainController -Filter * | select name, operatingsystem,HostName,site,IsGlobalCatalog,IsReadOnly,IPv4Address |
Export-Csv DomainControllers_Report.csv
