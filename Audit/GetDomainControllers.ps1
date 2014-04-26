Import-Module ActiveDirectory 

# Get Domain Controllers
$Servers = Get-ADDomainController -Filter *

# Get All Servers
#$Servers = Get-ADComputer -Filter * -Properties Name,Operatingsystem | Where-Object {$_.Operatingsystem -like "*server*"}

$Servers.count

ForEach ($Server in $Servers){

if (Test-Connection -Cn $Server.Name -BufferSize 16 -Count 1 -ea 0 -quiet) {
write-host -ForegroundColor Green $Server.Name "is online"
} Else {
write-host -ForegroundColor Red $Server.Name "is offline"
}


}

