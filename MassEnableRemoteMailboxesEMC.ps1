Function Enableonprem{

$PeopleToEnable = 'C:\SuperscriptProd\Test\UsernamesNeedEnabled.csv'

Import-Csv $PeopleToEnable | ForEach {

$User = $_.Username

Enable-RemoteMailbox -Identity $User -RemoteRoutingAddress ($User + '@callruby.com') -Verbose

}
}
Enableonprem