# Author: Michael Schönburg
# Stand: 04.12.2020

ForEach($u in $usersUnsuccessfull)
{
    $pathLog = "C:\Users\$($u)\AppData\Roaming\ITCE-Logfiles\Script_OL-MRU-Cleanup.2.log"
    Remove-Item $pathLog -Confirm:$false
}
