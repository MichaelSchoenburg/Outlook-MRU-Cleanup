# Author: Michael Schönburg
# Stand: 04.12.2020

cd "C:\users"
$listUsers = Get-ChildItem
$usersUnsuccessfull = @()

ForEach($user in $listUsers.Name)
{
    $pathLog = "C:\Users\$($user)\AppData\Roaming\ITCE-Logfiles\Script_OL-MRU-Cleanup.2.log"
    if (Test-Path $pathLog)
    {
        $logOutput = Get-Content $pathLog
        
        # Erfolgreich
        $Status = $logOutput.where{$_.StartsWith("Registryimport")}
        if ($Status)
        {
            Write-Host "$($user): $($Status)" -ForegroundColor Green
        }
        
        # Ordner existiert gar nicht
        $Status = $logOutput.where{$_.StartsWith("Ordner")}
        if ($Status)
        {
            Write-Host "$($user): $($Status)" -ForegroundColor Yellow
            $usersUnsuccessfull += $user
        }
    }
    else
    {
        Write-Host "$($user): Script wurde noch nicht ausgeführt" -ForegroundColor Gray
    }
}
