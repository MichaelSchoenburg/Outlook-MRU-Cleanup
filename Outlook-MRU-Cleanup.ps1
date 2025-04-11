# Author: Michael Schönburg
# Stand: 04.12.2020
# Zweck: Most Recently Used Lists in Outlook löschen (notwendig nach Migration)

$logFile = "$($env:APPDATA)\ITCE-Logfiles\Script_OL-MRU-Cleanup.2.log"
$tarPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\"
$tarFolder = "0a0d020000000000c000000000000046"
$regFile = "\\DomainController.Domain.TLD\netlogon\MRU.reg"

if (Test-Path $logFile)
{
    Write-Host "Logfile existiert bereits. Script wird nicht ausgeführt."
}
else
{    
    Start-Transcript -Path $logFile # überschreiben
    $Profile = Get-ChildItem $tarPath

    if ($Profile.Length -gt 1)
    {
        Write-Host "Der Benutzer hat mehr, als ein Outlook-Profil."

        # Finde Namen des genutzten Outlook-Profils
        $outlookApplication = New-Object -ComObject 'Outlook.Application'
        if ($outlookApplication.Application.DefaultProfileName)
        {
            $tarPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$($outlookApplication.Application.DefaultProfileName)"
        }
        else
        {
            Write-Host -ForegroundColor Red "Konnte outlookApplication.Application.DefaultProfileName nicht abrufen. Script wird beendet."
            Exit
        }
    }
    else
    {
        $tarPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$((Get-ChildItem $tarPath)[0])"
    }
    
    Write-Host "tarPath = $($tarPath)"

    # Registry bearbeiten
    if (Test-Path "$($tarPath)\$($tarFolder)")
    {
        Remove-Item -Path "$($tarPath)\$($tarFolder)"
    
        $ci = Get-ChildItem $tarPath
    
        if ($ci -notcontains $tarFolder)
        {
            
            Write-Host "Starte Registryimport..."

            $pinfo = New-Object System.Diagnostics.ProcessStartInfo
            $pinfo.FileName = "reg.exe"
            $pinfo.RedirectStandardError = $true
            $pinfo.RedirectStandardOutput = $true
            $pinfo.UseShellExecute = $false
            $pinfo.Arguments = "IMPORT $($regFile)"
            $p = New-Object System.Diagnostics.Process
            $p.StartInfo = $pinfo
            $p.Start() | Out-Null
            $stdout = $p.StandardOutput.ReadToEnd()
            $stderr = $p.StandardError.ReadToEnd()
            $p.WaitForExit()
            Write-Host "stdout: $stdout"
            Write-Host "stderr: $stderr"

            if ($p.ExitCode -eq 0)
            {
                Write-Host "Registryimport erfolgreich abgeschlossen."
            }
            else
            {
                Write-Host "Registryimport nicht erfolgreich abgeschlossen. Exitcode = $($p.ExitCode)"
            }
        }
        else
        {
            Write-Host "Löschen des Ordners nicht erfolgreich."
        }
    }
    else
    {
        Write-Host "Ordner $($tarPath)\$($tarFolder) existiert gar nicht."
    }

    Write-Host "Script wird beendet."

    Stop-Transcript
}
