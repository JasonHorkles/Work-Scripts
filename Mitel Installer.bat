@echo off
echo:
echo ================================================================
echo        ___  ___                 __  ___                 ___  __
echo ^|\/^| ^|  ^|  ^|__  ^|       ^| ^|\ ^| /__`  ^|   /\  ^|    ^|    ^|__  ^|__)
echo ^|  ^| ^|  ^|  ^|___ ^|___    ^| ^| \^| .__/  ^|  /~~\ ^|___ ^|___ ^|___ ^|  \
echo:
echo ===========================[ v1.0.0 ]===========================
echo:
cd "%USERPROFILE%\Desktop"
powershell -Command "$ProgressPreference = 'SilentlyContinue'; $tempLoc = \"$env:USERPROFILE\Desktop\temp-mitelconnect\"; Write-Output \"We are now downloading Mitel Connect! Please wait...`nNote: This may take several minutes!\"; mkdir $tempLoc -ErrorAction SilentlyContinue -Force > $null; Set-Location $tempLoc; Invoke-WebRequest \"https://upgrade.connect.mitelcloud.com/ClientInstall/213.100.5664.0/non-admin/MitelConnectGPO.zip\" -OutFile MitelConnectGPO.zip; Write-Output \"Download complete! We will now unzip the folder...\"; Expand-Archive MitelConnectGPO.zip -Force; Write-Output \"Unzip finished! We will now begin the Mitel Connect installation process...\"; Start-Sleep -Seconds 3; Set-Location $tempLoc\MitelConnectGPO; & msiexec /i \"Mitel Connect.msi\" /passive /norestart | Out-Null; Write-Output \"Installation complete! Let's check to make sure it installed properly...\"; $isInstalledAfter = Test-Path \"$env:LOCALAPPDATA\Programs\Mitel\Connect\Mitel.exe\"; if ($isInstalledAfter) {Write-Output \"Good news! Mitel Connect has been properly installed!\"; Remove-Item -Path \"$env:USERPROFILE\Desktop\Install Mitel.bat\" -Force -ErrorAction SilentlyContinue} else {Write-Output \"Oh no! Mitel Connect wasn't found! Please restart your computer and try again. If it fails again, please create an i2 IT ticket for more help.\"}; Start-Sleep -Seconds 3; Write-Output \"We will now clean up our changes...\"; Set-Location $env:USERPROFILE\Desktop; Remove-Item -Path $tempLoc -Recurse -Force; Write-Output \"We have successfully cleaned up!\";"
echo Script complete! It will delete itself in a moment...
timeout /T 5 /NOBREAK

@REM Delete itself
(goto) 2>nul & del "%~f0"
