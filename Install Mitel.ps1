$wshShell = New-Object -ComObject "WScript.Shell"
$urlShortcut = $wshShell.CreateShortcut(
  (Join-Path "$env:USERPROFILE\Desktop" "Test.url")
)
$urlShortcut.TargetPath = "https://upgrade01.sky.shoretel.com/ClientInstall/NonAdmin"
$urlShortcut.Save()

$ProgressPreference = 'SilentlyContinue'
$tempLoc = "$env:LOCALAPPDATA\temp-mitelconnect"
function Get-TimeStamp {
    return "[{0:yyyy/MM/dd} {0:HH:mm:ss}]" -f (Get-Date)
}

Write-Output "$(Get-TimeStamp) - Init"

mkdir $tempLoc -ErrorAction SilentlyContinue
Set-Location $tempLoc
Invoke-WebRequest "https://upgrade.connect.mitelcloud.com/ClientInstall/213.100.5664.0/non-admin/MitelConnectGPO.zip" -OutFile MitelConnectGPO.zip -Verbose

Write-Output "$(Get-TimeStamp) - Download complete!"

$hash = (Get-FileHash MitelConnectGPO.zip).Hash
Write-Output "Hash: $hash"

Expand-Archive MitelConnectGPO.zip -Force -Verbose
Set-Location $tempLoc\MitelConnectGPO

Get-ChildItem

# Run installer syncrhonously to know when it finishes
& msiexec /i "Mitel Connect.msi" /quiet /qn /norestart /L*v "$env:LOCALAPPDATA\Programs\Mitel\Connect\Install.log" | Out-Null
Write-Output "$(Get-TimeStamp) - Installation complete!"

$isInstalledAfter = Test-Path "$env:LOCALAPPDATA\Programs\Mitel\Connect\Mitel.exe"

Write-Output "$(Get-TimeStamp) - Found $isInstalledAfter! Cleaining up..."

Set-Location $env:LOCALAPPDATA
Remove-Item -Path "$env:LOCALAPPDATA\temp-mitelconnect" -Recurse -Force -Verbose

if ($isInstalledAfter) {
    Exit 0
}

Exit 1
