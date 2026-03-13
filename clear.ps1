# Clear lock screen cached users
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" /v LastLoggedOnUser /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" /v LastLoggedOnSAMUser /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI" /v LastLoggedOnDisplayName /f

Write-Host "Stopping Office / OneDrive processes..."

$apps = "OneDrive","Outlook","WinWord","Excel","PowerPoint","OneNote","Teams","olk","msedgewebview2"
foreach ($app in $apps) {
    Get-Process $app -ErrorAction SilentlyContinue | Stop-Process -Force
}

Start-Sleep -Seconds 5

Write-Host "Clearing Office authentication and licensing caches..."

$paths = @(
"$env:LOCALAPPDATA\Microsoft\Office\16.0\Licensing",
"$env:LOCALAPPDATA\Microsoft\Office\Licensing",
"$env:LOCALAPPDATA\Microsoft\Office\LicensingNext",
"$env:PROGRAMDATA\Microsoft\Office\Licensing",
"$env:LOCALAPPDATA\Microsoft\IdentityCache",
"$env:LOCALAPPDATA\Microsoft\OneAuth",
"$env:LOCALAPPDATA\Microsoft\TokenBroker",
"$env:LOCALAPPDATA\Microsoft\Office\16.0\OfficeFileCache",
"$env:LOCALAPPDATA\Microsoft\Credentials",
"$env:LOCALAPPDATA\Microsoft\Vault",
"$env:LOCALAPPDATA\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy"
)

foreach ($path in $paths) {
    if (Test-Path $path) {
        Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "Removing Outlook profiles..."

Remove-Item "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles" -Recurse -Force -ErrorAction SilentlyContinue

Remove-Item "$env:LOCALAPPDATA\Packages\Microsoft.OutlookForWindows_8wekyb3d8bbwe" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Clearing Office identity registry keys..."

Remove-Item "HKCU:\Software\Microsoft\Office\16.0\Common\Identity" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Microsoft\IdentityCRL" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Setting default Outlook profile..."

New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Force | Out-Null

New-ItemProperty `
-Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" `
-Name "DefaultProfile" `
-Value "Outlook" `
-PropertyType String `
-Force | Out-Null

# -------- Office Licensing Reset --------

Write-Host "Clearing Office subscription licensing..."

Remove-Item "HKCU:\Software\Microsoft\Office\16.0\Common\Licensing" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "HKCU:\Software\Microsoft\Office\16.0\Common\SignIn" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Restarting Office licensing service..."

Stop-Service ClickToRunSvc -Force -ErrorAction SilentlyContinue
Start-Service ClickToRunSvc

# ----------------------------------------

Write-Host "Clearing OneDrive configuration..."

Remove-Item "$env:LOCALAPPDATA\Microsoft\OneDrive\settings" -Recurse -Force -ErrorAction SilentlyContinue

# Additional OneDrive identity caches
Remove-Item "$env:LOCALAPPDATA\Microsoft\OneDrive\logs" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$env:LOCALAPPDATA\Microsoft\OneDrive\setup" -Recurse -Force -ErrorAction SilentlyContinue

# OneDrive account registry cache
Remove-Item "HKCU:\Software\Microsoft\OneDrive\Accounts" -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "Clearing Teams configuration..."

Write-Host "Resetting Microsoft Teams..."

# Classic Teams
$classicTeams = "$env:APPDATA\Microsoft\Teams"

if (Test-Path $classicTeams) {
    Remove-Item $classicTeams -Recurse -Force -ErrorAction SilentlyContinue
}

# New Teams (Teams 2.x)
$newTeams = "$env:LOCALAPPDATA\Packages\MSTeams_8wekyb3d8bbwe"

if (Test-Path $newTeams) {
    Remove-Item $newTeams -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Clearing Windows Credential Manager..."

$creds = cmdkey /list | Select-String "Target:" | ForEach-Object { ($_ -split ":")[1].Trim() }

foreach ($cred in $creds) {
    cmdkey /delete:$cred
}

Write-Host "Refreshing Azure AD token..."

dsregcmd /refreshprt

Write-Host "Configuring OneDrive first-run launch..."

New-ItemProperty `
-Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce" `
-Name "OneDriveSetup" `
-Value 'powershell.exe -WindowStyle Hidden -Command "Stop-Process -Name OneDrive -Force -ErrorAction SilentlyContinue; Start-Process ''C:\Program Files\Microsoft OneDrive\OneDrive.exe''"' `
-PropertyType String `
-Force

Write-Host ""
Write-Host "Reset complete. Reboot recommended."

Restart-Computer -Force