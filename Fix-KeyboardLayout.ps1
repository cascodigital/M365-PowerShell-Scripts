# Fix-KeyboardLayout.ps1
# Fixes Windows automatic keyboard layout switching from PT-BR (ABNT2) to EN-US
# Author: Hall
# Description: Forces ABNT2 keyboard layout on English Windows installation

# 1. Configure Language List (Display English + Input ABNT2)
# Creates the language object for English (US)
$LangList = New-WinUserLanguageList en-US

# Clears default input methods (which would include US keyboard)
$LangList[0].InputMethodTips.Clear()

# Adds specifically ABNT2 keyboard (Code: 0416:00000416) to English language
$LangList[0].InputMethodTips.Add('0416:00000416')

# Applies the new list (This removes phantom US keyboards)
Set-WinUserLanguageList $LangList -Force

# 2. Disable Keyboard Switching Shortcuts (Ctrl+Shift / Alt+Shift)
# Prevents accidental switching during use
$RegPath = "HKCU:\Keyboard Layout\Toggle"
if (!(Test-Path $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
Set-ItemProperty -Path $RegPath -Name "Hotkey" -Value "3"
Set-ItemProperty -Path $RegPath -Name "Language Hotkey" -Value "3"
Set-ItemProperty -Path $RegPath -Name "Layout Hotkey" -Value "3"

# 3. Force Override Input Method in User Profile
# Ensures Windows uses ABNT2 regardless of window language
$ProfilePath = "HKCU:\Control Panel\International\User Profile"
Set-ItemProperty -Path $ProfilePath -Name "InputMethodOverride" -Value "0416:00000416"

Write-Host "Configuration applied. Logoff/Login recommended to activate registry changes." -ForegroundColor Green
