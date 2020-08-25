<#
  Set-SpotlightImageAsWallpaper.ps1
  Author: LostLogic
  Last updated: 20200715
  Source: https://github.com/LostLogic/pouncing-powershell-kittens
  
  Description:
  Grabs the latest image value from registry and sets it as the current wallpaper for the logged-in user
  It will also set JPEG Import Quality to 100 to prevent image degredation
  
  Can be run as a scheduled task with the following command:
  Program: powershell.exe
  Arguments: -WindowStyle Hidden -ExecutionPolicy Bypass -File "PATH TO FILE\Set-SpotlightImageAsWallpaper.ps1"
  
  This will cause a flicker of a console app to launch. To work around that issue, use the RunPowershellScriptHidden.vbs in the repo and configure the paths. That again can be run as a scheduled task with the following command:
  Program: wscript.exe
  Arguments: "PATH TO FILE\RunPowershellScriptHidden.vbs"
  
  Set it to run as your user account when you are logged on with the highest priveleges checked. For optimal use, set it to recur every X minutes as you see fit.
#>

# Get current user SID
$userSID = ([Security.Principal.WindowsIdentity]::GetCurrent()).User.Value

# Location of current lockscreen from Spotlight
$currentLockscreenRegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Creative\$userSID"

# Get all Spotlight images
$spotlightImages = Get-ChildItem -Path $currentLockscreenRegPath -Recurse:$false | Select-Object Name

# Get the latest image (They are date/time stamped, so the last registry Key is the latest image)
$latestImage = (Get-ItemProperty -Path $spotlightImages[$spotlightImages.Count-1].Name.Replace("HKEY_LOCAL_MACHINE","HKLM:") | Select-Object landscapeImage).landscapeImage

# Check if wallpaper quality is set to 100, if not correct it
if(-not (Get-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name JPEGImportQuality -ErrorAction SilentlyContinue))
{
  New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name JPEGImportQuality -PropertyType DWord -Value 100
}
elseif ((Get-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name JPEGImportQuality).JPEGImportQuality -ne 100)
{
  Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name JPEGImportQuality -Value 100
}

# Check if the current wallpaper is identical to the current Spotlight image
if((Get-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallPaper).WallPaper -eq $latestImage)
{
  # We already have the current Spotlight image set as a wallpaper. Terminate the script with success code 0
  return 0
}

# Set the value of the desktop wallpaper to the value of the current Spotlight image
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallPaper -Value $latestImage

# Trigger a refresh of the desktop wallpaper. It won't trigger every time, so we need to loop it a few times to ensure the switch takes place. 60 should be more than enough
for($i = 0; $i -lt 60; $i++)
{
  & RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters ,1 ,True
  Start-Sleep -Seconds 1
}
