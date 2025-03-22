# Kill tasks
taskkill.exe /IM winword.exe /F 
taskkill.exe /IM outlook.exe /F 
taskkill.exe /IM excel.exe /F 
taskkill.exe /IM powerpnt.exe /F 
taskkill.exe /IM wbgx.exe /F 

# Remote C:\Worldox 
Remove-Item -Recurse C:\Worldox -Verbose -ErrorAction SilentlyContinue

# Remove AppData files for each user
$users = Get-ChildItem c:\users
foreach ($user in $users)
{
    $userdir = $user.name
   <# takeown /r /d y /f c:\users\$userdir\AppData\Local\Worldox 
    takeown /r /d y /f c:\users\$userdir\AppData\Roaming\Worldox
    takeown /r /d y /f c:\users\$userdir\AppData\Roaming\Microsoft\Word
    takeown /r /d y /f c:\users\$userdir\AppData\Roaming\Microsoft\Excel
    #>
   
   # Remove from AppData and recurse through all users
    Remove-Item -Recurse -Force c:\users\$userdir\AppData\Local\Worldox -Verbose -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force c:\users\$userdir\AppData\Roaming\Worldox -Verbose -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force c:\users\$userdir\AppData\Roaming\Microsoft\Excel\XLSTART\@WD*.* -Verbose -ErrorAction SilentlyContinue
    Remove-Item -Recurse -Force c:\users\$userdir\AppData\Roaming\Microsoft\Word\Startup\@WD*.* -Verbose 
    
   
}

# Recurse through all HKCU registry keys and delete Worldox Addins from each app

# Define the relative path to the Add-In key
# Define an array of Add-In registry subkeys to delete
$addInSubKeys = @(
    "Software\Microsoft\Office\Outlook\Addins\W6OLCAI2010.Connect",
    "Software\Microsoft\Office\Outlook\Addins\WDOLCAI2010.Connect"
)

# Get all user SIDs under HKEY_USERS
$hku = [Microsoft.Win32.Registry]::Users
$subKeys = $hku.GetSubKeyNames()

foreach ($sid in $subKeys) {
    foreach ($addInSubKey in $addInSubKeys) {
        try {
            $fullKeyPath = "$sid\$addInSubKey"
            $regPath = "Registry::HKEY_USERS\$fullKeyPath"
            
            if (Test-Path $regPath) {
                Remove-Item -Path $regPath -Recurse -Force
                Write-Output "Deleted: $regPath"
            } else {
                Write-Output "Key not found: $regPath"
            }
        } catch {
            Write-Warning "Failed to delete key '$addInSubKey' for SID $sid : $_"
        }
    }
}
