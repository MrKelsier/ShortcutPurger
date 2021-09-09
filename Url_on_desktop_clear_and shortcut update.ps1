#TIMEOUT=180000
#!PS
#kesstn 2021

$ErrorActionPreference = 'SilentlyContinue'

$DELshortcuts = "Shortcutname1","shortcutname2","shortcutname3" #list the shortcuts you want deleted from the desktop here
$NEWShortcutName = "MyRandomShorcut" #The name of the shortcut
$NEWShortcutAddress = "http://imthebest.com" #Website address




$users = Get-ChildItem -Path "C:\Users"
$users | ForEach-Object {
    foreach ($DELShortcut in $DELshortcuts){
        if (Test-Path "C:\Users\$($_.Name)\Desktop\$DELShortcut"){
            Remove-Item -Path "C:\Users\$($_.Name)\Desktop\$DELShortcut"
                write-host $_.name "has old $DELShortcut, Removing"
        }
    }
}


$Shell = New-Object -ComObject ("WScript.Shell")
$alldesktops =  [Environment]::GetFolderPath("CommonDesktopDirectory")
$ShortCut = $Shell.CreateShortcut($alldesktops + "\" + $NEWShortcutName)
if (test-path "C:\Program Files (x86)\Internet Explorer\iexplore.exe"){
    $ShortCut.TargetPath = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
    $ShortCut.WorkingDirectory = "C:\Program Files (x86)\Internet Explorer"
}
else {
    $ShortCut.TargetPath = "C:\Program Files\Internet Explorer\iexplore.exe"
    $ShortCut.WorkingDirectory = "C:\Program Files\Internet Explorer"
    
}
$ShortCut.Arguments = $NEWShortcutAddress
$ShortCut.WindowStyle = 1
$ShortCut.IconLocation = "iexplore.exe, 0"

$ShortCut.Save()
write-host "$NEWShortcutName created"


