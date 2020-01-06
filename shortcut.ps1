# SET CUSTIOM DIR NAME
$TARGET_DIR_NAME = "_shortcuts"
# joinig path for new folder 
$target_path = Join-Path -Path $PWD -ChildPath "/$TARGET_DIR_NAME"
# if no such folder, create one
If(!(test-path $target_path))
{
      New-Item -ItemType Directory -Force -Path $target_path
}
# WScript Shell for creating a shortcut 
$WScriptShell = New-Object -ComObject WScript.Shell

# Enumerate all subdirs
$dir = dir $PWD | ?{$_.PSISContainer}
foreach ($d in $dir){
    # Try to find executable 1 lvl deep
    $name = Join-Path -Path $d.FullName -ChildPath "/*.exe"
    $exes = dir $name
    # For all found execuatles make a shortcut
    foreach($exe in $exes) {
      $dupe = $false
      # Extracting executable name for makign a shortcut and general logging
      $exename = Split-Path $exe -leaf
      # Making path where shortcut will be created     
      $shortcut_path = Join-Path -Path $target_path -ChildPath ("/$exename".replace("exe", "lnk"))
      # Checking if it allready exists
      If(Test-Path $shortcut_path)
      {
          # Setting dupe flag
          $dupe = $true
          # Extracting original dir name
          # https://stackoverflow.com/a/10318305
          $origin_dir = Split-Path (Split-Path $exe -Parent) -Leaf
          # Appending filename with dir name 
          $shortcut_path = $shortcut_path.Replace(".", "_$origin_dir.")
          # Logging
          Write-Host "[!] $exename alredy exists, creating with new name..." -ForegroundColor Yellow
      }
      # Create a shortcut
      # https://dotnet-helpers.com/powershell/create-shortcuts-on-desktops-using-powershell/
      $Shortcut = $WScriptShell.CreateShortcut($shortcut_path)
      $Shortcut.TargetPath = $exe
      $Shortcut.Save()
      # Selecting highlight color based if shortcut is dupe or not
      # Powershell inline if is soooooooooooo bad...
      # https://stackoverflow.com/a/25682508
      $color = (&{if($dupe) {"Yellow"} else {"Green"}}) 
      # Log a successfull shortcut creation
      Write-Host "[x] $exename" -ForegroundColor $color
    } 
}
