# Step 1: Prepare the environment to run the powershell script. Run this Script as the Domain logged user. This looks at their local outlook file in %appdata%
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force
$currentUsername = $env:USERNAME
$userDirs = Get-ChildItem -Path 'C:\Users\' -Directory
$matchFound = $false 
foreach ($userDir in $userDirs) {
    if($userDir.Name -eq $currentUsername){
        Write-Host "The directory for the logged user is: $($userDir.FullName)"
        $currentUserDir = $userDir.FullName
        $matchFound = $true
        break
    }
}
if(-not $matchFound) {
    Write-Host "There is no directory that matches the user logged in" #Make sure the user is logged-in and part of the domain. 
}

# Step 2: We will close outlook if its open and close all applications that take ownership of office processes.
$processNames = @('outlook', 'communicator', 'lync', 'ucmapi', 'msedge', 'msedgewebview2') # Kill ALL processes that use Office resources, you can check which ones with sysinternals process explorer
foreach ($processName in $processNames) {
	$process = Get-Process $processName -ErrorAction SilentlyContinue
	if($process) {
		Stop-Process -Name $processName -Force
	}
} 

$reg = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles" #Remove the profile from the local registry
try {
    $profiles = (Get-ChildItem -Path $reg -ErrorAction Stop).Name
} catch {
    Write-Host "Error: $_"
    exit
}
try {
    Remove-Item -Path registry::$profiles -Recurse -ErrorAction Stop
} catch {
	exit
}

$currentUsername = $env:USERNAME
$userDirs = Get-ChildItem -Path 'C:\Users\' -Directory
foreach ($userDir in $userDirs) {
    if($userDir.Name -eq $currentUsername){
        $ostDir = Join-Path -Path $userDir.FullName -ChildPath 'AppData\Local\Microsoft\Outlook\' #Start looking for OST/PST files locally on the PC
        if (Test-Path -Path $ostDir) {
			$ostFiles = Get-ChildItem -Path $ostDir -Recurse -Include "*@domain.com*.ost", "*@domain.com*.nst" #Look for the mailbox on the right format
            Write-Host "Total OST files: $($ostFiles.Count)"
            foreach ($ostFile in $ostFiles) {
                Write-Host "OST file: $($ostFile.FullName)" # delete all local profiles
            }
            foreach ($ostFile in $ostFiles) {
                try {
                    Remove-Item -Path $ostFile.FullName -Force 
                    Write-Host "Deleted OST file: $($ostFile.FullName)"
                    Start-Sleep -Seconds 5
                } catch {
                    Write-Host "Error deleting OST file: $($ostFile.FullName)"
                    Write-Host "Error details: $($_.Exception.Message)"
                    continue
                }
            }
        }
    }
}
$allFiles = Get-ChildItem -Path $ostDir -Recurse
Write-Host "Total files: $($allFiles.Count)"
foreach ($file in $allFiles) {
    Write-Host "File: $($file.FullName)"
}
# Setup a new profile 'outlook' for the current user logged in.
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover" -Name 'ZeroConfigExchange' -Value 1 -Type DWord
try {
    New-Item -Path $reg -Name 'outlook' -Force -ErrorAction Stop | Out-Null
} catch {
    Write-Host "Error: $_"
    exit
}
try {
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name DefaultProfile -Value "outlook" -ErrorAction Stop
} catch {
    Write-Host "Error: $_"
    exit
}
try {
    Start-Process 'outlook' -ErrorAction Stop -ArgumentList '/profile "outlook"'
} catch {
    Write-Host "Error: $_"
    exit
}

# Step 3: End of process, Set policy back for security
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Restricted -Force
