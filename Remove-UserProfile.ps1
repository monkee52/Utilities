Set-Location (Join-Path $env:SystemDrive "Users")

$keepGoing = $true

while ($keepGoing) {
    "------------------------------------------------"
    
    $profiles = @()
    
    $profiles += Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" | ? {
        $sid = New-Object System.Security.Principal.SecurityIdentifier ($_.PSChildName)
        $user = $sid.Translate([System.Security.Principal.NTAccount])
        
        -not (($sid.ToString() -eq "S-1-5-18") -or ($sid.ToString() -eq "S-1-5-19") -or ($sid.ToString() -eq "S-1-5-20"))
    } | % { New-Object PSObject -Property @{ User = $user.Value; SID = $sid.ToString(); Path = ($_ | Get-ItemProperty -Name "ProfileImagePath").ProfileImagePath } }
    
    for ($i = 0; $i -lt $profiles.Length; $i++) {
        "[" + ($i + 1).ToString() + "] " + $profiles[$i].User + " (" + $profiles[$i].SID + ") = " + $profiles[$i].Path
    }
    
    $keepGoing2 = $true
    $finalOption = $null
    
    while ($keepGoing2) {
        $opt = Read-Host "Profile"
        
        if ($opt -eq "") {
            $keepGoing = $false
            break
        }
        
        try {
            $opt = [Convert]::ToInt32($opt, 10)
        } catch {

        }
        
        if ($opt -gt 0 -and $opt -le $profiles.Length) {
            $keepGoing2 = $false
            $finalOption = $opt - 1
        }
    }
    
    if ($finalOption -ne $null) {
        $profile = $profiles[$finalOption]
        
        Remove-Item ("HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\" + $profile.SID)
        Remove-Item -Recurse -Force $profile.Path
    }
}
