# Allow WAC delegation
$wacServer = Get-ADComputer -Identity "MyWacServer"

$wacManagedMachines = (Get-ADGroupMember -Identity "ShadowGroup_Servers" -Recursive) + (Get-ADGroupMember -Identity "Domain Controllers") | ? { ($_ | Get-ADComputer -Properties OperatingSystem).OperatingSystem.StartsWith("Windows Server") }

$wacManagedMachines | % {
	Write-Host "Allow remote administration for" $_.Name
	
	$_ | Set-ADComputer -PrincipalsAllowedToDelegateToAccount $wacServer
}
