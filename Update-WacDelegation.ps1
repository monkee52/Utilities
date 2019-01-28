# Allow WAC delegation
$wacServer = Get-ADComputer -Identity "MyWacServer"

$wacManagedMachines = (Get-ADGroupMember -Identity "ShadowGroup_Servers" -Recursive) + (Get-ADGroupMember -Identity "Domain Controllers")

$wacManagedMachines | % {
	$_ | Set-ADComputer -PrincipalsAllowedToDelegateToAccount $wacServer
}
