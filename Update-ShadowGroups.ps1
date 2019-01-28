# BEGIN CONFIG

$rootOU = "OU=MyContainer,DC=MyDomain"
$shadowGroupOU = "OU=ShadowGroups,OU=Groups,OU=MyContainer,DC=MyDomain"

$shadowGroupNameFormat = "ShadowGroup_{0}"

$exclude = @(
	"OU=Groups,OU=MyContainer,DC=MyDomain",
	"OU=NewUsers,OU=MyContainer,DC=MyDomain",
	"OU=NewComputers,OU=MyContainer,DC=MyDomain"
)

$aliases = @{
	"ShadowGroup_Users" = "My_Users";
	"ShadowGroup_Servers" = "My_Servers";
	"ShadowGroup_Hypervisors" = "My_Hypervisors";
}

# END CONFIG

$exclude += $shadowGroupOU

function Recurse-OU {
	param(
		[String]$DistinguishedName
	)
	
	[Array]$results = @()
	
	Get-ADOrganizationalUnit -Filter * -SearchBase $DistinguishedName -SearchScope OneLevel | Sort-Object -Property DistinguishedName | % {
		$results += $_
		$results += Recurse-OU -DistinguishedName $_.DistinguishedName
	}
	
	return $results
}

$ouList = Recurse-OU -DistinguishedName $rootOU

# Depth first
[Array]::Reverse($ouList)

# Pass 1 - create groups
$ouList | ? { $exclude -notcontains $_.DistinguishedName } | % {
	$name = $shadowGroupNameFormat -f $_.Name
	
	try {
		$group = Get-ADGroup -Identity $name
		Write-Host $name "exists"
	} catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
		$group = New-ADGroup -Name $name -Path $shadowGroupOU -GroupScope Global -GroupCategory Security
		Write-Host $name "created"
	}
}

# Pass 2 - modify groups (mark then sweep?)
$ouList | ? { $exclude -notcontains $_.DistinguishedName } | % {
	$name = $shadowGroupNameFormat -f $_.Name
	$group = Get-ADGroup -Identity $name
	
	Write-Host "Processing" $name
	
	$members = @{}
	
	# Get candidate computers
	Get-ADComputer -Filter * -SearchBase $_.DistinguishedName -SearchScope OneLevel | % {
		$members.Add($_.DistinguishedName, @{ Object = $_; Action = "add" })
	}
	
	# Get candidate users
	Get-ADUser -Filter * -SearchBase $_.DistinguishedName -SearchScope OneLevel | % {
		$members.Add($_.DistinguishedName, @{ Object = $_; Action = "add" })
	}
	
	# Get sub shadow groups
	Get-ADOrganizationalUnit -Filter * -SearchBase $_.DistinguishedName -SearchScope OneLevel | ? { $exclude -notcontains $_.DistinguishedName } | % {
		$subName = $shadowGroupNameFormat -f $_.Name
		$subGroup = Get-ADGroup -Identity $subName
		
		$members.Add($subGroup.DistinguishedName, @{ Object = $subGroup; Action = "add" })
	}
	
	# Determine action for existing shadow group members (none/keep, or remove)
	$group | Get-ADGroupMember | % {
		if ($members.ContainsKey($_.DistinguishedName)) {
			$members[$_.DistinguishedName].Action = "none"
		} else {
			$members.Add($_.DistinguishedName, @{ Object = $_; Action = "remove" })
		}
	}
	
	# Process actions
	$members.Values | % {
		if ($_.Action -eq "add") {
			$group | Add-ADGroupMember -Members $_.Object
			Write-Host "Adding" $_.Object.Name "to" $name
		} elseif ($_.Action -eq "remove") {
			$group | Remove-ADGroupMember -Members $_.Object
			Write-Host "Removing" $_.Object.Name "from" $name
		} elseif ($_.Action -eq "none") {
			Write-Host "Keeping" $_.Object.Name "in" $name
		}
	}
}

# Pass 3 - Modify alias groups (mark then sweep?)
$aliases.Keys | % {
	$group = Get-ADGroup -Identity $_
	$destGroup = Get-ADGroup -Identity $aliases[$_]
	
	Write-Host "Processing alias from" $group.Name "to" $destGroup.Name

	$members = @{}
	
	# Get shadow group members
	$group | Get-ADGroupMember | % {
		$members.Add($_.DistinguishedName, @{ Object = $_; Action = "add" })
	}
	
	# Get target group members
	$destGroup | Get-ADGroupMember | % {
		if ($members.ContainsKey($_.DistinguishedName)) {
			$members[$_.DistinguishedName].Action = "none"
		} else {
			$members.Add($_.DistinguishedName, @{ Object = $_; Action = "remove" })
		}
	}
	
	# Process actions
	$members.Values | % {
		if ($_.Action -eq "add") {
			$destGroup | Add-ADGroupMember -Members $_.Object
			Write-Host "Adding" $_.Object.Name "to" $destGroup.Name
		} elseif ($_.Action -eq "remove") {
			Write-Host "Removing" $_.Object.Name "from" $destGroup.Name
		} elseif ($_.Action -eq "none") {
			Write-Host "Keeping" $_.Object.Name "in" $destGroup.Name
		}
	}
}
