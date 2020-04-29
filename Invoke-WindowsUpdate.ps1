$criteria = "IsInstalled=0 AND IsHidden=0"

function Get-UpdateDescription {
	param(
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Position = 0)]
		$Update,
		[Parameter(Mandatory = $false)]
		[bool]$NoCategories = $false
	)

	$description = $Update.Title
	
	if ($Update.KBArticleIDs.Count -gt 0) {
		$description += " ("
		
		for ($j = 0; $j -lt $Update.KBArticleIDs.Count; $j += 1) {
			if ($j -gt 0) {
				$description += ","
			}
			
			$description += "KB" + $Update.KBArticleIDs.Item($j)
		}
		
		$description += ")"
	}

	if (-not $NoCategories) {
		$description += " Categories: "
	
		for ($j = 0; $j -lt $Update.Categories.Count; $j++) {
			$category = $Update.Categories.Item($j)
		
			if ($j -gt 0) {
				$description += ","
			}
		
			$description += $category.Name
		}
	}

	return $description
}

enum OperationResultCode {
	NotStarted
	InProgress
	Succeeded
	SucceededWithErrors
	Failed
	Aborted
}

enum InstallationRebootBehavior {
	NeverReboots
	AlwaysRequiresReboot
	CanRequestReboot
}

$session = New-Object -ComObject Microsoft.Update.Session
$searcher = $session.CreateUpdateSearcher()

$serviceManager = New-Object -ComObject Microsoft.Update.ServiceManager

# Install Microsoft Update
$_ = $serviceManager.AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 7, $null)

foreach ($service in $serviceManager.Services) {
	if ($service.Name -eq "Microsoft Update") {
		$searcher.ServerSelection = 3
		$searcher.ServiceID = $service.ServiceID

		break
	}
}

Write-Host "Searching for updates..."

$searchResult = $searcher.Search($critera).Updates

$selectedUpdates = New-Object -ComObject Microsoft.Update.UpdateColl

$yesToAllFlag = $false
$noToAllFlag = $false
$suspendFlag = $false
$exclusiveAdded = $false

$yesToAllEulaFlag = $false
$noToAllEulaFlag = $false

$optYes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Install the update"
$optYesToAll = New-Object System.Management.Automation.Host.ChoiceDescription "Yes to &All", "Installs the update and all subsequent updates"
$optNo = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Does not install the update"
$optNoToAll = New-Object System.Management.Automation.Host.ChoiceDescription "No to A&ll", "Does not install the update and all subsequent updates"
$optSuspend = New-Object System.Management.Automation.Host.ChoiceDescription "&Suspend", "Exit the updater"

$options = [System.Management.Automation.Host.ChoiceDescription[]]($optYes, $optYesToAll, $optNo, $optNoToAll, $optSuspend)
$optionsYesNo = [System.Management.Automation.Host.ChoiceDescription[]]($optYes, $optNo)

for ($i = 0; $i -lt $searchResult.Count; $i += 1) {
	if ($suspendFlag) {
		break
	}

	if ($exclusiveAdded) {
		continue
	}
	
	$update = $searchResult.Item($i)
	
	if ($yesToAllFlag) {
		$_ = $selectedUpdates.Add($update)
		continue
	}
	
	if ($noToAllFlag) {
		continue
	}
	
	# Get description
	$description = $Update | Get-UpdateDescription
	
	# Check exclusivity
	if ((($update.InstallationBehavior.Impact -eq 2) -and ($selectedUpdates.Count -gt 0)) -or $exclusiveAdded) {
		continue
	}

	switch ([InstallationRebootBehavior]$update.InstallationBehavior.RebootBehavior) {
		([InstallationRebootBehavior]::CanRequestReboot) {
			Write-Warning "The below update may require a restart once installed."
		}
		([InstallationRebootBehavior]::AlwaysRequiresReboot) {
			Write-Warning "The below update will require a restart once installed."
		}
	}

	$result = $Host.UI.PromptForChoice("Do you want to install this update?", $description, $options, 0)
	$add = $false
	
	switch ($result) {
		0 {
			$add = $true
		}
		1 {
			$add = $true
			$yesToAllFlag = $true
		}
		2 {
			continue
		}
		3 {
			$noToAllFlag = $true
			continue
		}
		4 {
			$suspendFlag = $true
		}
	}

	#TODO: Add EULA handling

	if ($add) {
		$_ = $selectedUpdates.Add($update)

		if ($update.InstallationBehavior.Impact -eq 2) {
			$exclusiveAdded = $true
		}
	}
}

if (-not $suspendFlag -and $selectedUpdates.Count -gt 0) {
	$downloader = $session.CreateUpdateDownloader()
	$downloader.Updates = $selectedUpdates
	
	Write-Host "Downloading updates..."

	$_ = $downloader.Download()
	
	$installer = New-Object -ComObject Microsoft.Update.Installer
	$installer.Updates = $selectedUpdates
	
	Write-Host "Installing updates..."

	$result = $Installer.Install()

	Write-Host ("Result: {0}" -f ([OperationResultCode]$result.ResultCode))

	for ($i = 0; $i -lt $selectedUpdates.Count; $i += 1) {
		$description = $selectedUpdates.Item($i) | Get-UpdateDescription
		$updateResult = $result.GetUpdateResult($i)
		$updateResultCode = [OperationResultCode]$updateResult.ResultCode

		Write-Host ("{0}: {1}" -f $description, $updateResultCode)

		if ($updateResultCode -ge [OperationResultCode]::SucceededWithErrors) {
			if ($updateResult.HResult -eq -2145116147) {
				Write-Warning "This update needed additional downloaded content. Please re-run this program."
			} else {
				Write-Warning ("HRESULT 0x{0:x8}" -f $updateResult.HResult)
			}
		}
	}

	if ($result.RebootRequired) {
		switch ($Host.UI.PromptForChoice("A restart is required to finish installing updates.", "Do you want to restart now?", $optionsYesNo, 0)) {
			0 {
				Restart-Computer
			}
			1 {
		
			}
		}
	}
}
