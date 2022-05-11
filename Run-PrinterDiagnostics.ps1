<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2021 v5.8.195
	 Created on:   	1/5/2022 8:55 AM
	 Created by:   	David Just
	 Filename: Run-PrinterDiagnostics 
	 Version: 1.0.5
	===========================================================================
	.DESCRIPTION
		Utility for basic printer diagnostics and troubleshooting
#>
####### Variables ##########
$PrintServiceLog = Get-WinEvent -Listlog "Microsoft-Windows-PrintService/Operational"
$printers = Get-Printer | Where-Object { $_.name -notlike "*Microsoft*" -and $_.name -notlike "*Fax*" -and $_.name -notlike "*OneNote*" }
$activeUser = (((quser | select-string 'active') -split ' ').replace('>', ''))[0]
$currentUser = whoami
$currentUser = ($currentUser -split '\\')[1]
############################

###### Functions ##########

function Welcome
{
	
	Write-Host "####################################################################################################"
	Write-Host "Run-PrintDiagNostics Version 1.0.5`nAuthor: David Just" -ForegroundColor DarkCyan -BackgroundColor Black
	Write-Host "Welcome to the Printer Diagnostics Utility `nWhere we try to make printers slightly less painful! `nCurrently Running as $($currentUser.toUpper())" -ForegroundColor Green -BackgroundColor Black
	Write-Host "#################################################################################################### `r"
	if (($activeUser -ne $currentUser))
	{
		Write-Warning "Caution, you are currently running as $($currentUser.toUpper()).`nYou must run this tool as the logged on user in order to work with shared printers`nRunning as SYSTEM or another user will only show system wide printers!"
	}
	
}

function FetchPrintLogs
{
	If (!(Get-WinEvent -Listlog "Microsoft-Windows-PrintService/Operational" | select -ExpandProperty IsEnabled))
	{
		Write-Host "System Log [Microsoft-Windows-PrintService/Operational] is disabled...Trying to enable now" -ForegroundColor Yellow -BackgroundColor Black
		sleep 1
		$scriptblock = @'
$log=Get-WinEvent -ListLog 'Microsoft-Windows-PrintService/Operational';$log.IsEnabled=$true;$log.SaveChanges()
'@
		
		$bytes = [System.Text.Encoding]::Unicode.GetBytes($scriptblock)
		$encodedCommand = [Convert]::ToBase64String($bytes)
		Start-Process "Powershell.exe" -ArgumentList "-EncodedCommand $encodedCommand" -Verb runas
		sleep 2
		if ((Get-WinEvent -Listlog "Microsoft-Windows-PrintService/Operational" | select -ExpandProperty IsEnabled))
		{
			Write-Host "Successfully Enabled Logging!`nPlease try to print a test page and check back for event logs.`n" -ForegroundColor Green -BackgroundColor Black
			pause
			Clear-Host
		}
		else
		{
			Write-Host "Failed to enable Print Service Operation Log!"
			sleep 2
			Clear-Host
		}
		
	}
	else
	{
		Get-WinEvent -LogName "Microsoft-Windows-PrintService/Operational" | Out-GridView
		Clear-Host
	}
	
}

function RestartSpooler #Clears and restarts print spooler service.
{
	Write-Host "Restarting Print Spooler..." -ForegroundColor White -BackgroundColor Black
	$scriptblock = @'
Get-Service Spooler | Stop-Service -force ; cmd /c "del %systemroot%\System32\spool\printers* /Q" ; Get-Service Spooler | Start-Service
'@
	$bytes = [System.Text.Encoding]::Unicode.GetBytes($scriptblock)
	$encodedCommand = [Convert]::ToBase64String($bytes)
	Start-Process "Powershell.exe" -ArgumentList "-EncodedCommand $encodedCommand" -Verb runas
	if ((Get-Service Spooler | select -ExpandProperty Status) -like "Running")
	{
		Write-Host "Successfully Restarted Spooler!" -ForegroundColor Green -BackgroundColor Black
		$Prompt = Read-Host "Would you like to print a test page? Y/N"
		switch ($Prompt)
		{
			Y{ PrintTestPage }
			N{ }
			
		}
		sleep 1
		Clear-Host
	}
	else
	{
		Write-Warning "Spooler service is not running!"
		sleep 2
		Clear-Host
	}
}

function PrintTestPage
{
	
	$index = 1
	foreach ($printer in $printers)
	{
		
		Write-Host "[$index] $($printer.name)"
		$index++
		
	}
	
	try
	{
		$Selection = Read-Host "Select which printer to send a test page to. Press Q to return to mainmenu" -ErrorAction 'Stop'
		if ($Selection -like "Q") { return }
		else {$Selection = [int]$Selection}
		
	}
	Catch
	{
		Write-Host "Please enter a valid menu selection." -ForegroundColor Red -BackgroundColor Black
		Pause
		PrintTestPage
	}
	$SelectedPrinter = $printers[$Selection - 1]
	$TestPrint = Invoke-CimMethod -MethodName printtestpage -InputObject (Get-CimInstance win32_printer | Where-Object { $_.name -like $SelectedPrinter.Name })
	if ($TestPrint.ReturnValue -eq 0)
	{
		Write-Host "Test Page Sent Successfully! `n" -ForegroundColor Green -BackgroundColor Green
		sleep 1
		Clear-Host
	}
	else
	{
		Write-Host "Test page failed!" -ForegroundColor Red -BackgroundColor Green
		sleep 1
	}
	
	
}

function OpenPrinterCP
{
	param (
	[String]$Option	
	)
	switch ($Option)
	{
		Modern {
			Start-Process ms-settings:printers
			Clear-Host
		}
		Classic {
			Start-Process control printers
			Clear-Host
		}
	}
	
}



function MainMenu
{
	$log = Get-WinEvent -Listlog "Microsoft-Windows-PrintService/Operational" | select -ExpandProperty IsEnabled
	if ($log) { $status = 'Enabled' }
	else { $status = 'Disabled' }
	Write-Host "Printer Service Log Status: $($status)" -Foregroundcolor DarkCyan -BackGroundColor Black
	Write-Host "Installed Printers:"
	$printers | select Name, Portname, Drivername | Format-Table -AutoSize
	Write-Host "Main Menu"
	Write-Host "Options: `n[1] Review Print Service Logs`n[2] Restart and Clear Print Spooler`n[3] Add a new printer`n[4] Open Classic Printer Control Panel`n[5] Open Modern Printer Settings Page`n[6] Print Test Page`n[7] Quit" -ForegroundColor Cyan -BackgroundColor Black
	try
	{
		[int]$MenuSelection = Read-Host "Please Enter A Selection" -ErrorAction 'Stop'
	}
	Catch
	{		
		Write-Host "Please select a menu option 1 - 7" -ForegroundColor Red -BackgroundColor Black
		Pause
		Clear-Host		
	}
	switch ($MenuSelection)
	{
		1{ FetchPrintLogs } #Review Print Logs
		2{ RestartSpooler } #Restart and Clear Spooler
		3{ cmd /c "C:\Windows\System32\rundll32.exe PRINTUI.DLL, PrintUIEntry /im"; return } #Add a new printer
		4{ OpenPrinterCP -Option Classic }
		5{ OpenPrinterCP -Option Modern }
		6{ PrintTestPage }
		7{ Write-Host "Thank you for using the Print Diagnostic Tool. Happy Printing!"; sleep 1; exit } #Quit
	}
		
}

Welcome
sleep 1
while ($true){
MainMenu
}