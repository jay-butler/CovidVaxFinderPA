Function Get-AllentownAppt
{
<#
.SYNOPSIS
	Loads the Allenttown PA vaccine site's appointment page periodically. If there is an appointment available, the page will be loaded in a browser and there is an optional audbible alert. Audio does not work in PowerShell Core.
.DESCRIPTION
	The appoinment page displays the days on which there are available appointments. Clicking one of these links goes to a page to select a time. Once a time is selected, there is a form with name, address, date of birth, etc. that must be completed and submitted. To speed the form filling, use your password manager's or your browser's functionality. 
.NOTES
	Author: Jay Butler.
.PARAMETER CycleSeconds
	Number of seconds between page reloads. Set this number low enough to find the appointment availability quickly, but not so low that the script continually reloads the website. 
.PARAMETER LogFilePath
	The results will be written to a file. Specify a folder for the log file if you want to place it anywhere other than the default of a folder named "CovidVaxFinder" within your home folder.
.PARAMETER SilentMode
	Specify this to suppress the audible alerts.
.EXAMPLE
	Invoke-ExDbaAzDataFileBlobStorageConfiguration -TargetSQLInstances 'sspc1dbd-001,51818' -Environment AzureCloud;
#>

[cmdletbinding()]

	PARAM
	(
	[Parameter(Mandatory=$false)]
	[ValidateRange(30,3600)]
		[int]$CycleSeconds = 30,
	[Parameter(Mandatory=$false)]
	[ValidateScript({Test-Path $_ -PathType ‘Container’})] 
		[string]$LogFilePath = ('{0}\CovidVaxFinder\' -f $Home),
	[Parameter(Mandatory=$false)]
		[switch]$SilentMode = $false
	)

	BEGIN
	{
		$PSBoundParameters.GetEnumerator() | ForEach-Object { Write-Verbose $_ };

		if ($LogFilePath[-1] -ne '\') 
			{$LogFilePath += '\';}

		[string]$OutFile = ('{0}AllentownAppts_{1}.txt' -f $LogFilePath, (((Get-Date -Format s) -replace 'T','_') -replace ':','-'))
		Write-Host ('Log file: {0}' -f $OutFile );
	}

	PROCESS
	{
		[boolean]$WindowOpen = $false;

		# Not the landing page, but the next page where days with available appointments are listed.
		[string]$PageURL = 'https://allentownpaclinics.schedulemeappointments.com/?mode=startreg&b2ZmZXJpbmdpZD00MjYx';
		# This is the magic that will find the information about availabel appointments (if there are any)...
		[regex]$match = [Regex]::New('((?<ApptCount>\d?[,]?\d{1,3})(?<ApptLabel> appointment[s]? remaining))');

		While ($true)
		{
			$StartCycleTime = Get-Date;
			Write-Debug ('{0}	Still alive' -f ((Get-Date -Format s) -replace 'T', ' '));

			# Get the scheduler page
			$page = Invoke-WebRequest -Uri $PageURL -SessionVariable ACF_sess;
			$Links = $page.Links | Where-Object {$_.Class -like 'entry calendar-entry-link calendar-entry-link-offering*'};

			# This is some older code used to detect appointents. The regex expression above and evaluation
			# below replaced this. But, the page does chaneg from time to time, so maybe these conditions 
			# will be needed again.
			<#
			# If this string is missing, there are appointments
			$NoApptMsg = $page.Content.IndexOf(' There are no openings available for this offering <em>(623)</em>.');
			if ($NoApptMsg -le 0) {$NoApptMsg = $page.Content.IndexOf('<span class="BDHerrorinfo">We are sorry but the clinic you have chosen is full.</span>');}
			if ($NoApptMsg -le 0) {$NoApptMsg = $page.Content.IndexOf('<span class="BDHerroremphasis">Error:</span>');}
			if ($NoApptMsg -le 0) {$NoApptMsg = $page.Content.IndexOf('Sorry but there are no dates available at this time');}
			$ApptMsg = $page.Content.IndexOf('M»&nbsp;1st Dose of Covid Vaccine');
			#>

			# Look for the link that schedules the appointments
			if ($Links)
			{
				# Regex search to find the appointment numbers for every day there are appointments...
				$AllAppts = $match.Matches($page.Content);
				$AppCount = ($AllAppts | ForEach-Object {$_.Groups | Where-Object {$_.Name -eq 'ApptCount'}}).Value;
				$AppSum = $AppCount | Measure-object -Sum;

				[string]$message = ('***** {0:N0} appointment(s) on {1:N0} days in Allentown ({2})' -f $AppSum.Sum, $AppSum.Count, ([string[]]$AppCount -join ', ') );
				$MsgColor = 'Green';

				# if( $page.Content -match '(?<NbrAppts>\d?[,]?\d+)( appointment[s]? remaining)')
				# 	{ $NbrAppts = [int]$Matches.NbrAppts; }
				# [string]$message = ('{0}	***** {1:N0} appointment(s) in Allentown available' -f ((Get-Date -Format s) -replace 'T', ' '),$NbrAppts );

				if (-not $WindowOpen)
				{
					# Open the page in a browser ASAP
					Start-Process $PageURL;
					if (-not $SilentMode)
						{
						# Play an audible alert
						for ($i=0; $i -lt 2; $i++) {[console]::beep(550,220)};
						# Chirp a few times if there are a lot of appointments
						for ($i=0; $i -lt ($AppSum.Sum / 50); $i++) {[console]::beep(750,120)};
						} #end IF silent
					# Then, remember some stuff
					[boolean]$WindowOpen = $true;
					[datetime]$WindowStart = Get-Date;
				} # End Nested If window open/closed

			} # End If appointments present
			else
			{
				if ($WindowOpen)
				{
					if (-not $SilentMode)
						{
						[console]::beep(900,750);
						} #End IF silent
					[boolean]$WindowOpen = $false;
					[string]$message = ('{0}	------> Window duration was {1:N2} seconds.' -f ((Get-Date -Format s) -replace 'T', ' '), $((New-Timespan -Start $WindowStart).TotalSeconds));
					Write-Host $message  -ForegroundColor White -BackgroundColor DarkBlue;
					Add-Content -Path $OutFile -Value $message;
				} # End Nested If
				[string]$message = ('No appointments in Allentown');
				$MsgColor = 'Yellow';
			} # End Else

			Add-Content -Path $OutFile -Value $message;
			Write-Host ('{0}	' -f ((Get-Date -Format s) -replace 'T', ' ')) -ForegroundColor White -NoNewline;
			Write-Host $message -ForegroundColor $MsgColor;

			# Run exactly at the cylce time, so subtract the amountof time it took to do all the stuff above...
			Start-Sleep -Milliseconds (($CycleSeconds * 1000) - (New-Timespan -Start $StartCycleTime).TotallMilliseconds);
		} # End While True
	} # End PROCESS

	END
	{
	} # End END

} # End FUNCTION

Get-AllentownAppt -CycleSeconds 30;