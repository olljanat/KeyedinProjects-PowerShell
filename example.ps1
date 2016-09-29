Import-Module .\KeyedinProjects.psm1

# Do login
$AuthenticationToken = Invoke-KeyedinLogin -UserName "john.doe@example.com" -Password "PASSWORD"

# Get data
$SystemConfig = Get-KeyedinSystemConfig
$Projects = Get-KeyedinProjects
$Activities = Get-KeyedinActivities
$Tasks = Get-KeyedinTasks -StartDate $(Get-Date -Date "2016-01-01") -EndDate $(Get-Date -Date "2016-12-31") -ResourceCode $SystemConfig.ResourceCode
$Timesheets = Get-KeyedinTimesheets -StartDate $(Get-Date -Date "2016-01-01") -EndDate $(Get-Date -Date "2016-12-31") -ResourceCode $SystemConfig.ResourceCode


# Find your activity code
$Activity = $Activities | Where-Object {$_.Description -like "*abc"}

# Create project + task list of active projects/tasks and export it to CSV
# You can open that list to Excel and find project codes and task keys from it (you need them when you add new timesheets).
$ActiveProjects = $Projects | Where-Object {$_.IsActive -eq $True} | Sort-Object Code
ForEach ($Project in $ActiveProjects) {
	$ActiveProjectTasks = $Tasks | Where-Object {($_.IsActive -eq $True) -and ($_.ProjectCode -eq $Project.Code)}
	ForEach ($Task in $ActiveProjectTasks) {
		[array]$ProjectTaskArray += New-Object -TypeName PSObject -Property @{
			"ProjectCode" = $Project.Code
			"ProjectDescription" = $Project.Description
			"TaskKey" = $Task.Key
			"TaskDescription" = $Task.Description
		}
	}
}
$ProjectTaskArray | Select-Object ProjectCode,ProjectDescription,TaskKey,TaskDescription | Export-Csv .\ProjectTaskArray.csv -Delimiter ";" -Encoding UTF8 -NoTypeInformation

# Add timesheet
Add-KeyedinTimesheet -ActivityCode "123" -TaskKey "12345" -HoursDecimal 7.5 -IsChargeable $False -IsOvertime $False -ProjectCode "54321" -TimesheetDate $(Get-Date -Date "2016-01-01")

# Do logout
Invoke-KeyedinLogout