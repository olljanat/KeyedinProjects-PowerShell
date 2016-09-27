# Global parameters for all functions
$AppVersion = "2.4.1"
$DeviceID = ([guid]::NewGuid()).Guid
$global:URI=""
$global:RequestHeader = @"
<Envelope xmlns:key="http://www.keyedin.com/" xmlns="http://schemas.xmlsoap.org/soap/envelope/">
  <Header>
    <key:AuthenticationHeader>
      <key:AuthenticationToken></key:AuthenticationToken>
      <key:DeviceID>$DeviceID</key:DeviceID>
      <key:AppVersion>$AppVersion</key:AppVersion>
    </key:AuthenticationHeader>
  </Header>
  <Body>
"@
$global:RequestFooter = "</Body></Envelope>"

Function Invoke-Login {
    <#
    .SYNOPSIS
        Invoke login to Keyedin Projects mobile API

    .PARAMETER UserName
        Your Keyedin Projects username. Usually on email format.

    .PARAMETER Password
        Your Keyedin Projects password
		
    .PARAMETER Server
        Your Keyedin Projects server. Options are "Europe", "USA" and your own installation server name with http:// or https:// -prefix.
		
    .EXAMPLE
        Invoke-Login -UserName "john.doe@example.com" -Password "PASSWORD" -Server "Europe"
    #>
	param (
		[Parameter(Mandatory=$true)][string]$UserName,
		[Parameter(Mandatory=$true)][string]$Password,
		[Parameter(Mandatory=$true)][string]$Server
	)
	
	If ($Server -eq "Europe") {
		$global:URI = "https://mobile.keyedinprojects.co.uk/mobileservices.svc"
	} ElseIf ($Server -eq "USA") {
		$global:URI = "https://mobile.keyedinprojects.com/mobileservices.svc"
	} Else {
		$global:URI = $Server + "/mobileservices.svc"
	}
	
	$LoginRequest = "<key:Login><key:username>" + $UserName + "</key:username><key:password>" + $Password + "</key:password></key:Login>"
	$Request = $RequestHeader + $LoginRequest + $RequestFooter
	
	$Headers = @{"SOAPAction" = "http://www.keyedin.com/IProjectMobileService/Login"}
	$Result = (Invoke-WebRequest -Uri $URI -ContentType "text/xml" -Headers $Headers -Method POST -Body $Request -UserAgent "")
	
	[xml]$ResultXML = $Result.Content
	If ($ResultXML.Envelope.Body.LoginResponse.LoginResult.Token.Length -ne 172) {
		throw "Login failed. Check username and password"
	} Else {
		$AuthenticationToken = $ResultXML.Envelope.Body.LoginResponse.LoginResult.Token
		$global:RequestHeader = $RequestHeader -replace "<key:AuthenticationToken></key:AuthenticationToken>","<key:AuthenticationToken>$AuthenticationToken</key:AuthenticationToken>"
	}
}

Function Invoke-Logout {
    <#
    .SYNOPSIS
        Invoke logout from Keyedin Projects mobile API

    .EXAMPLE
        Invoke-Logout
    #>
	$LogoutRequest = "<key:Logout />"
	$Request = $RequestHeader + $LogoutRequest + $RequestFooter
	
	$Headers = @{"SOAPAction" = "http://www.keyedin.com/IProjectMobileService/Logout"}
	$Result = (Invoke-WebRequest -Uri $URI -ContentType "text/xml" -Headers $Headers -Method POST -Body $Request -UserAgent "")
	
	[xml]$ResultXML = $Result.Content
	return $ResultXML.Envelope.Body.LogoutResponse.LogoutResult.Succeeded
}

Function Get-SystemConfig {
    <#
    .SYNOPSIS
        Get your Keyedin Projects settings.

    .EXAMPLE
        Get-SystemConfig
    #>
	$SystemConfigRequest = '<key:GetSystemConfig />'
	$Request = $RequestHeader + $SystemConfigRequest + $RequestFooter
	
	$Headers = @{"SOAPAction" = "http://www.keyedin.com/IProjectMobileService/GetSystemConfig"}
	$Result = (Invoke-WebRequest -Uri $URI -ContentType "text/xml" -Headers $Headers -Method POST -Body $Request -UserAgent "")
	
	[xml]$ResultXML = $Result.Content
	return $ResultXML.Envelope.Body.GetSystemConfigResponse.GetSystemConfigResult
}

Function Get-Projects {
    <#
    .SYNOPSIS
        Get your projects.

    .EXAMPLE
        Get-Projects
    #>
	$ProjectsRequest = '<key:GetProjects><FilterDTO xmlns="http://www.keyedin.com/" xmlns:i="http://www.w3.org/2001/XMLSchema-instance"><FilterItems /></FilterDTO></key:GetProjects>'
	$Request = $RequestHeader + $ProjectsRequest + $RequestFooter
	
	$Headers = @{"SOAPAction" = "http://www.keyedin.com/IProjectMobileService/GetProjects"}
	$Result = (Invoke-WebRequest -Uri $URI -ContentType "text/xml" -Headers $Headers -Method POST -Body $Request -UserAgent "")
	
	[xml]$ResultXML = $Result.Content
	return $ResultXML.Envelope.Body.GetProjectsResponse.GetProjectsResult.ProjectDTO
}

Function Get-Activities {
    <#
    .SYNOPSIS
        Get your activies.

    .EXAMPLE
        Get-Activities
    #>
	$ActivitiesRequest = '<key:GetActivities><FilterDTO xmlns="http://www.keyedin.com/" xmlns:i="http://www.w3.org/2001/XMLSchema-instance"><FilterItems /></FilterDTO></key:GetActivities>'
	$Request = $RequestHeader + $ActivitiesRequest + $RequestFooter
	
	$Headers = @{"SOAPAction" = "http://www.keyedin.com/IProjectMobileService/GetActivities"}
	$Result = (Invoke-WebRequest -Uri $URI -ContentType "text/xml" -Headers $Headers -Method POST -Body $Request -UserAgent "")
	
	[xml]$ResultXML = $Result.Content
	return $ResultXML.Envelope.Body.GetActivitiesResponse.GetActivitiesResult.ActivityDTO
}

Function Get-Tasks {
    <#
    .SYNOPSIS
        Get your project tasks

    .PARAMETER StartDate
        Start date for task search limit.

    .PARAMETER EndDate
        End date for task search limit.
		
    .PARAMETER ResourceCode
        Your recourse code. Your can find it from Get-SystemConfig response.
		
    .EXAMPLE
        $SystemConfig = Get-SystemConfig
		$Tasks = Get-Tasks -StartDate $(Get-Date -Date "2016-01-01") -EndDate $(Get-Date -Date "2016-12-31") -ResourceCode $SystemConfig.ResourceCode
    #>
	param (
		[Parameter(Mandatory=$true)][DateTime]$StartDate,
		[Parameter(Mandatory=$true)][DateTime]$EndDate,
		[Parameter(Mandatory=$true)][string]$ResourceCode
	)
	[string]$StartDateString = Get-Date -Date $StartDate -Format "yyyy-MM-dd"
	[string]$EndDateString = Get-Date -Date $EndDate -Format "yyyy-MM-dd"
	$TasksRequest = @"
    <key:GetTasks>
      <FilterDTO xmlns=`"http://www.keyedin.com/`" xmlns:i=`"http://www.w3.org/2001/XMLSchema-instance`">
        <FilterItems>
          <FilterItemDTO>
            <Field>FinishDate</Field>
            <Operator>LessThanOrEqualTo</Operator>
            <Value>$EndDateString</Value>
          </FilterItemDTO>
          <FilterItemDTO>
            <Field>StartDate</Field>
            <Operator>GreaterThanOrEqualTo</Operator>
            <Value>$StartDateString</Value>
          </FilterItemDTO>
          <FilterItemDTO>
            <Field>IsActive</Field>
            <Operator>EqualTo</Operator>
            <Value>true</Value>
          </FilterItemDTO>
          <FilterItemDTO>
            <Field>ResourceCode</Field>
            <Operator>EqualTo</Operator>
            <Value>$ResourceCode</Value>
          </FilterItemDTO>
        </FilterItems>
      </FilterDTO>
    </key:GetTasks>
"@
	$Request = $RequestHeader + $TasksRequest + $RequestFooter
	
	$Headers = @{"SOAPAction" = "http://www.keyedin.com/IProjectMobileService/GetTasks"}
	$Result = (Invoke-WebRequest -Uri $URI -ContentType "text/xml" -Headers $Headers -Method POST -Body $Request -UserAgent "")
	
	[xml]$ResultXML = $Result.Content
	return $ResultXML.Envelope.Body.GetTasksResponse.GetTasksResult.ProjectTaskDTO
}

Function Get-Timesheets {
    <#
    .SYNOPSIS
        Get your timesheets

    .PARAMETER StartDate
        Start date for task search limit.

    .PARAMETER EndDate
        End date for task search limit.
		
    .PARAMETER ResourceCode
        Your recourse code. Your can find it from Get-SystemConfig response.
		
    .EXAMPLE
        $SystemConfig = Get-SystemConfig
		$Timesheets = Get-Timesheets -StartDate $(Get-Date -Date "2016-09-01") -EndDate $(Get-Date -Date "2016-09-30") -ResourceCode $SystemConfig.ResourceCode
    #>
	param (
		[Parameter(Mandatory=$true)][string]$StartDate,
		[Parameter(Mandatory=$true)][string]$EndDate,
		[Parameter(Mandatory=$true)][string]$ResourceCode
	)
	$TimesheetsRequest = @"
    <key:GetTimesheets>
      <FilterDTO xmlns=`"http://www.keyedin.com/`" xmlns:i=`"http://www.w3.org/2001/XMLSchema-instance`">
        <FilterItems>
          <FilterItemDTO>
            <Field>Date</Field>
            <Operator>LessThanOrEqualTo</Operator>
            <Value>$EndDate</Value>
          </FilterItemDTO>
          <FilterItemDTO>
            <Field>Date</Field>
            <Operator>GreaterThanOrEqualTo</Operator>
            <Value>$StartDate</Value>
          </FilterItemDTO>
          <FilterItemDTO>
            <Field>ResourceCode</Field>
            <Operator>EqualTo</Operator>
            <Value>$ResourceCode</Value>
          </FilterItemDTO>
        </FilterItems>
      </FilterDTO>
    </key:GetTimesheets>
"@
	$Request = $RequestHeader + $TimesheetsRequest + $RequestFooter
	
	$Headers = @{"SOAPAction" = "http://www.keyedin.com/IProjectMobileService/GetTimesheets"}
	$Result = (Invoke-WebRequest -Uri $URI -ContentType "text/xml" -Headers $Headers -Method POST -Body $Request -UserAgent "")
	
	[xml]$ResultXML = $Result.Content
	return $ResultXML.Envelope.Body.GetTimesheetsResponse.GetTimesheetsResult.TimesheetDTO
}

Function Add-Timesheet {
    <#
    .SYNOPSIS
        Add/extend timesheet.

    .PARAMETER ActivityCode
        Code of activity. You can find it from Get-Activities

    .PARAMETER TaskKey
        Key of task. You can find it from Get-Tasks
		
    .PARAMETER HoursDecimal
        Hours used for this activity.
		
    .PARAMETER IsChargeable
        Is this activity chargeable or not?
		
    .PARAMETER ProjectCode
        Code of project. You can find it from Get-Projects
		
    .PARAMETER TimesheetDate
        Date when activity is done.
		
    .EXAMPLE
        Add-Timesheet -ActivityCode "123" -TaskKey "12345" -HoursDecimal 7.5 -IsChargeable $False -ProjectCode "54321" -TimesheetDate $(Get-Date -Date "2016-01-01")
    #>
	param (
		[Parameter(Mandatory=$true)][string]$ActivityCode,
		[Parameter(Mandatory=$true)][string]$TaskKey,
		[Parameter(Mandatory=$true)][double]$HoursDecimal,
		[Parameter(Mandatory=$true)][bool]$IsChargeable,
		[Parameter(Mandatory=$true)][string]$ProjectCode,
		[Parameter(Mandatory=$true)][DateTime]$TimesheetDate
	)
	[string]$TimesheetDateString = Get-Date -Date $TimesheetDate -Format "yyyy-MM-dd"
	$TimesheetDateString +=	"T00:00:00"
	[string]$HoursDecimalString = $HoursDecimal -replace ",","."
	If ($IsChargeable -eq $True) {
		$IsChargeableString = "true"
	} Else {
		$IsChargeableString = "false"
	}
	
	$InsertOrAmendTimesheetRequest = @"
    <key:InsertOrAmendTimesheet>
      <TimesheetDTO xmlns=`"http://www.keyedin.com/`" xmlns:i=`"http://www.w3.org/2001/XMLSchema-instance`">
        <Activity>
          <Code>$ActivityCode</Code>
        </Activity>
        <ApprovalStatus>1</ApprovalStatus>
        <Assignment i:nil="true" />
        <CustomFields />
        <HoursAndMinutes i:nil="true" />
        <HoursDecimal>$HoursDecimalString</HoursDecimal>
        <IsChargeable>$IsChargeableString</IsChargeable>
        <IsOvertime>false</IsOvertime>
        <IsSubmitted>false</IsSubmitted>
        <Key>-1</Key>
        <LastEditDate>0001-01-01T00:00:00</LastEditDate>
        <Notes />
        <Project>
          <Code>$ProjectCode</Code>
        </Project>
        <Task>
			<Key>$TaskKey</Key>
		</Task>
        <TimesheetDate>$TimesheetDateString</TimesheetDate>
      </TimesheetDTO>
    </key:InsertOrAmendTimesheet>
"@
	$Request = $RequestHeader + $InsertOrAmendTimesheetRequest + $RequestFooter
	
	$Headers = @{"SOAPAction" = "http://www.keyedin.com/IProjectMobileService/InsertOrAmendTimesheet"}
	$Result = (Invoke-WebRequest -Uri $URI -ContentType "text/xml" -Headers $Headers -Method POST -Body $Request -UserAgent "")
	
	[xml]$ResultXML = $Result.Content
	return $ResultXML.Envelope.Body.InsertOrAmendTimesheetResponse.InsertOrAmendTimesheetResult.Succeeded
}