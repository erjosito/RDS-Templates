param(
  [Parameter(mandatory = $false)]
  [object]$WebHookData
)
# If runbook was called from Webhook, WebhookData will not be null.
if ($WebHookData) {

  # Collect properties of WebhookData
  $WebhookName = $WebHookData.WebhookName
  $WebhookHeaders = $WebHookData.RequestHeader
  $WebhookBody = $WebHookData.RequestBody

  # Collect individual headers. Input converted from JSON.
  $From = $WebhookHeaders.From
  $Input = (ConvertFrom-Json -InputObject $WebhookBody)
}
else
{
  Write-Error -Message 'Runbook was not started from Webhook' -ErrorAction stop
}
$AADTenantId = $Input.AADTenantId
$SubscriptionID = $Input.SubscriptionID
$TenantGroupName = $Input.TenantGroupName
$ResourceGroupName = $Input.ResourceGroupName
$TenantName = $Input.TenantName
$HostpoolName = $Input.hostpoolname
$BeginPeakTime = $Input.BeginPeakTime
$EndPeakTime = $Input.EndPeakTime
$TimeDifference = $Input.TimeDifference
$SessionThresholdPerCPU = $Input.SessionThresholdPerCPU
$MinimumNumberOfRDSH = $Input.OffpeakMinimumNumberOfRDSH
$PeakMinimumNumberOfRDSH = $Input.PeakMinimumNumberOfRDSH
$LimitSecondsToForceLogOffUser = $Input.LimitSecondsToForceLogOffUser
$LogOffMessageTitle = $Input.LogOffMessageTitle
$LogOffMessageBody = $Input.LogOffMessageBody
$MaintenanceTagName = $Input.MaintenanceTagName
$LogAnalyticsWorkspaceId = $Input.LogAnalyticsWorkspaceId
$LogAnalyticsPrimaryKey = $Input.LogAnalyticsPrimaryKey
$RDBrokerURL = $Input.RDBrokerURL
$AutomationAccountName = $Input.AutomationAccountName
$ConnectionAssetName = $Input.ConnectionAssetName
$WeekDaysPeak = $Input.WeekDaysPeak  # Expected as a space-separated string, such as "1 2 3 4 5" (0=Sunday, 6=Saturday)
$BankHolidays = $Input.BankHolidays  # Expected as a space-separated string with MM/DD, such as "5/1 12/25"
$ScaleUpDuringOffpeak = $Input.ScaleUpDuringOffPeak  # Expected "yes" or "no"

# Set defaults if some of the parameters havent been passed on
if (!$WeekDaysPeak) {
    $WeekDaysPeak = "1 2 3 4 5"  # Defaults to Mon-Fri
}
if (!$MinimumNumberOfRDSH) {
    $MinimumNumberOfRDSH = 0  # Defaults to 0
}
if (!$PeakMinimumNumberOfRDSH) {
    $PeakMinimumNumberOfRDSH = $MinimumNumberOfRDSH  # Defaults to the Offpeak minimum
}
$ScaleUpDuringOffpeak = ($ScaleUpDuringOffpeak -eq "yes")  # Defaults to $false
if (!$BankHolidays) {
    $BankHolidaysArray = @()
} else {
    $BankHolidaysArray = $BankHolidays -split " "
}

# Transform the string for week days into an array
$WeekDaysPeakArray = $WeekDaysPeak -split " "

# Control variables to prevent endless loops
# Used mainly in loops waiting to start/stop VMs: 5 minutes waiting seems to be sufficient time
$StartStopWaitTime = 30  # seconds
$StartStopMaxCycles = 10 # times to go through the waittime

Set-ExecutionPolicy -ExecutionPolicy Undefined -Scope Process -Force -Confirm:$false
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope LocalMachine -Force -Confirm:$false
# Setting ErrorActionPreference to stop script execution when error occurs
$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Function to convert from UTC to Local time
function Convert-UTCtoLocalTime
{
    param(
      $TimeDifferenceInHours
    )

    $UniversalTime = (Get-Date).ToUniversalTime()
    $TimeDifferenceMinutes = 0
    if ($TimeDifferenceInHours -match ":") {
      $TimeDifferenceHours = $TimeDifferenceInHours.Split(":")[0]
      $TimeDifferenceMinutes = $TimeDifferenceInHours.Split(":")[1]
    }
    else {
      $TimeDifferenceHours = $TimeDifferenceInHours
    }
    #Azure is using UTC time, justify it to the local time
    $ConvertedTime = $UniversalTime.AddHours($TimeDifferenceHours).AddMinutes($TimeDifferenceMinutes)
    return $ConvertedTime
}

# Function for to add logs to log analytics workspace
function Add-LogEntry
    {
    param(
        [Object]$LogMessageObj,
        [string]$LogAnalyticsWorkspaceId,
        [string]$LogAnalyticsPrimaryKey,
        [string]$LogType,
        $TimeDifferenceInHours
    )

    if ($LogAnalyticsWorkspaceId) {

        foreach ($Key in $LogMessage.Keys) {
            switch ($Key.substring($Key.Length - 2)) {
                '_s' { $sep = '"'; $trim = $Key.Length - 2 }
                '_t' { $sep = '"'; $trim = $Key.Length - 2 }
                '_b' { $sep = ''; $trim = $Key.Length - 2 }
                '_d' { $sep = ''; $trim = $Key.Length - 2 }
                '_g' { $sep = '"'; $trim = $Key.Length - 2 }
                default { $sep = '"'; $trim = $Key.Length }
            }
            $LogData = $LogData + '"' + $Key.substring(0,$trim) + '":' + $sep + $LogMessageObj.Item($Key) + $sep + ','
        }
        $TimeStamp = Convert-UTCtoLocalTime -TimeDifferenceInHours $TimeDifferenceInHours
        $LogData = $LogData + '"TimeStamp":"' + $timestamp + '"'

        #Write-Verbose "LogData: $($LogData)"
        $json = "{$($LogData)}"

        $PostResult = Send-OMSAPIIngestionFile -customerId $LogAnalyticsWorkspaceId -sharedKey $LogAnalyticsPrimaryKey -Body "$json" -logType $LogType -TimeStampField "TimeStamp"
        #Write-Verbose "PostResult: $($PostResult)"
        if ($PostResult -ne "Accepted") {
        Write-Error "Error posting to OMS - $PostResult"
        }
    }
}

function Send-Message {
    param(
        [string]$Message,
        [string]$HostpoolName,
        [string]$LogAnalyticsWorkspaceId,
        [string]$LogAnalyticsPrimaryKey,
        $TimeDifferenceInHours
    )
    Write-Output $Message
    if ($LogAnalyticsWorkspaceId -and $LogAnalyticsPrimaryKey) {
        $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Failed to authenticate Azure: $($_.exception.message)" }
        Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -logType "WVDTenantScale_CL" -TimeDifferenceInHours $TimeDifference        
    }

}

# Display some the control variables
$Message = "Running with these control variables: 'MinimumNumberOfRDSH' = '$MinimumNumberOfRDSH', `
           'PeakMinimumNumberOfRDSH': '$PeakMinimumNumberOfRDSH', 'WeekDaysPeak': '$WeekDaysPeak', 'BankHolidays': '$BankHolidays', `
           'ScaleUpDuringOffpeak': '$ScaleUpDuringOffpeak' 'SessionThresholdPerCPU': '$SessionThresholdPerCPU'"
Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey

# Collect the credentials from Azure Automation Account Assets
$Connection = Get-AutomationConnection -Name $ConnectionAssetName

# Authenticating to Azure
Clear-AzContext -Force
$AZAuthentication = Connect-AzAccount -ApplicationId $Connection.ApplicationId -TenantId $AADTenantId -CertificateThumbprint $Connection.CertificateThumbprint -ServicePrincipal
if (!$AZAuthentication) {
    Send-Message -Message "Failed to authenticate Azure: $($_.exception.message)" -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
    exit
} else {
    $AzObj = $AZAuthentication | Out-String
    $Message = "Authenticating as service principal for Azure. Result: `n$AzObj"
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
}
# Set the Azure context with Subscription
$AzContext = Set-AzContext -SubscriptionId $SubscriptionID
if (!$AzContext) {
    Write-Error "Please provide a valid subscription"
    exit
} else {
    $AzSubObj = $AzContext | Out-String
    $Message = "Sets the Azure subscription. Result: `n$AzSubObj"
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
}

# Authenticating to WVD
try {
    $WVDAuthentication = Add-RdsAccount -DeploymentUrl $RDBrokerURL -ApplicationId $Connection.ApplicationId -CertificateThumbprint $Connection.CertificateThumbprint -AADTenantId $AadTenantId
}
catch {
    $Message = "Failed to authenticate WVD: $($_.exception.message)"
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
}
$WVDObj = $WVDAuthentication | Out-String
$Message = "Authenticating as service principal for WVD. Result: `n$WVDObj"
Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey

# Function to check and update the loadbalancer type is BreadthFirst
function UpdateLoadBalancerTypeInPeakandOffPeakwithBredthFirst {
    param(
      [string]$HostpoolLoadbalancerType,
      [string]$TenantName,
      [string]$HostpoolName,
      [int]$MaxSessionLimitValue,
      [string]$LogAnalyticsWorkspaceId,
      [string]$LogAnalyticsPrimaryKey,
      $TimeDifference
    )
    if ($HostpoolLoadbalancerType -ne "BreadthFirst") {
        Write-Output "Changing hostpool load balancer type:'BreadthFirst' Current Date Time is: $CurrentDateTime"
        $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Changing hostpool load balancer type:'BreadthFirst' Current Date Time is: $CurrentDateTime" }
        Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -logType "WVDTenantScale_CL" -TimeDifferenceInHours $TimeDifference
        $EditLoadBalancerType = Set-RdsHostPool -TenantName $TenantName -Name $HostpoolName -BreadthFirstLoadBalancer -MaxSessionLimit $MaxSessionLimitValue
        if ($EditLoadBalancerType.LoadBalancerType -eq 'BreadthFirst') {
            Write-Output "Hostpool load balancer type in peak hours is 'BreadthFirst'"
            $LogMessage = @{ hostpoolName_s = $HostpoolName; logmessage_s = "Hostpool load balancer type in peak hours is 'BreadthFirst'" }
            Add-LogEntry -LogMessageObj $LogMessage -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -logType "WVDTenantScale_CL" -TimeDifferenceInHours $TimeDifference
        }
    }
}

# Function to Check if the session host is allowing new connections, and enable it of not
function Set-AllowNewSession
{
    param(
        [string]$TenantName,
        [string]$HostpoolName,
        [string]$SessionHostName
    )
    # Check if the session host is allowing new connections
    $StateOftheSessionHost = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHostName
    if (!($StateOftheSessionHost.AllowNewSession)) {
        Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHostName -AllowNewSession $true
    }
}

# Function to start a Session Host 
function Start-SessionHost
{
    param(
        [string]$VMName,
        [string]$LogAnalyticsWorkspaceId,
        [string]$LogAnalyticsPrimaryKey,
        $TimeDifference
    )
    try {
        Get-AzVM | Where-Object { $_.Name -eq $VMName } | Start-AzVM -AsJob | Out-Null
    }
    catch {
        $Message = "Failed to start Azure VM: $($VMName) with error: $($_.exception.message)"
        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
        exit
    }
}

# Function to stop a Session Host
function Stop-SessionHost
{
    param(
        [string]$VMName,
        [string]$LogAnalyticsWorkspaceId,
        [string]$LogAnalyticsPrimaryKey,
        $TimeDifference
    )
    try {
        Get-AzVM | Where-Object { $_.Name -eq $VMName } | Stop-AzVM -Force -AsJob | Out-Null
    }
    catch {
        $Message = "Failed to stop Azure VM: $($VMName) with error: $($_.exception.message)"
        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
        exit
    }
}

# Function to check if the Session host is available
function Wait-SessionHostAvailable
{
    param(
        [string]$TenantName,
        [string]$HostpoolName,
        [string]$SessionHostName
    )
    $WaitCycle=0
    $IsHostAvailable = $false
    while (!$IsHostAvailable -and $WaitCycle -lt $StartStopMaxCycles) {
        $WaitCycle += 1
        $SessionHostStatus = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHostName
        if ($SessionHostStatus.Status -eq "Available") {
            $IsHostAvailable = $true
        } else {
            Start-Sleep -Seconds $StartStopWaitTime
        }
    }
    return $IsHostAvailable
}

# Converting date time from UTC to Local
$CurrentDateTime = Convert-UTCtoLocalTime -TimeDifferenceInHours $TimeDifference
$CurrentDayOfWeek = $CurrentDateTime.DayOfWeek.value__.ToString()

# Set context to the appropriate tenant group
$CurrentTenantGroupName = (Get-RdsContext).TenantGroupName
if ($TenantGroupName -ne $CurrentTenantGroupName) {
    $Message = "Running switching to the $TenantGroupName context"
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
    Set-RdsContext -TenantGroupName $TenantGroupName
}

$BeginPeakDateTime = [datetime]::Parse($CurrentDateTime.ToShortDateString() + ' ' + $BeginPeakTime)
$EndPeakDateTime = [datetime]::Parse($CurrentDateTime.ToShortDateString() + ' ' + $EndPeakTime)

# check the calculated end time is later than begin time in case of time zone
if ($EndPeakDateTime -lt $BeginPeakDateTime) {
    $EndPeakDateTime = $EndPeakDateTime.AddDays(1)
}

# Checking given host pool name exists in Tenant
$HostpoolInfo = Get-RdsHostPool -TenantName $TenantName -Name $HostpoolName
if (!$HostpoolInfo) {
    $Message = "Hostpoolname '$HostpoolName' does not exist in the tenant of '$TenantName'. Ensure that you have entered the correct values."
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
    exit
}

# Setting up appropriate load balacing type based on PeakLoadBalancingType in Peak hours
$HostpoolLoadbalancerType = $HostpoolInfo.LoadBalancerType
[int]$MaxSessionLimitValue = $HostpoolInfo.MaxSessionLimit
if ($CurrentDateTime -ge $BeginPeakDateTime -and $CurrentDateTime -le $EndPeakDateTime) {
    UpdateLoadBalancerTypeInPeakandOffPeakwithBredthFirst -TenantName $TenantName -HostPoolName $HostpoolName -MaxSessionLimitValue $MaxSessionLimitValue -HostpoolLoadbalancerType $HostpoolLoadbalancerType -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -TimeDifference $TimeDifference
} else {
    UpdateLoadBalancerTypeInPeakandOffPeakwithBredthFirst -TenantName $TenantName -HostPoolName $HostpoolName -MaxSessionLimitValue $MaxSessionLimitValue -HostpoolLoadbalancerType $HostpoolLoadbalancerType -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -TimeDifference $TimeDifference
}
$Message = "Starting WVD tenant hosts scale optimization: Current Date Time is: $CurrentDateTime, current week day is $CurrentDayOfWeek"
Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
# Check the after changing hostpool loadbalancer type
$HostpoolInfo = Get-RdsHostPool -TenantName $TenantName -Name $HostPoolName

# Check if the hostpool have session hosts
$ListOfSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -ErrorAction Stop | Sort-Object SessionHostName
if (!$ListOfSessionHosts) {
    $Message = "Session hosts does not exist in the Hostpool of '$HostpoolName'. Ensure that the hostpool has at least one session host"
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
    exit
}

# Check whether today is a bank holiday
$TodayIsBankHoliday = $false
foreach ($BankHoliday in $BankHolidaysArray) {
    $BankHolidayDetails = $BankHoliday.ToString() -split "/"
    if ($BankHolidayDetails.Length -eq 2)
    {
        $BankHolidayDay = $BankHolidayDetails[1].ToString()
        $BankHolidayMonth = $BankHolidayDetails[0].ToString()
        if (($CurrentDateTime.Day.ToString() -eq $BankHolidayDay) -and ($CurrentDateTime.Month.ToString() -eq $BankHolidayMonth)) {
                $Message = "Today appears to be a bank holiday"
                Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                $TodayIsBankHoliday = $true
        }    
    } else {
        $Message = "The date $BankHoliday seems to be malformed. It should be in the format MM/DD"
        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
    }
}

# Check if it is during the peak or off-peak time
if (($CurrentDateTime -ge $BeginPeakDateTime) -and ($CurrentDateTime -le $EndPeakDateTime) -and ($WeekDaysPeakArray.Contains($CurrentDayOfWeek)) -and (-not $TodayIsBankHoliday)) {
    ##############################################
    #                Peak hours                 #
    ##############################################
    $Message = "It is peak hours now."
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey

    # Peak hours check and remove the MinimumNoOfRDSH value dynamically stored in automation variable                            
    $AutomationAccount = Get-AzAutomationAccount -ErrorAction Stop | Where-Object { $_.AutomationAccountName -eq $AutomationAccountName }
    $OffPeakUsageMinimumNoOfRDSH = Get-AzAutomationVariable -Name "$HostpoolName-OffPeakUsage-MinimumNoOfRDSH" -ResourceGroupName $AutomationAccount.ResourceGroupName -AutomationAccountName $AutomationAccount.AutomationAccountName -ErrorAction SilentlyContinue
    if ($OffPeakUsageMinimumNoOfRDSH) {
        Remove-AzAutomationVariable -Name "$HostpoolName-OffPeakUsage-MinimumNoOfRDSH" -ResourceGroupName $AutomationAccount.ResourceGroupName -AutomationAccountName $AutomationAccount.AutomationAccountName
    }

    # Initialize variables
    # Check the number of running session hosts
    [int]$NumberOfRunningHost = 0
    # Total of running cores
    [int]$TotalRunningCores = 0
    # Total capacity of sessions of running VMs
    $AvailableSessionCapacity = 0
    #Initialize variable for to skip the session host which is in maintenance.
    $SkipSessionhosts = 0
    $SkipSessionhosts = @()

    # Calculate current number of sessions
    $HostPoolUserSessions = Get-RdsUserSession -TenantName $TenantName -HostPoolName $HostpoolName
    foreach ($SessionHost in $ListOfSessionHosts) {
        $SessionHostName = $SessionHost.SessionHostName | Out-String
        $VMName = $SessionHostName.Split(".")[0]

        # Check if VM is in maintenance (has been tagged with the $MaintenanceTagName, the value of the tag is not relevant)
        $RoleInstance = Get-AzVM -Status | Where-Object { $_.Name.Contains($VMName) }
        if ($RoleInstance.Tags.Keys -contains $MaintenanceTagName) {
            $Message = "Session host is in maintenance: $VMName, so script will skip this VM"
            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
            $SkipSessionhosts += $SessionHost
            continue
        }
        $AllSessionHosts = Compare-Object $ListOfSessionHosts $SkipSessionhosts | Where-Object { $_.SideIndicator -eq '<=' } | ForEach-Object { $_.InputObject }

        $Message = "Checking session host $($SessionHost.SessionHostName): $($SessionHost.Sessions) and status: $($SessionHost.Status)"
        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey

        if ($SessionHostName.ToLower().Contains($RoleInstance.Name.ToLower())) {
            # Check if the Azure vm is running       
            if ($RoleInstance.PowerState -eq "VM running") {
                [int]$NumberOfRunningHost = [int]$NumberOfRunningHost + 1
                # Calculate available capacity of sessions            
                $RoleSize = Get-AzVMSize -Location $RoleInstance.Location | Where-Object { $_.Name -eq $RoleInstance.HardwareProfile.VmSize }
                $AvailableSessionCapacity = $AvailableSessionCapacity + $RoleSize.NumberOfCores * $SessionThresholdPerCPU
                [int]$TotalRunningCores = [int]$TotalRunningCores + $RoleSize.NumberOfCores
            }
        }
    }

    $Message = "Current number of running hosts: $NumberOfRunningHost. Total number of running cores: $($TotalRunningCores.ToString()). Available session capacity: $($AvailableSessionCapacity.ToString())"
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey

    # PeakMinimumNuberOfRDSH was set and we need more hosts
    if ($NumberOfRunningHost -lt $PeakMinimumNumberOfRDSH) {

        $Message = "Current number of running session hosts ($NumberOfRunningHost) is less than minimum requirements for peak hours ($PeakMinimumNumberOfRDSH), starting session host."
        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
        # Start VM to meet the minimum requirement            
        foreach ($SessionHost in $AllSessionHosts.SessionHostName) {
            # Check whether the number of running VMs now meets the minimum or not
            if ($NumberOfRunningHost -lt $PeakMinimumNumberOfRDSH) {
                $VMName = $SessionHost.Split(".")[0]
                $RoleInstance = Get-AzVM -Status | Where-Object { $_.Name.Contains($VMName) }
                if ($SessionHost.ToLower().Contains($RoleInstance.Name.ToLower())) {
                    # Check if the Azure VM is running and if the session host is healthy
                    $SessionHostInfo = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost
                    if ($RoleInstance.PowerState -ne "VM running" -and $SessionHostInfo.UpdateState -eq "Succeeded") {
                        # Enable new connections for this host (the scale down process might have set the session host in drain mode)
                        Set-AllowNewSession -TenantName $TenantName -HostPoolName $HostpoolName -SessionHostName $SessionHost
                        # Start the Az VM
                        $Message = "Starting Azure VM: $VMName and waiting for it to complete."
                        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                        Start-SessionHost -VMName $VMName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -TimeDifference $TimeDifference
                        # Wait for the VM to Start
                        $IsVMStarted = $false
                        while (!$IsVMStarted) {
                            $RoleInstance = Get-AzVM -Status | Where-Object { $_.Name -eq $VMName }
                            if ($RoleInstance.PowerState -eq "VM running") {
                                $IsVMStarted = $true
                                $Message = "Azure VM has been started: $($RoleInstance.Name)."
                                Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                            }
                        }
                        # Wait for the VM to become available to WVD
                        $SessionHostIsAvailable = Wait-SessionHostAvailable -TenantName $TenantName -HostPoolName $HostpoolName -SessionHost $SessionHost
                        if ($SessionHostIsAvailable) {
                            $Message = "'$SessionHost' session host status is 'Available'"
                            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                        }
                        else {
                            $Message = "'$SessionHost' session host does not configured properly with deployagent or does not started properly"
                            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                        }
                        # Calculate available capacity of sessions
                        $RoleSize = Get-AzVMSize -Location $RoleInstance.Location | Where-Object { $_.Name -eq $RoleInstance.HardwareProfile.VmSize }
                        $AvailableSessionCapacity = $AvailableSessionCapacity + $RoleSize.NumberOfCores * $SessionThresholdPerCPU
                        [int]$NumberOfRunningHost = [int]$NumberOfRunningHost + 1
                        [int]$TotalRunningCores = [int]$TotalRunningCores + $RoleSize.NumberOfCores
                        if ($NumberOfRunningHost -ge $MinimumNumberOfRDSH) {
                            break;
                        }
                    }
                }
            }
        }
    }
    # We are over the minimum number of hosts (typically because it is not the first run during this peak hours cycle)
    # We still might need to scale up
    else
    {
        $Message = "Current number of running session hosts ($NumberOfRunningHost) is at least the minimum requirements for peak hours ($PeakMinimumNumberOfRDSH)"
        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
        # Check if the available capacity meets the number of sessions or not
        $Message = "Current total number of user sessions: $($HostPoolUserSessions.Count), available session capacity is: $AvailableSessionCapacity"
        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
        # If we need to scale up
        if ($HostPoolUserSessions.Count -ge $AvailableSessionCapacity) {
            $Message = "Current available session capacity is less than demanded user sessions, starting session host"
            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey

            # Running out of capacity, we need to start more VMs if there are any 
            foreach ($SessionHost in $AllSessionHosts.SessionHostName) {
                if ($HostPoolUserSessions.Count -ge $AvailableSessionCapacity) {
                    $VMName = $SessionHost.Split(".")[0]
                    $RoleInstance = Get-AzVM -Status | Where-Object { $_.Name.Contains($VMName) }
                    if ($SessionHost.ToLower().Contains($RoleInstance.Name.ToLower())) {
                        # Check if the Azure VM is running and if the session host is healthy
                        $SessionHostInfo = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHost
                        if ($RoleInstance.PowerState -ne "VM running" -and $SessionHostInfo.UpdateState -eq "Succeeded") {
                            # Validating session host is allowing new connections
                            Set-AllowNewSession -TenantName $TenantName -HostPoolName $HostpoolName -SessionHostName $SessionHost
                            # Start the Az VM
                            $Message = "Starting Azure VM: $VMName and waiting for it to complete."
                            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                            Start-SessionHost -VMName $VMName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -TimeDifference $TimeDifference
                            # Wait for the VM to Start
                            $IsVMStarted = $false
                            while (!$IsVMStarted) {
                                $RoleInstance = Get-AzVM -Status | Where-Object { $_.Name -eq $VMName }
                                if ($RoleInstance.PowerState -eq "VM running") {
                                    $IsVMStarted = $true
                                    $Message = "Azure VM has been Started: $($RoleInstance.Name)."
                                    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                                }
                            }
                            $SessionHostIsAvailable = Wait-SessionHostAvailable -TenantName $TenantName -HostPoolName $HostpoolName -SessionHost $SessionHost
                            if ($SessionHostIsAvailable) {
                                $Message = "'$SessionHost' session host status is 'Available'"
                                Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                            }
                            else {
                                $Message = "'$SessionHost' session host does not configured properly with deployagent or does not started properly"
                                Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                            }
                            # Calculate new available capacity of sessions with the new RDSH
                            $RoleSize = Get-AzVMSize -Location $RoleInstance.Location | Where-Object { $_.Name -eq $RoleInstance.HardwareProfile.VmSize }
                            $AvailableSessionCapacity = $AvailableSessionCapacity + $RoleSize.NumberOfCores * $SessionThresholdPerCPU
                            [int]$NumberOfRunningHost = [int]$NumberOfRunningHost + 1
                            [int]$TotalRunningCores = [int]$TotalRunningCores + $RoleSize.NumberOfCores
                            $Message = "New available session capacity is: $AvailableSessionCapacity"
                            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                            if ($AvailableSessionCapacity -gt $HostPoolUserSessions.Count) {
                                break
                            }
                        }
                    }
                }
            }
        } else {
            $Message = "Current available session capacity is enough to satisfy the current session demand, no need to scale up"
            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
        }
    }
}
# Note there is no scale down during peak hours
else
{
    ##############################################
    #               Off-peak hours               #
    ##############################################
    $Message = "It is Off-peak hours"
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
    $Message = "Verifying whether we need to scale down the number of WVD session hosts."
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
    $Message = "Processing hostpool $($HostpoolName)"
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
       
    # Check the number of running session hosts
    [int]$NumberOfRunningHost = 0
    # Total number of running cores
    [int]$TotalRunningCores = 0
    #Initialize variable for to skip the session host which is in maintenance.
    $SkipSessionhosts = @()

    # Check if the minimum number of rdsh VM's are running and available
    $CheckMinimumNumberOfRDShIsRunning = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Where-Object { $_.Status -eq "Available" }
    if (!$CheckMinimumNumberOfRDShIsRunning) {
        $NumberOfRunningHost = 0
    } else {
        $NumberOfRunningHost = $CheckMinimumNumberOfRDShIsRunning.Length
    }

  
    # Total list of hosts (running and not running)
    $ListOfSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName

    # See if we need to add session hosts
    if ($NumberOfRunningHost -lt $MinimumNumberOfRDSH) {
        $Message = "There are too few session hosts in this pool ($NumberOfRunningHost), starting hosts to reach the minimum ($MinimumNumberOfRDSH hosts)"
        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
        foreach ($SessionHostName in $ListOfSessionHosts.SessionHostName) {
            # Loop through the VMs until we reach the minimum number of hosts
            if ($NumberOfRunningHost -lt $MinimumNumberOfRDSH) {
                $VMName = $SessionHostName.Split(".")[0]
                $RoleInstance = Get-AzVM -Status | Where-Object { $_.Name.Contains($VMName) }
                # Check the session host is in maintenance
                if ($RoleInstance.Tags.Keys -contains $MaintenanceTagName) {
                    continue
                }
                # Check if the session host is allowing new connections, and enable it
                Set-AllowNewSession -TenantName $TenantName -HostPoolName $HostpoolName -SessionHostName $SessionHostName
                Start-SessionHost -VMName $VMName
                # Wait for the VM to Start
                $IsVMStarted = $false
                while (!$IsVMStarted) {
                    $RoleInstance = Get-AzVM -Status | Where-Object { $_.Name -eq $VMName }
                    if ($RoleInstance.PowerState -eq "VM running") {
                        $IsVMStarted = $true
                    }
                }
                # Wait for the VM to start
                $SessionHostIsAvailable = Wait-SessionHostAvailable -TenantName $TenantName -HostPoolName $HostpoolName -SessionHost $SessionHost
                if ($SessionHostIsAvailable) {
                    $NumberOfRunningHost += 1
                    $Message = "'$SessionHost' session host status is now 'Available'"
                    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                }
                else {
                    $Message = "Session host '$SessionHost' has not started properly. Maybe it is not configured properly with the deploy agent?"
                    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                }
            }
        }
    } else {
        if (!$NumberOfRunningHost) {
            $NumberOfRunningHost = 0
        }
        $Message = "This pool has $NumberOfRunningHost hosts running, (the minimum is $MinimumNumberOfRDSH): no need to scale up"
        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
    }

    # Calculate existing number of sessions, to see whether VMs have to be scaled down
    $ListOfSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Sort-Object Sessions
    foreach ($SessionHost in $ListOfSessionHosts) {
        $SessionHostName = $SessionHost.SessionHostName
        $VMName = $SessionHostName.Split(".")[0]
        $RoleInstance = Get-AzVM -Status | Where-Object { $_.Name.Contains($VMName) }
        # Check the session host is in maintenance
        if ($RoleInstance.Tags.Keys -contains $MaintenanceTagName) {
            $Message = "Session host is in maintenance: $VMName, so the script will skip this VM"
            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
            $SkipSessionhosts += $SessionHost
            continue
        } else {
            $Message = "Checking session host $SessionHost on VM $VMName..."
            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey    
        }

        # Remove the hosts in maintenance mode from the list of hosts
        $AllSessionHosts = $ListOfSessionHosts | Where-Object { $SkipSessionhosts -notcontains $_ }

        # Go through each host to calculate current capacity and total number of cores
        if ($SessionHostName.ToLower().Contains($RoleInstance.Name.ToLower())) {
            # Check if the Azure VM is running or not
            if ($RoleInstance.PowerState -eq "VM running") {
                $Message = "Checking session host: $($SessionHost.SessionHostName.ToString()): $($SessionHost.Sessions) sessions and status $($SessionHost.Status)"
                Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                [int]$NumberOfRunningHost = [int]$NumberOfRunningHost + 1
                # Calculate available capacity of sessions  
                $RoleSize = Get-AzVMSize -Location $RoleInstance.Location | Where-Object { $_.Name -eq $RoleInstance.HardwareProfile.VmSize }
                [int]$TotalRunningCores = [int]$TotalRunningCores + $RoleSize.NumberOfCores
            } else {
                $Message = "Session host $SessionHost is powered off"
                Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey    
                # Disconnect stale sessions to this host here. There shouldnt be any sessions to a powered off VM
                $ShutdownVMSessions = $(Get-RdsUserSession -TenantName $TenantName -HostpoolName $HostpoolName | Where-Object { $_.SessionHostName -eq $SessionHost.SessionHostName})
                foreach ($session in $ShutdownVMSessions) {
                    $Message = "Session $($session.SessionId) from user $($session.UserPrincipalName) and state $($session.SessionState) is connected to session host $($session.SessionHostName), although that VM is powered off!"
                    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                    # Uncomment next line to invoke the session logoff
                    # Invoke-RdsUserSessionLogoff -TenantName $TenantName -HostPoolName $HostPoolName -SessionHostName $session.SessionHostName -SessionId $session.SessionId -Force -NoUserPrompt
                }
            }
        }
    }
    # Defined minimum no of RDSH is the value received from Webhook Data
    [int]$DefinedMinimumNumberOfRDSH = [int]$MinimumNumberOfRDSH

    # Retrieve the minimum number of RDS Hosts (OffPeakUsageMinimumNoOfRDSH) from an Azure Automation Variable
    # OffPeakUsageMinimumNoOfRDSH will be >= than DefinedMinimumNumberOfRDSH
    $Message = "Verifying automation variable '$HostpoolName-OffPeakUsage-MinimumNoOfRDSH'..."
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey    
    $AutomationAccount = Get-AzAutomationAccount -ErrorAction Stop | Where-Object { $_.AutomationAccountName -eq $AutomationAccountName }
    $OffPeakUsageMinimumNoOfRDSH = Get-AzAutomationVariable -Name "$HostpoolName-OffPeakUsage-MinimumNoOfRDSH" -ResourceGroupName $AutomationAccount.ResourceGroupName -AutomationAccountName $AutomationAccount.AutomationAccountName -ErrorAction SilentlyContinue
    if ($OffPeakUsageMinimumNoOfRDSH) {
        $Message = "Azure Automation variable '$HostpoolName-OffPeakUsage-MinimumNoOfRDSH' retrieved with value $OffPeakUsageMinimumNoOfRDSH"
        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
        [int]$MinimumNumberOfRDSH = $OffPeakUsageMinimumNoOfRDSH.Value
        if($MinimumNumberOfRDSH -lt $DefinedMinimumNumberOfRDSH) {
            $Message = "Don't enter the value of the Azure Automation variable '$HostpoolName-OffPeakUsage-MinimumNoOfRDSH' manually, which is dynamically stored value by script. You have entered manually, so script will stop now."
            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
            Exit
        }
    } else {
        $Message = "Azure Automation variable '$HostpoolName-OffPeakUsage-MinimumNoOfRDSH' could not be retrieved"
        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
    }

    # We can now start to shut down VMs until reaching $MinimumNumberOfRDSH, which will be:
    # 1) The value received from the webhook, if this is the first run during this offpeak cycle
    # 2) A higher number if the farm was scaled up during the previous offpeak cycle
    if ($NumberOfRunningHost -gt $MinimumNumberOfRDSH) {
        $Message = "There are too many session hosts ($NumberOfRunningHost), we need to shut down some to come to the minimum ($MinimumNumberOfRDSH)"
        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
        foreach ($SessionHost in $AllSessionHosts) {
            # Only handle unresponsive hosts
            if ($SessionHost.Status -ne "NoHeartbeat" -or $SessionHost.Status -ne "Unavailable") {
                # See if we need to shut down this one to reach the target of $MinimumNumberOfRDSH
                if ($NumberOfRunningHost -gt $MinimumNumberOfRDSH) {
                    $SessionHostName = $SessionHost.SessionHostName
                    $VMName = $SessionHostName.Split(".")[0]
                    $Message = "Trying to stop Azure VM $VMName and waiting for it to complete."
                    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                    if ($SessionHost.Sessions -eq 0) {
                        $Message = "VM $VMName has no active sessions, shutting down now"
                        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                            Stop-SessionHost -VMName $VMName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -TimeDifference $TimeDifference
                    } else {
                        # Set the session host to drain mode (do not accept new connections)
                        try {
                            Set-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHostName -AllowNewSession $false -ErrorAction Stop
                        }
                        catch {
                            $Message = "Unable to set it to allow connections on session host: $SessionHostName with error: $($_.exception.message)"
                            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                            exit
                        }
                        # Notify all users in the RDSH to log off
                        try {
                            $HostPoolUserSessions = Get-RdsUserSession -TenantName $TenantName -HostPoolName $HostpoolName | Where-Object { $_.SessionHostName -eq $SessionHostName }
                        }
                        catch {
                            $Message = "Failed to retrieve user sessions in hostpool: $($HostpoolName) with error: $($_.exception.message)"
                            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                            exit
                        }
                        $HostUserSessionCount = ($HostPoolUserSessions | Where-Object -FilterScript { $_.SessionHostName -eq $SessionHostName }).Count
                        $Message = "Counting the current sessions on the host $SessionHostName :$HostUserSessionCount"
                        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                        $ExistingSession = 0
                        foreach ($session in $HostPoolUserSessions) {
                            if ($session.SessionHostName -eq $SessionHostName -and $session.SessionState -eq "Active") {
                                if ($LimitSecondsToForceLogOffUser -ne 0) {
                                    # Send notification
                                    try {
                                        Send-RdsUserSessionMessage -TenantName $TenantName -HostPoolName $HostpoolName -SessionHostName $SessionHostName -SessionId $session.SessionId -MessageTitle $LogOffMessageTitle -MessageBody "$($LogOffMessageBody) You will be logged off in $($LimitSecondsToForceLogOffUser) seconds." -NoUserPrompt -ErrorAction Stop
                                    }
                                    catch {
                                        $Message = "Failed to send message to user with error: $($_.exception.message)"
                                        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                                        exit
                                    }
                                    $Message = "Script was sent a log off message to user: $($Session.UserPrincipalName | Out-String)"
                                    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                                }
                            }
                            $ExistingSession = $ExistingSession + 1
                        }
                        # Wait for n seconds before forceful logoff
                        Start-Sleep -Seconds $LimitSecondsToForceLogOffUser

                        # Log off users forcefully
                        if ($LimitSecondsToForceLogOffUser -ne 0) {
                            # Force users to log off
                            $Message = "Forcing users to log off."
                            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                            foreach ($Session in $HostPoolUserSessions) {
                                if ($Session.SessionHostName -eq $SessionHostName) {
                                    #Log off user
                                    try {
                                        Invoke-RdsUserSessionLogoff -TenantName $TenantName -HostPoolName $HostpoolName -SessionHostName $Session.SessionHostName -SessionId $Session.SessionId -NoUserPrompt -Force -ErrorAction Stop
                                        $ExistingSession = $ExistingSession - 1
                                    }
                                    catch {
                                        $Message = "Failed to log off user with error: $($_.exception.message)"
                                        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                                        exit
                                    }
                                    $Message = "Forcibly logged off the user: $($Session.UserPrincipalName | Out-String)"
                                    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                                }
                            }
                        }

                        # Verify the session count before shutting down the VM
                        if ($ExistingSession -eq 0) {
                                # Shutdown the Azure VM
                                $Message = "Stopping Azure VM: $VMName and waiting for it to complete."
                                Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                                Stop-SessionHost -VMName $VMName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -TimeDifference $TimeDifference
                        } else {
                            $Message = "Could not stop Azure VM $VMName, there are still $ExistingSession sessions connected to it."
                            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                        }
                    }

                    # Wait for the VM to stop. Added logic to prevent infinite loops
                    $WaitCycle=0
                    $IsVMStopped = $false
                    $Message = "Waiting for VM to stop: $($RoleInstance.Name)."
                    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                    while ((!$IsVMStopped) -and ($WaitCycle -lt $StartStopMaxCycles)) {
                        $WaitCycle += 1
                        $RoleInstance = Get-AzVM -Status | Where-Object { $_.Name -eq $VMName }
                        if ($RoleInstance.PowerState -eq "VM deallocated") {
                            $IsVMStopped = $true
                            $Message = "Azure VM has been stopped: $($RoleInstance.Name)."
                            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                        } else {
                            Start-Sleep -Seconds $StartStopWaitTime
                        }
                    }
                    if (!$IsVMStopped) {
                        $Message = "Could not stop Azure VM: $($RoleInstance.Name)."
                        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                    }

                    # Wait until the session host status is Unavailable, and remove it from drain mode
                    $Message = "Waiting for session host $SessionHostName to become unavailable to remove it from drain mode"
                    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                    $IsSessionHostNoHeartbeat = $false
                    $WaitCycle=0
                    while (!$IsSessionHostNoHeartbeat -and ($WaitCycle -lt $StartStopMaxCycles)) {
                        $WaitCycle += 1
                        $SessionHostInfo = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName -Name $SessionHostName
                        if ($SessionHostInfo.UpdateState -eq "Succeeded" -and $($SessionHostInfo.Status -eq "NoHeartbeat" -or $SessionHost.Status -eq "Unavailable" )) {
                            $IsSessionHostNoHeartbeat = $true
                            # Ensure the Azure VMs that are off have allow new connections mode set to True
                            if ($SessionHostInfo.AllowNewSession -eq $false) {
                                $Message = "Setting AllowNewConnections to true for session host $SessionHostName"
                                Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                                Set-AllowNewSession -TenantName $TenantName -HostPoolName $HostpoolName -SessionHostName $SessionHostName
                            }
                        } else {
                            Start-Sleep -Seconds $StartStopWaitTime
                        }
                    }
                    if (!$IsSessionHostNoHeartbeat) {
                        $Message = "Could not set AllowNewConnections to true for session host $SessionHostName"
                        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                    }

                    $RoleSize = Get-AzVMSize -Location $RoleInstance.Location | Where-Object { $_.Name -eq $RoleInstance.HardwareProfile.VmSize }
                    #decrement number of running session host
                    [int]$NumberOfRunningHost = [int]$NumberOfRunningHost - 1
                    [int]$TotalRunningCores = [int]$TotalRunningCores - $RoleSize.NumberOfCores
                }
            }
        }
    } # Breadth-first session hosts shutdown in off peak hours

    # Now store in an automation variable the new value of $MinimumNumberOfRDSH if we had already
    # However this value should have not changed at this point
    $AutomationAccount = Get-AzAutomationAccount -ErrorAction Stop | Where-Object { $_.AutomationAccountName -eq $AutomationAccountName }
    $OffPeakUsageMinimumNoOfRDSH = Get-AzAutomationVariable -Name "$HostpoolName-OffPeakUsage-MinimumNoOfRDSH" -ResourceGroupName $AutomationAccount.ResourceGroupName -AutomationAccountName $AutomationAccount.AutomationAccountName -ErrorAction SilentlyContinue
    if ($OffPeakUsageMinimumNoOfRDSH) {
        [int]$MinimumNumberOfRDSH = $OffPeakUsageMinimumNoOfRDSH.Value
        $NoConnectionsofhost = 0
        if ($NumberOfRunningHost -le $MinimumNumberOfRDSH) {
            foreach ($SessionHost in $AllSessionHosts) {
                if ($SessionHost.Status -eq "Available" -and $SessionHost.Sessions -eq 0) {
                    $NoConnectionsofhost = $NoConnectionsofhost + 1
                }
            }
            # Calculate how many more hosts we have above the minimum
            $NoConnectionsofhost = $NoConnectionsofhost-$DefinedMinimumNumberOfRDSH
            # Not too sure about this...
            if ($NoConnectionsofhost -gt $DefinedMinimumNumberOfRDSH) {
                [int]$MinimumNumberOfRDSH = [int]$MinimumNumberOfRDSH - $NoConnectionsofhost
                $Message = "Setting automation variable $HostpoolName-OffPeakUsage-MinimumNoOfRDSH to $MinimumNumberOfRDSH"
                Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                Set-AzAutomationVariable -Name "$HostpoolName-OffPeakUsage-MinimumNoOfRDSH" -ResourceGroupName $AutomationAccount.ResourceGroupName -AutomationAccountName $AutomationAccount.AutomationAccountName -Encrypted $false -Value $MinimumNumberOfRDSH
            }
        }
    }

    # Verify existing sessions, and if they are more than 90% of the maximum of the supported sessions, scale up
    # Added global variable $ScaleUpDuringOffPeak at the beginning of the script to control whether scale up takes place or not, because
    #   some customers where seeing the farm never going down during offpeak hours.
    $HostpoolMaxSessionLimit = $HostpoolInfo.MaxSessionLimit
    $HostpoolSessionCount = (Get-RdsUserSession -TenantName $TenantName -HostPoolName $HostpoolName).Count
    if ($ScaleUpDuringOffPeak -eq $true) {
        $Message = "The session count in host pool is $HostpoolSessionCount, the session limit is $HostpoolMaxSessionLimit. Verifying now whether scale up is required..."
    } else {
        $Message = "The session count in host pool is $HostpoolSessionCount, the session limit is $HostpoolMaxSessionLimit. Scale up during offpeak time is disabled"
    }
    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
    if ($HostpoolSessionCount -ne 0 -and $ScaleUpDuringOffpeak)
    {
        # Calculate the how many sessions will allow in minimum number of RDSH VMs in off peak hours and calculate TotalAllowSessions Scale Factor
        $TotalAllowSessionsInOffPeak = [int]$MinimumNumberOfRDSH * $HostpoolMaxSessionLimit
        $SessionsScaleFactor = $TotalAllowSessionsInOffPeak * 0.90
        $ScaleFactor = [math]::Floor($SessionsScaleFactor)
        # Scaling up if required
        if ($HostpoolSessionCount -ge $ScaleFactor) {
            $Message = "Scaling up the farm to accomodate for the high session count"
            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
            $AllSessionHosts = Get-RdsSessionHost -TenantName $TenantName -HostPoolName $HostpoolName | Where-Object { $_.Status -eq "NoHeartbeat" -or $_.Status -eq "Unavailable" }
            $AllSessionHosts = (Compare-Object -ReferenceObject $AllSessionHosts -DifferenceObject $SkipSessionhosts).InputObject | Select-Object -Property * -Unique
            foreach ($SessionHost in $AllSessionHosts) {
                # Check the session host status and if the session host is healthy before starting the host
                if ($SessionHost.UpdateState -eq "Succeeded") {
                    $SessionHostName = $SessionHost.SessionHostName | Out-String
                    $VMName = $SessionHostName.Split(".")[0]
                    $Message = "Existing number of sessions too close to the maximum session limit, starting VM $VMName..."
                    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                    # Validating session host is allowing new connections
                    Set-AllowNewSession -TenantName $TenantName -HostPoolName $HostpoolName -SessionHostName $SessionHost.SessionHostName
                    # Start the Az VM
                    $Message = "Starting Azure VM: $VMName and waiting for it to complete."
                    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                    Start-SessionHost -VMName $VMName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey -TimeDifference $TimeDifference
                    # Wait for the VM to start
                    $IsVMStarted = $false
                    while (!$IsVMStarted) {
                        $RoleInstance = Get-AzVM -Status | Where-Object { $_.Name -eq $VMName }
                        if ($RoleInstance.PowerState -eq "VM running") {
                            $IsVMStarted = $true
                            $Message = "Azure VM has been Started: $($RoleInstance.Name)."
                            Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                        }
                    }

                    # Wait for the sessionhost is available
                    $SessionHostIsAvailable = Wait-SessionHostAvailable -TenantName $TenantName -HostPoolName $HostpoolName -SessionHost $SessionHost.SessionHostName
                    if ($SessionHostIsAvailable) {
                        $Message = "$($SessionHost.SessionHostName | Out-String) session host status is 'Available'"
                        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                    } else {
                        $Message = "'$($SessionHost.SessionHostName | Out-String)' session host does not configured properly with deployagent or does not started properly"
                        Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                    }

                    # Increment the number of running session host
                    [int]$NumberOfRunningHost = [int]$NumberOfRunningHost + 1

                    # Increment the number of minimumnumberofrdsh and store it in an Azure Automation variable for the next time we run
                    [int]$MinimumNumberOfRDSH = [int]$MinimumNumberOfRDSH + 1
                    $OffPeakUsageMinimumNoOfRDSH = Get-AzAutomationVariable -Name "$HostpoolName-OffPeakUsage-MinimumNoOfRDSH" -ResourceGroupName $AutomationAccount.ResourceGroupName -AutomationAccountName $AutomationAccount.AutomationAccountName -ErrorAction SilentlyContinue
                    if (!$OffPeakUsageMinimumNoOfRDSH) {
                        New-AzAutomationVariable -Name "$HostpoolName-OffPeakUsage-MinimumNoOfRDSH" -ResourceGroupName $AutomationAccount.ResourceGroupName -AutomationAccountName $AutomationAccount.AutomationAccountName -Encrypted $false -Value $MinimumNumberOfRDSH -Description "Dynamically generated minimumnumber of RDSH value"
                    } else {
                        Set-AzAutomationVariable -Name "$HostpoolName-OffPeakUsage-MinimumNoOfRDSH" -ResourceGroupName $AutomationAccount.ResourceGroupName -AutomationAccountName $AutomationAccount.AutomationAccountName -Encrypted $false -Value $MinimumNumberOfRDSH
                    }

                    # Calculate available capacity of sessions
                    $RoleSize = Get-AzVMSize -Location $RoleInstance.Location | Where-Object { $_.Name -eq $RoleInstance.HardwareProfile.VmSize }
                    $AvailableSessionCapacity = $TotalAllowSessions + $HostpoolInfo.MaxSessionLimit
                    [int]$TotalRunningCores = [int]$TotalRunningCores + $RoleSize.NumberOfCores
                    $Message = "New available session capacity is: $AvailableSessionCapacity"
                    Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
                    break
                }
            }
        }
    }
}

# Good bye info
$Message = "HostpoolName: $HostpoolName, TotalRunningCores: $TotalRunningCores NumberOfRunningHosts: $NumberOfRunningHost"
Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
$Message = "End WVD tenant scale optimization."
Send-Message -Message $Message -HostPoolName $HostpoolName -LogAnalyticsWorkspaceId $LogAnalyticsWorkspaceId -LogAnalyticsPrimaryKey $LogAnalyticsPrimaryKey
