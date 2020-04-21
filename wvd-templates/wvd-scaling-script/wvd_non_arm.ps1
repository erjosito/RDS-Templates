###################################
#
# Command samples for WVD
#
###################################

# Check AZ CLI is logged in, and it is the right sub
az account show

# Secrets for WVD Admin app SP
$wvdadmin_sp_appid = $(az keyvault secret show --vault-name cloudtrooper --name wvdadmin-sp-appid --query value -o tsv)
$wvdadmin_sp_secret = $(az keyvault secret show --vault-name cloudtrooper --name wvdadmin-sp-secret --query value -o tsv)

# Create SP (it only needs to be done once)
Import-Module AzureAD
$aadContext = Connect-AzureAD
$svcPrincipal = New-AzureADApplication -AvailableToOtherTenants $true -DisplayName "Windows Virtual Desktop Svc Principal"
$svcPrincipalCreds = New-AzureADApplicationPasswordCredential -ObjectId $svcPrincipal.ObjectId

# Secrets for TenantCreator SP
wvd_sp_appid=$(az keyvault secret show --vault-name cloudtrooper --name wvd-sp-appid --query value -o tsv)
wvd_sp_secret=$(az keyvault secret show --vault-name cloudtrooper --name wvd-sp-secret --query value -o tsv)

# Login to WVD (admin@cloudtrooper.net)
Add-RdsAccount -DeploymentUrl "https://rdbroker.wvd.microsoft.com"

# Create Tenant
New-RdsTenant -Name cloudtrooper -AadTenantId ecd38d6d-544b-494c-9b29-ff3d6a31c040 -AzureSubscriptionId 3e78e84b-6750-44b9-9d57-d9bba935237a
$tenant = Get-RdsTenant

# Tenant role assignments
New-RdsRoleAssignment -RoleDefinitionName "RDS Owner" -ApplicationId $svcPrincipal.AppId -TenantName $tenant.TenantName
Get-RdsRoleAssignment 

# Test SP sign in
$creds = New-Object System.Management.Automation.PSCredential($svcPrincipal.AppId, (ConvertTo-SecureString $svcPrincipalCreds.Value -AsPlainText -Force))
Add-RdsAccount -DeploymentUrl "https://rdbroker.wvd.microsoft.com" -Credential $creds -ServicePrincipal -AadTenantId $aadContext.TenantId.Guid

# Host pools
$hostpool = Get-RdsHostPool -TenantName $tenant.TenantName

# App group
$appgroup = Get-RdsAppGroup -TenantName $tenant.TenantName -HostPoolName $hostpool.HostPoolName

# New app & app group:
$appgroupname = "myapps"
New-RdsAppGroup $tenant.TenantName $hostpool.HostPoolName $appgroupname -ResourceType "RemoteApp"
Get-RdsStartMenuApp $tenant.TenantName $hostpool.HostPoolName $appgroupname
New-RdsRemoteApp $tenant.TenantName $hostpool.HostPoolName $appgroupname -Name WordPad -AppAlias wordpad
Get-RdsRemoteApp $tenant.TenantName $hostpool.HostPoolName $appgroupname
$upn = "julia@cloudtrooper.net"
Add-RdsAppGroupUser -TenantName $tenant.TenantName `
                    -HostPoolName $hostpool.HostPoolName `
                    -AppGroupName $appgroupname `
                    -UserPrincipalName $upn

# Users in app group
$upn = "lucas@cloudtrooper.net"
Get-RdsAppGroupUser -TenantName $tenant.TenantName -HostPoolName $hostpool.HostPoolName -AppGroupName $appgroup.AppGroupName
Add-RdsAppGroupUser -TenantName $tenant.TenantName `
                                                 -HostPoolName $hostpool.HostPoolName `
                    -AppGroupName $appgroup.AppGroupName `
                                                 -UserPrincipalName $upn

# Session hosts
Get-RdsSessionHost $tenant.TenantName $hostpool.HostPoolName 

# Host pool properties
$property = "audiocapturemode:i:1;"
$property = "audiomode:i:0;audiocapturemode:i:1;"
Set-RdsHostPool -TenantName $tenant.TenantName -Name $hostpool.HostPoolName -CustomRdpProperty $property
Set-RdsHostPool -TenantName $tenant.TenantName -Name $hostpool.HostPoolName -BreadthFirstLoadBalancer -MaxSessionLimit 10
Set-RdsHostPool -TenantName $tenant.TenantName -Name $hostpool.HostPoolName -DepthFirstLoadBalancer -MaxSessionLimit 10
Set-RdsHostPool $tenant.TenantName $hostpool.HostPoolName -AssignmentType Automatic|Direct
Set-RdsSessionHost $tenant.TenantName $hostpool.HostPoolName -Name <sessionhostname> -AssignedUser $upn

# Sending data to Log Analytics
$subscriptionId = $(az account show --query id -o tsv)
Set-RdsTenant -Name $tenant.TenantName -AzureSubscriptionId $subscriptionId -LogAnalyticsWorkspaceId <String> -LogAnalyticsPrimaryKey <String>



# Mounting NFS shares:
# As admin:
net use X: \\wvdfslogixcloudtrooper.file.core.windows.net\fslogix <storage-account-key> /user:Azure\wvdfslogixcloudtrooper
# As user:
net use X: \\wvdfslogixcloudtrooper.file.core.windows.net\fslogix\lucas
# Permissions
icacls z: /grant <user-email>:(F)
icacls z: /grant wvd-users:(M)
icacls z: /grant "Creator Owner":(OI)(CI)(IO)(M)



# Creating logic app for automation
# $resourceGroupName = Read-Host -Prompt "Enter the name of the resource group for the new Azure Logic App"
$resourceGroupName = "wvdautoscale"
# $aadTenantId = Read-Host -Prompt "Enter your Azure AD tenant ID"
$aadTenantId = "ecd38d6d-544b-494c-9b29-ff3d6a31c040"
# $subscriptionId = Read-Host -Prompt "Enter your Azure Subscription ID"
$subscriptionId = "3e78e84b-6750-44b9-9d57-d9bba935237a"
# $tenantName = Read-Host -Prompt "Enter the name of your WVD tenant"
$tenantName = "cloudtrooper"
# $hostPoolName = Read-Host -Prompt "Enter the name of the host pool you'd like to scale"
$hostPoolName = "cloudtrooper"
# $recurrenceInterval = Read-Host -Prompt "Enter how often you'd like the job to run in minutes, e.g. '15'"
$recurrenceInterval = "120"
# $beginPeakTime = Read-Host -Prompt "Enter the start time for peak hours in local time, e.g. 9:00"
$beginPeakTime = "5:00"
# $endPeakTime = Read-Host -Prompt "Enter the end time for peak hours in local time, e.g. 18:00"
$endPeakTime = "5:15"
# $timeDifference = Read-Host -Prompt "Enter the time difference between local time and UTC in hours, e.g. +5:30"
$timeDifference = "+2"
# $sessionThresholdPerCPU = Read-Host -Prompt "Enter the maximum number of sessions per CPU that will be used as a threshold to determine when new session host VMs need to be started during peak hours"
$sessionThresholdPerCPU = 0.5
# $minimumNumberOfRdsh = Read-Host -Prompt "Enter the minimum number of session host VMs to keep running during off-peak hours"
$minimumNumberOfRdsh = 0
# $limitSecondsToForceLogOffUser = Read-Host -Prompt "Enter the number of seconds to wait before automatically signing out users. If set to 0, users will be signed out immediately"
$limitSecondsToForceLogOffUser = 120
# $logOffMessageTitle = Read-Host -Prompt "Enter the title of the message sent to the user before they are forced to sign out"
$logOffMessageTitle = "Please sign off, this machine is going to be powered down"
# $logOffMessageBody = Read-Host -Prompt "Enter the body of the message sent to the user before they are forced to sign out"
$logOffMessageBody = "Please sign off, this machine is going to be powered down"
# $location = Read-Host -Prompt "Enter the name of the Azure region where you will be creating the logic app"
$location = "westeurope"
# $connectionAssetName = Read-Host -Prompt "Enter the name of the Azure RunAs connection asset"
$connectionAssetName = "AzureRunAsConnection"
# $webHookURI = Read-Host -Prompt "Enter the URI of the WebHook returned by when you created the Azure Automation Account"
$webHookURI = "https://s2events.azure-automation.net/webhooks?token=xyKqBttIQD1TnBcjob4LlIf62wzjITlMNwkOZ6g2tTc%3d"
# $automationAccountName = Read-Host -Prompt "Enter the name of the Azure Automation Account"
$automationAccountName = "wvd1138"
# $maintenanceTagName = Read-Host -Prompt "Enter the name of the Tag associated with VMs you don't want to be managed by this scaling tool"
$maintenanceTagName = "DoNotPowerOff"
.\createazurelogicapp.ps1 -ResourceGroupName $resourceGroupName `
  -AADTenantID $aadTenantId `
  -SubscriptionID $subscriptionId `
  -TenantName $tenantName `
  -HostPoolName $hostPoolName `
  -RecurrenceInterval $recurrenceInterval `
  -BeginPeakTime $beginPeakTime `
  -EndPeakTime $endPeakTime `
  -TimeDifference $timeDifference `
  -SessionThresholdPerCPU $sessionThresholdPerCPU `
  -MinimumNumberOfRDSH $minimumNumberOfRdsh `
  -LimitSecondsToForceLogOffUser $limitSecondsToForceLogOffUser `
  -LogOffMessageTitle $logOffMessageTitle `
  -LogOffMessageBody $logOffMessageBody `
  -Location $location `
  -ConnectionAssetName $connectionAssetName `
  -WebHookURI $webHookURI `
  -AutomationAccountName $automationAccountName `
  -MaintenanceTagName $maintenanceTagName

