#Requires -RunAsAdministrator 
#Requires -Version 7.0

<#
.SYNOPSIS
Configure Windows Time Service to synchronize time with NTP servers.

.DESCRIPTION
This script configures Windows Time Service (w32time) to synchronize time with public NTP peer servers.
This script allows you to select NTP sources from Facebook, Google, or NTP Pool Project.
Be aware that this script will reset and reconfigure the Windows Time Service.
This script logs the output to a log file "%TEMP%\Configure-NTPSync.ps1.yyyyMMddHHmmss.log".

Author: https://github.com/mielipuolinen
Disclaimer: No warranty expressed or implied. Use at your own risk.

.PARAMETER NTPSource
Specifies the NTP source to use for time synchronization.
Valid values are "Facebook", "Google", and "NTPPool". Default value is "Facebook".

.PARAMETER SkipTimeAPI
Skips the TimeAPI.io check to compare the system clock with the current time from TimeAPI.io.

.PARAMETER Unattended
Runs the script in unattended mode. The script will exit after completion without user interaction.

.EXAMPLE
Configure Windows Time Service with Facebook's (set as default) NTP peer servers as a source:

Run Powershell 7 as Administrator > Copy&Paste the command below > Press Enter
irm https://raw.githubusercontent.com/mielipuolinen/PowerShell-Scripts/master/Configure-NTPSync.ps1 | iex

.EXAMPLE
Configure Windows Time service with Google's NTP peer servers as a source:

Run Powershell 7 as Administrator > Copy&Paste the command below > Press Enter
$NTP="Google"; irm https://raw.githubusercontent.com/mielipuolinen/PowerShell-Scripts/master/Configure-NTPSync.ps1 > "$env:TEMP\Configure-NTPSync.ps1"; & "$env:TEMP\Configure-NTPSync.ps1" -NTPSource $NTP; rm "$env:TEMP\Configure-NTPSync.ps1"

.EXAMPLE
Configure Windows Time service with NTP Pool Project's (pool.ntp.org) NTP peer servers as a source:

Run Powershell 7 as Administrator > Copy&Paste the command below > Press Enter
$NTP="NTPPool"; irm https://raw.githubusercontent.com/mielipuolinen/PowerShell-Scripts/master/Configure-NTPSync.ps1 > "$env:TEMP\Configure-NTPSync.ps1"; & "$env:TEMP\Configure-NTPSync.ps1" -NTPSource $NTP; rm "$env:TEMP\Configure-NTPSync.ps1"


.INPUTS
No pipeline inputs are accepted.

.OUTPUTS
No forwardable outputs are generated.

.LINK
https://github.com/mielipuolinen/PowerShell-Scripts/blob/master/Configure-NTPSync.ps1

.NOTES
Version: 1.2
Date: 2025-03-16
Usage: See examples. PowerShell 7.0 or later is required. PowerShell must be run as Administrator.
Author: https://github.com/mielipuolinen
Disclaimer: No warranty expressed or implied. Use at your own risk.

https://learn.microsoft.com/en-us/windows-server/networking/windows-time-service/windows-time-service-tools-and-settings
https://learn.microsoft.com/en-us/troubleshoot/windows-server/active-directory/configure-w32ime-against-huge-time-offset
https://engineering.fb.com/2020/03/18/production-engineering/ntp-service/
https://serverfault.com/questions/334682/ws2008-ntp-using-time-windows-com-0x9-time-always-skewed-forwards
https://developers.google.com/time/smear
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Facebook", "Google", "NTPPool", IgnoreCase=$true)]
    [string]$NTPSource = "Facebook",
    [Parameter(Mandatory=$false)]
    [switch]$SkipTimeAPI = $false,
    [Parameter(Mandatory=$false)]
    [switch]$Unattended = $false
)

Set-StrictMode -Version 3.0
# Prohibits references to uninitialized variables. This includes uninitialized variables in strings.
# Prohibits references to non-existent properties of an object.
# Prohibits function calls that use the syntax for calling methods.
# Prohibit out of bounds or unresolvable array indexes.
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/set-strictmode

$ScriptVersion = "v1.2"

####################################################################################################
Write-Host @'

    _____                __  _                             _   _  _______  _____      _____                       
   / ____|              / _|(_)                           | \ | ||__   __||  __ \    / ____|                      
  | |      ___   _ __  | |_  _   __ _  _   _  _ __  ___   |  \| |   | |   | |__) |  | (___   _   _  _ __    ___   
  | |     / _ \ | '_ \ |  _|| | / _` || | | || '__|/ _ \  | . ` |   | |   |  ___/    \___ \ | | | || '_ \  / __|  
  | |____| (_) || | | || |  | || (_| || |_| || |  |  __/  | |\  |   | |   | |        ____) || |_| || | | || (__   
   \_____|\___/ |_| |_||_|  |_| \__, | \__,_||_|   \___|  |_| \_|   |_|   |_|       |_____/  \__, ||_| |_| \___|  
                                 __/ |                                                        __/ |               
'@ -ForegroundColor Cyan

Write-Host "                                |___/" -ForegroundColor Cyan -NoNewline
Write-Host "     g i t h u b . c o m / m i e l i p u o l i n e n    " -ForegroundColor Magenta -NoNewline
Write-Host "|___/          $($ScriptVersion)" -ForegroundColor Cyan
Write-Host ""

$ScriptLogFile = "$env:TEMP\Configure-NTPSync.ps1.$(Get-Date -f "yyyyMMddHHmmss").log"
Start-Transcript -Path $ScriptLogFile -NoClobber -Append | Out-Null

# Writing to the log file, invisible in the console
$CursorPosition = $Host.UI.RawUI.CursorPosition
Write-Host "" -NoNewline
Write-Host "Configure NTP Sync $($ScriptVersion)" -NoNewline
$Host.UI.RawUI.CursorPosition = $CursorPosition
Write-Host "github.com/mielipuolinen" -NoNewline
$Host.UI.RawUI.CursorPosition = $CursorPosition
Write-Host "                                                  "

####################################################################################################
# Set registry key paths
$RegKeyPath_W32TimeConfig = "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config"
$RegKeyPath_W32TimeParameters = "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"
$RegKeyPath_W32TimeProviders= "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient"

# Invoke-Expression commands
$Cmd_W32TM_Config = "w32tm /query /configuration"
$Cmd_W32TM_Status = "w32tm /query /status"
$Cmd_W32TM_Peers = "w32tm /query /peers"
$Cmd_W32TM_TZ = "w32tm /tz"


####################################################################################################
# Select NTP Peers
Switch ($NTPSource) {
    "Facebook" {
        Write-Host "`n🌐 Selected Facebook's NTP peers as a time source" -ForegroundColor Green
        $NTPPeerList = "time1.facebook.com,0x8 time2.facebook.com,0x8 time3.facebook.com,0x8 time4.facebook.com,0x8 time5.facebook.com,0x8"
        $StriptChartServer = "time.facebook.com"
    }
    "Google" {
        Write-Host "`n🌐 Selected Google's NTP peers as a time source" -ForegroundColor Green
        $NTPPeerList = "time1.google.com,0x8 time2.google.com,0x8 time3.google.com,0x8 time4.google.com,0x8"
        $StriptChartServer = "time.google.com"
    }
    "NTPPool" {
        Write-Host "`n🌐 Selected NTP Pool Project's NTP peers as a time source" -ForegroundColor Green
        $NTPPeerList = "0.pool.ntp.org,0x8 1.pool.ntp.org,0x8 2.pool.ntp.org,0x8 3.pool.ntp.org,0x8"
        $StriptChartServer = "pool.ntp.org"
    }
    default { Write-Error "Unable to determine NTP source" -ErrorAction Stop }
}

# List Selected NTP Peer Servers
$NTPPeerList -split " " | ForEach-Object {
    Write-Host "`t$_" -ForegroundColor Cyan
}


####################################################################################################
# Check current date and time in Windows, and determine if it's about accurate without w32time service
# TODO: Refactor, use Invoke-WebRequest and read headers to get the current time from trusted sources (google, MS, etc.)
Write-Host "`n🔍 Checking if System clock is approximately on time" -ForegroundColor Green

if(!$SkipTimeAPI){

    $WebRequest_MaxAttempts = 2
    $WebRequest_Attempt = 0
    $WebRequest_RetryDelay = 5 #seconds
    $DateTime_TimeAPI = $null

    while (($WebRequest_Attempt -lt $WebRequest_MaxAttempts) -and ($null -eq $DateTime_TimeAPI)) {

        # Doubling as a poor man's rate limiter
        Start-Sleep -Seconds ($WebRequest_RetryDelay*$WebRequest_Attempt)

        try {
            Write-Host "`t🔗 Connecting to TimeAPI.io (attempt #$($WebRequest_Attempt+1))" -ForegroundColor Cyan
            $WebRequest = Invoke-WebRequest -Uri "https://timeapi.io/api/time/current/zone?timeZone=UTC" `
                        -SkipCertificateCheck -ConnectionTimeoutSeconds 10 -ErrorAction Stop
            Write-Host "`t✅ Connected to TimeAPI.io" -ForegroundColor Cyan
            $DateTime_TimeAPI = ($WebRequest.Content | ConvertFrom-Json).dateTime
            $DateTime_System = Get-Date -AsUTC
            $DateTime_Difference = New-TimeSpan -Start $DateTime_System -End $DateTime_TimeAPI
        } finally { $WebRequest_Attempt++ }
    }

    if ($null -eq $DateTime_TimeAPI) {
        Write-Host "`t⚠️ Unable to compare clock sources: TimeAPI.io service is unavailable" -ForegroundColor Yellow
    } else {
        
        Write-Host "`tTimeAPI.io clock: $DateTime_TimeAPI" -ForegroundColor Cyan
        Write-Host "`tSystem clock: $DateTime_System" -ForegroundColor Cyan
        Write-Host "`tDifference: $($DateTime_Difference.TotalSeconds) seconds" -ForegroundColor Cyan

        $AcceptableTimeDifference = 1 # +- seconds
        if ($DateTime_Difference.TotalSeconds -gt $AcceptableTimeDifference) {
            Write-Host "`t⚠️ System clock is out of sync" -ForegroundColor Yellow
        } elseif ($DateTime_Difference.TotalSeconds -lt $AcceptableTimeDifference*-1) {
            Write-Host "`t⚠️ System clock is out of sync" -ForegroundColor Yellow
        } else {
            Write-Host "`t✅ System clock is approximately on time" -ForegroundColor Cyan
        }
    }

}else{ Write-Host "`t⚠️ User requested to skip TimeAPI.io check" -ForegroundColor Yellow }

####################################################################################################
# Check if Windows Time Service exists
Write-Host "`n🔍 Checking Windows Time Service" -ForegroundColor Green

$Service_W32Time = Get-Service -Name "w32time" -ErrorAction SilentlyContinue
if ($null -eq $Service_W32Time) {
    Write-Host "`t⚠️ Windows Time Service does not exist on this system" -ForegroundColor Yellow
    Write-Host "`t🔄 Registering Windows Time Service" -ForegroundColor Cyan
    Write-Host "`t$(Invoke-Expression "w32tm /register")" -ForegroundColor Cyan
    Start-Sleep -Seconds 5
} else {
    Write-Host "`t✅ Windows Time Service exists" -ForegroundColor Cyan
}

# Check if Windows Time Service startup type is set to Manual
$Service_W32Time = Get-Service -Name "w32time" -ErrorAction Stop
if ($Service_W32Time.StartType -ne "Manual") {
    try{
        Write-Host "`t⚠️ Windows Time Service startup type is set to: $($Service_W32Time.StartType)" -ForegroundColor Yellow
        Write-Host "`t🔄 Setting Windows Time Service startup type to Manual" -ForegroundColor Cyan
        Set-Service -Name w32time -StartupType Manual
        Start-Sleep -Seconds 5
    } catch {
        Write-Error "Unable to set Windows Time Service startup type to Automatic: $($_.Exception.Message)" -ErrorAction Stop
    }
} else {
    Write-Host "`t✅ Windows Time Service startup type is set to Manual" -ForegroundColor Cyan
}

# Start Windows Time Service if not running
if ((Get-Service -Name "w32time").Status -ne "Running") {
    try{
        Write-Host "`t⚠️ Windows Time Service is not running" -ForegroundColor Yellow
        Write-Host "`t🔄 Starting Windows Time Service" -ForegroundColor Cyan
        Start-Service -Name w32time -ErrorAction Stop
        Start-Sleep -Seconds 5
    } catch {
        Write-Error "Unable to start Windows Time Service: $($_.Exception.Message)" -ErrorAction Stop
    }
} else {
    Write-Host "`t✅ Windows Time Service is running" -ForegroundColor Cyan
}


####################################################################################################
# Create a strip chart
# Dsplay the time offset between the local computer and the selected NTP peer servers
# /stripchart: Displays a strip chart of the offset between the local computer and the target computer.
# /period: Specifies the time interval between samples in seconds.
# /samples: Specifies the number of samples to collect.
Write-Host "`n🔍 Creating Strip Chart" -ForegroundColor Green
(Invoke-Expression "w32tm /stripchart /computer:$StriptChartServer /period:1 /samples:5") -split "`n" | ForEach-Object {
    Write-Host "`t$_" -ForegroundColor Cyan
}


####################################################################################################
# Capture current configuration and status
Write-Host "`n🔍 Reading Windows Time Service's status before applying configurations" -ForegroundColor Green
Write-Host "`t⚡ $Cmd_W32TM_Config" -ForegroundColor Cyan
$W32TM_Config_Before = Invoke-Expression $Cmd_W32TM_Config
Write-Host "`t⚡ $Cmd_W32TM_Status" -ForegroundColor Cyan
$W32TM_Status_Before = Invoke-Expression $Cmd_W32TM_Status
Write-Host "`t⚡ $Cmd_W32TM_Peers" -ForegroundColor Cyan
$W32TM_Peers_Before = Invoke-Expression $Cmd_W32TM_Peers
Write-Host "`t⚡ $Cmd_W32TM_TZ" -ForegroundColor Cyan
$W32TM_TZ_Before = Invoke-Expression $Cmd_W32TM_TZ

# List peers before applying configurations
Write-Host "`n🌐 NTP Peers before applying configurations" -ForegroundColor Green
$W32TM_Peers_Before -split "`n" | ForEach-Object {
    Write-Host "`t$_" -ForegroundColor Cyan
}

# List timezone before applying configurations
Write-Host "`n🕒 Timezone before applying configurations" -ForegroundColor Green
$W32TM_TZ_Before -split "`n" | ForEach-Object {
    Write-Host "`t$_" -ForegroundColor Cyan
}


####################################################################################################
# Unregisters the time service, and removes all configuration information from the registry.
Write-Host "`n🔧 Unregistering Windows Time Service and removing any existing configurations" -ForegroundColor Green
Write-Host "`t$(Invoke-Expression "w32tm /unregister")" -ForegroundColor Cyan
Start-Sleep -Seconds 5

# Stop Windows Time Service
Write-Host "`n🔄 Stopping Windows Time Service" -ForegroundColor Green
try{ Stop-Service -Name w32time } catch{}
# .{Stop-Service -Name w32time} 2>$null # Hack to suppress an expected error message, which doesn't follow the ErrorAction preference.

# Registers the time service to run as a service, and adds default configuration to the registry.
Write-Host "`n🔧 Registering Windows Time Service and applying default configurations" -ForegroundColor Green
Write-Host "`t$(Invoke-Expression "w32tm /register")" -ForegroundColor Cyan
Start-Sleep -Seconds 5


####################################################################################################
# UtilizeSSlTimeData
# Specifies whether the Windows Time service uses SSL time data that is received from the time source.
# Default value is 1.
# https://learn.microsoft.com/en-us/windows-server/networking/windows-time-service/windows-server-2016-improvements#secure-time-seeding
# https://arstechnica.com/security/2023/08/windows-feature-that-resets-system-clocks-based-on-random-data-is-wreaking-havoc/
# https://serverfault.com/questions/1131670/windows-server-time-service-jumps-into-the-future-and-partially-back
Write-Host "`n🔧 Disable SSL time data utilization" -ForegroundColor Green
Write-Host "`tRegistry Key: $RegKeyPath_W32TimeConfig" -ForegroundColor Cyan

try {
    $UtilizeSslTimeData = (Get-ItemProperty -Path $RegKeyPath_W32TimeConfig -Name "UtilizeSslTimeData")
    $UtilizeSslTimeData = $UtilizeSslTimeData.UtilizeSslTimeData
}
catch { $UtilizeSslTimeData = -1 }

if($UtilizeSslTimeData -eq 0){
    Write-Host "`tSSL time data utilization is already disabled" -ForegroundColor Cyan
}else{
    Write-Host "`tUtilizeSSlTimeData: 0" -ForegroundColor Cyan
    Set-ItemProperty -Path $RegKeyPath_W32TimeConfig -Name "UtilizeSslTimeData" -Value 1 -Type DWord
}

# Allow time correction up to +-25 hours
#   Daylight saving time bugs can cause 1-hour time differences.
#   AM or PM misconfiguration can cause a 12-hour time difference.
#   Day or date mistakes can cause a 24-hour time difference.
# MaxPosPhaseCorrection: Specifies the maximum positive time correction in seconds that the service can make.
# MaxNegPhaseCorrection: Specifies the maximum negative time correction in seconds that the service can make.
# Setting both to 90000 seconds (25 hours) allows a maximum correction of 25 hours.
Write-Host "`n🔧 Configuring maximum time correction limits" -ForegroundColor Green
Write-Host "`tRegistry Key: $RegKeyPath_W32TimeConfig" -ForegroundColor Cyan
Write-Host "`tMaxPosPhaseCorrection: 90000 seconds (25 hours)" -ForegroundColor Cyan
Set-ItemProperty -Path $RegKeyPath_W32TimeConfig -Name "MaxPosPhaseCorrection" -Value 90000 -Type DWord
Write-Host "`tMaxNegPhaseCorrection: 90000 seconds (25 hours)" -ForegroundColor Cyan
Set-ItemProperty -Path $RegKeyPath_W32TimeConfig -Name "MaxNegPhaseCorrection" -Value 90000 -Type DWord

# Change MinPollInterval to 2^9 (~8,5 minutes, 512 seconds)
# Default value is 2^10 (~17 minutes, 1024 seconds)
# Value is represented in base 2, so 2^9 = 512 seconds
# The minimum polling interval is the shortest time that the Windows Time service will wait between time synchronization attempts.
$RegKey_W32TimeConfig_Value_MinPollInterval = 9 # 2^9 = 512 seconds
$RegKey_W32TimeConfig_Value_MinPollInterval_InSeconds = [math]::Pow(2, $RegKey_W32TimeConfig_Value_MinPollInterval)
Write-Host "`n🔧 Configuring NTP minimum polling interval" -ForegroundColor Green
Write-Host "`tRegistry Key: $RegKeyPath_W32TimeConfig" -ForegroundColor Cyan
Write-Host "`tMinPollInterval: $RegKey_W32TimeConfig_Value_MinPollInterval_InSeconds seconds" -ForegroundColor Cyan
Set-ItemProperty -Path $RegKeyPath_W32TimeConfig -Name "MinPollInterval" -Type DWord `
    -Value $RegKey_W32TimeConfig_Value_MinPollInterval

# Change MaxPollInterval to 2^14 (~4,5 hours, 16384 seconds)
# Default value is 2^15 (~9.1 hours, 32768 seconds)
# Value is represented in base 2, so 2^9 = 512 seconds
# The maximum polling interval is the longest time that the Windows Time service will wait between time synchronization attempts.
$RegKey_W32TimeConfig_Value_MaxPollInterval = 14 # 2^14 = 16384 seconds
$RegKey_W32TimeConfig_Value_MaxPollInterval_InSeconds = [math]::Pow(2, $RegKey_W32TimeConfig_Value_MaxPollInterval)
Write-Host "`n🔧 Configuring NTP maximum polling interval" -ForegroundColor Green
Write-Host "`tRegistry Key: $RegKeyPath_W32TimeConfig" -ForegroundColor Cyan
Write-Host "`tMaxPollInterval: $RegKey_W32TimeConfig_Value_MaxPollInterval_InSeconds seconds" -ForegroundColor Cyan
Set-ItemProperty -Path $RegKeyPath_W32TimeConfig -Name "MaxPollInterval" -Type DWord `
    -Value $RegKey_W32TimeConfig_Value_MaxPollInterval

# Set Client Type to NTP
# Default client NT5DS, the client synchronizes time with a domain controller in the domain hierarchy
Write-Host "`n🔧 Configuring NTP client type" -ForegroundColor Green
Set-ItemProperty -Path $RegKeyPath_W32TimeParameters -Name Type -Value "NTP" -Type String

# Enable NtpClient
Write-Host "`n🔧 Enabling NTP client" -ForegroundColor Green
Set-ItemProperty -Path $RegKeyPath_W32TimeProviders -Name Enabled -Value 1 -Type DWord

# Start Windows Time Service
Write-Host "`n🔄 Starting Windows Time Service" -ForegroundColor Green
Start-Service -Name w32time -ErrorAction Stop
Start-Sleep -Seconds 5


####################################################################################################
# Configure NTP servers
# /config: Modifies the configuration of the Windows Time service.
# /update: Notifies the Windows Time service that the configuration changed, causing the changes to take effect.
# /manualpeerlist:<peers>: Specifies the list of peers from which the Windows Time service obtains time stamps.
# /syncfromflags:MANUAL: Specifies that the Windows Time service is to use the manual peer list when synchronizing time.
# /reliable:NO: Set whether this computer is a reliable time source. This setting is only meaningful on DCs.
Write-Host "`n🔧 Configuring NTP servers" -ForegroundColor Green
$Cmd_W32TM_ConfigManualPeerList = "w32tm /config /manualpeerlist:`"$NTPPeerList`" /syncfromflags:MANUAL /reliable:NO /update"
Write-Host "`t$(Invoke-Expression $Cmd_W32TM_ConfigManualPeerList)" -ForegroundColor Cyan

# Restart Windows Time Service
Write-Host "`n🔄 Restarting Windows Time Service" -ForegroundColor Green
Restart-Service -Name w32time -ErrorAction Stop
Start-Sleep -Seconds 5


####################################################################################################
# Resynchronize the computer clock and rediscover network sources
Write-Host "`n🔄 Resynchronizing System clock" -ForegroundColor Green
# /resync: Synchronizes the computer clock with the time source, and then checks the time source for accuracy.
# /rediscover: Redetects the network configuration and rediscovers network sources, then resynchronizes.
#   Redetect Network Configuration: The Windows Time service will check the current network settings and configurations.
#   Rediscover Network Sources: It will search for available NTP servers or other time sources based on the updated network configuration.
#   Resynchronize: The service will then synchronize the computer clock with the newly discovered time sources.
#   This is particularly useful if there have been changes in the network environment, such as new NTP servers being added, changes in network topology, or updates to DNS settings.
(Invoke-Expression "w32tm /resync /rediscover") -split "`n" | ForEach-Object {
    Write-Host "`t$_" -ForegroundColor Cyan
}
Start-Sleep -Seconds 5


####################################################################################################
# Create a strip chart
# Dsplay the time offset between the local computer and the selected NTP peer servers
# /stripchart: Displays a strip chart of the offset between the local computer and the target computer.
# /period: Specifies the time interval between samples in seconds.
# /samples: Specifies the number of samples to collect.
Write-Host "`n🔍 Creating Strip Chart" -ForegroundColor Green
(Invoke-Expression "w32tm /stripchart /computer:$StriptChartServer /period:1 /samples:5") -split "`n" | ForEach-Object {
    Write-Host "`t$_" -ForegroundColor Cyan
}

# Capture current configuration and status
Write-Host "`n🔍 Reading Windows Time Service's status after applying configurations" -ForegroundColor Green
Write-Host "`t⚡ $Cmd_W32TM_Config" -ForegroundColor Cyan
$W32TM_Config_After = Invoke-Expression $Cmd_W32TM_Config
Write-Host "`t⚡ $Cmd_W32TM_Status" -ForegroundColor Cyan
$W32TM_Status_After = Invoke-Expression $Cmd_W32TM_Status
Write-Host "`t⚡ $Cmd_W32TM_Peers" -ForegroundColor Cyan
$W32TM_Peers_After = Invoke-Expression $Cmd_W32TM_Peers
Write-Host "`t⚡ $Cmd_W32TM_TZ" -ForegroundColor Cyan
$W32TM_TZ_After = Invoke-Expression $Cmd_W32TM_TZ

# List peers after applying configurations
Write-Host "`n🌐 NTP Peers after applying configurations" -ForegroundColor Green
$W32TM_Peers_After -split "`n" | ForEach-Object {
    Write-Host "`t$_" -ForegroundColor Cyan
}

# List timezone after applying configurations
Write-Host "`n🕒 Timezone after applying configurations" -ForegroundColor Green
$W32TM_TZ_After -split "`n" | ForEach-Object {
    Write-Host "`t$_" -ForegroundColor Cyan
}


####################################################################################################
# Before & After: Windows Time Service status
Write-Host "`n🔍 Before & After: Windows Time Service configuration" -ForegroundColor Green
Write-Host "`tw32tm /query /configuration" -ForegroundColor Cyan

$W32TM_Config_Before = $W32TM_Config_Before | Where-Object { $_ -ne "" }
$W32TM_Config_After = $W32TM_Config_After | Where-Object { $_ -ne "" }
$W32TM_Config_Before_LineCount = $W32TM_Config_Before.Count
$W32TM_Config_After_LineCount = $W32TM_Config_After.Count
$W32TM_Config_MaxLineCount = [math]::Max($W32TM_Config_Before_LineCount, $W32TM_Config_After_LineCount)
$W32TM_Config_BeforeAndAfter = @()

for ($i = 0; $i -lt $W32TM_Config_MaxLineCount; $i++) {
    $W32TM_Config_BeforeAndAfter += [PSCustomObject]@{
        "Before" = if ($i -lt $W32TM_Config_Before_LineCount) { $W32TM_Config_Before[$i] } else { "" }
        "After" = if ($i -lt $W32TM_Config_After_LineCount) { $W32TM_Config_After[$i] } else { "" }
    }
}

$W32TM_Config_BeforeAndAfter | Format-Table -Wrap


####################################################################################################
# Before & After: Windows Time Service status
Write-Host "`n🔍 Before & After: Windows Time Service status" -ForegroundColor Green
Write-Host "`tw32tm /query /status" -ForegroundColor Cyan

$W32TM_Status_Before = $W32TM_Status_Before | Where-Object { $_ -ne "" }
$W32TM_Status_After = $W32TM_Status_After | Where-Object { $_ -ne "" }
$W32TM_Status_Before_LineCount = $W32TM_Status_Before.Count
$W32TM_Status_After_LineCount = $W32TM_Status_After.Count
$W32TM_Status_MaxLineCount = [math]::Max($W32TM_Status_Before_LineCount, $W32TM_Status_After_LineCount)
$W32TM_Status_BeforeAndAfter = @()

for ($i = 0; $i -lt $W32TM_Status_MaxLineCount; $i++) {
    $W32TM_Status_BeforeAndAfter += [PSCustomObject]@{
        "Before" = if ($i -lt $W32TM_Status_Before_LineCount -and $W32TM_Status_Before[$i] -ne "") { $W32TM_Status_Before[$i] } else { "" }
        "After" = if ($i -lt $W32TM_Status_After_LineCount -and $W32TM_Status_After[$i] -ne "") { $W32TM_Status_After[$i] } else { "" }
    }
}

$W32TM_Status_BeforeAndAfter | Format-Table -Wrap


####################################################################################################
# Finishing touches
Write-Host "`n✅ NTP Configuration Complete!`n" -ForegroundColor Green
Write-Host "Log file: $ScriptLogFile" -ForegroundColor Gray

if($Unattended){
    Write-Host "`nUnattended mode: Exiting script. Bye! 👋" -ForegroundColor Gray
}else{
    Write-Host "`nPlease read the output above before exiting." -ForegroundColor Gray
    Write-Host "Waiting for 3 seconds to avoid accidental exiting." -ForegroundColor Gray
    Start-Sleep -Seconds 3
    Write-Host "Press any key to exit. Bye! 👋" -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
}

Write-Host ""
Stop-Transcript | Out-Null
