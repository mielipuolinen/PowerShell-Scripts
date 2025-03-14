# https://learn.microsoft.com/en-us/windows-server/networking/windows-time-service/windows-time-service-tools-and-settings
# https://learn.microsoft.com/en-us/troubleshoot/windows-server/active-directory/configure-w32ime-against-huge-time-offset
# https://engineering.fb.com/2020/03/18/production-engineering/ntp-service/
# https://serverfault.com/questions/334682/ws2008-ntp-using-time-windows-com-0x9-time-always-skewed-forwards
# https://developers.google.com/time/smear

# Ensure PowerShell is running as Administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "`n⚠️ Please run this script as Administrator." -ForegroundColor Red
    exit
}

Write-Host "`n🕒 Configuring NTP Sync" -ForegroundColor Green

# Set Registry Paths
$w32TimeConfig = "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Config"
$w32TimeParameters = "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\Parameters"
$w32TimeProviders = "HKLM:\SYSTEM\CurrentControlSet\Services\W32Time\TimeProviders\NtpClient"

# NTP Servers
$ntpServers = "time1.facebook.com,0x8 time2.facebook.com,0x8 time3.facebook.com,0x8 time4.facebook.com,0x8 time5.facebook.com,0x8"
#$ntpServers = "time1.google.com,0x8 time2.google.com,0x8 time3.google.com,0x8 time4.google.com,0x8"
#$ntpServers = "0.pool.ntp.org,0x8 1.pool.ntp.org,0x8 2.pool.ntp.org,0x8 3.pool.ntp.org,0x8"

# Starting Windows Time Service
Write-Host "`n🔄 Starting Windows Time Service" -ForegroundColor Green
net start w32time
Start-Sleep -Seconds 5

# Capture current configuration and status
Write-Host "`n🔍 Capturing current NTP configuration and status" -ForegroundColor Green
$beforeConfig = w32tm /query /configuration
$beforeStatus = w32tm /query /status

# Unregisters the time service, and removes all configuration information from the registry.
Write-Host "`n🔧 Unregistering the time service and removing existing time configuration" -ForegroundColor Green
W32tm /unregister

# Restart Windows Time Service
Write-Host "`n🔄 Stopping Windows Time Service" -ForegroundColor Green
net stop w32time

# Registers the time service to run as a service, and adds default configuration to the registry.
Write-Host "`n🔧 Registering the time service and adding default time configuration" -ForegroundColor Green
W32tm /register

# Allow time correction up to +-25 hours
#   Daylight saving time bugs can cause 1-hour time differences.
#   AM or PM misconfiguration can cause a 12-hour time difference.
#   Day or date mistakes can cause a 24-hour time difference.
# MaxPosPhaseCorrection: Specifies the maximum positive time correction in seconds that the service can make.
# MaxNegPhaseCorrection: Specifies the maximum negative time correction in seconds that the service can make.
# Setting both to 90000 seconds (25 hours) allows a maximum correction of 25 hours.
Write-Host "`n🔧 Configuring maximum time correction limits" -ForegroundColor Green
Set-ItemProperty -Path $w32TimeConfig -Name MaxPosPhaseCorrection -Value 90000 -Type DWord
Set-ItemProperty -Path $w32TimeConfig -Name MaxNegPhaseCorrection -Value 90000 -Type DWord

# Change MinPollInterval to ~8,5 minutes (512 seconds, 2^9)
# Default value is 10 (1024 seconds, ~17 minutes)
Write-Host "`n🔧 Configuring NTP minimum polling interval" -ForegroundColor Green
Set-ItemProperty -Path $w32TimeConfig -Name MinPollInterval -Value 9 -Type DWord

# Change MaxPollInterval to ~2,3 hours (8192 seconds, 2^13)
# Default value is 15 (32768 seconds, ~9.1 hours)
Write-Host "`n🔧 Configuring NTP maximum polling interval" -ForegroundColor Green
Set-ItemProperty -Path $w32TimeConfig -Name MaxPollInterval -Value 13 -Type DWord

# Set Client Type to NTP
# Default client NT5DS, the client synchronizes time with a domain controller in the domain hierarchy
Write-Host "`n🔧 Configuring NTP client type" -ForegroundColor Green
Set-ItemProperty -Path $w32TimeParameters -Name Type -Value "NTP" -Type String

# Enable NtpClient
Write-Host "`n🔄 Enabling NTP client" -ForegroundColor Green
Set-ItemProperty -Path $w32TimeProviders -Name Enabled -Value 1 -Type DWord

# Restart Windows Time Service
Write-Host "`n🔄 Starting Windows Time Service" -ForegroundColor Green
net start w32time

# Configure NTP servers
# /config: Modifies the configuration of the Windows Time service.
# /update: Notifies the Windows Time service that the configuration changed, causing the changes to take effect.
# /manualpeerlist:<peers>: Specifies the list of peers from which the Windows Time service obtains time stamps.
# /syncfromflags:MANUAL: Specifies that the Windows Time service is to use the manual peer list when synchronizing time.
# /reliable:NO: Set whether this computer is a reliable time source. This setting is only meaningful on DCs.
Write-Host "`n🔧 Configuring NTP servers" -ForegroundColor Green
w32tm /config /manualpeerlist:`"$ntpServers`" /syncfromflags:MANUAL /reliable:NO /update

# Restart Windows Time Service
Write-Host "`n🔄 Restarting Windows Time Service" -ForegroundColor Green
net stop w32time
net start w32time

# Resynchronize the computer clock and rediscover network sources
Write-Host "`n🔄 Resynchronizing clock" -ForegroundColor Green
# /resync: Synchronizes the computer clock with the time source, and then checks the time source for accuracy.
# /rediscover: Redetects the network configuration and rediscovers network sources, then resynchronizes.
#   Redetect Network Configuration: The Windows Time service will check the current network settings and configurations.
#   Rediscover Network Sources: It will search for available NTP servers or other time sources based on the updated network configuration.
#   Resynchronize: The service will then synchronize the computer clock with the newly discovered time sources.
#   This is particularly useful if there have been changes in the network environment, such as new NTP servers being added, changes in network topology, or updates to DNS settings.
w32tm /resync /rediscover
Start-Sleep -Seconds 5

# Capture new configuration and status
Write-Host "`n🔍 Capturing updated NTP configuration and status" -ForegroundColor Green
$afterConfig = w32tm /query /configuration
$afterStatus = w32tm /query /status

# Compare Before & After
Write-Host "`n🔍 Comparing NTP configuration and status" -ForegroundColor Green
Write-Host "Gray: No Change, Red: Before, Green: After`n" -ForegroundColor Gray

$diffConfig = Compare-Object -ReferenceObject ($beforeConfig -split "`n") -DifferenceObject ($afterConfig -split "`n") -IncludeEqual
$diffStatus = Compare-Object -ReferenceObject ($beforeStatus -split "`n") -DifferenceObject ($afterStatus -split "`n") -IncludeEqual

Write-Host "`n`tw32tm /query /configuration`n" -ForegroundColor Cyan
$diffConfig | ForEach-Object {
    if ($_.SideIndicator -eq "==") {
        Write-Host "`t`t$($_.InputObject)" -ForegroundColor Gray
    } elseif ($_.SideIndicator -eq "<=") {
        Write-Host "`t`t$($_.InputObject)" -ForegroundColor Red
    } elseif ($_.SideIndicator -eq "=>") {
        Write-Host "`t`t$($_.InputObject)" -ForegroundColor Green
    }
}

Write-Host "`n`tw32tm /query /status`n" -ForegroundColor Cyan
$diffStatus | ForEach-Object {
    if ($_.SideIndicator -eq "==") {
        Write-Host "`t`t$($_.InputObject)" -ForegroundColor Gray
    } elseif ($_.SideIndicator -eq "<=") {
        Write-Host "`t`t$($_.InputObject)" -ForegroundColor Red
    } elseif ($_.SideIndicator -eq "=>") {
        Write-Host "`t`t$($_.InputObject)" -ForegroundColor Green
    }
}

Write-Host "`n✅ NTP Configuration Complete!" -ForegroundColor Green
