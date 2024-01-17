<#
.SYNOPSIS
    This PowerShell script checks for the installed version of Microsoft 365 (M365) Apps, and based on that information, it will update to the latest version.
.DESCRIPTION

.NOTES
    File Name      : Update Microsoft 365 Apps
    Authors        : Joakim Aandal/Shehan Silva
    Version History:
        - Version 1.0 01.08.2024: Starts the Service and then Update Microsoft 365 (M365) Apps
        - Version 2.0 01.16.2024: Verify the installed Microsoft 365 (M365) Apps using Office releases, and based on that information, it will update to the latest version
#>

param (
    [String]$OfficeC2RClientPath = "$env:SystemDrive\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe",
    [String]$logDir = "$env:windir\XXX\XXX\XXX\XXX"
)

try {

# Configure powershell logging
$SaveVerbosePreference = $VerbosePreference
$VerbosePreference = 'continue'
$VMTime = Get-Date
$LogTime = $VMTime.ToUniversalTime()
mkdir $logDir -Force
Start-Transcript -Path "$logDir\m365_log.txt" -Append -IncludeInvocationHeader
Write-Host "################# New Script Run #################"
Write-host "Current time (UTC-0): $LogTime"

If ((Get-Service -Name "ClickToRunSvc").StartupType -ne "Automatic") {
    Set-Service -Name ClickToRunSvc -StartupType Automatic
    Write-Verbose "Setting Service to Automatic" -Verbose
}
Else {
    Write-Verbose "Service already set to Automatic" -Verbose
}

If ((Get-Service -Name "ClickToRunSvc").Status -eq 'Running') {
    Write-Verbose "Microsoft Office Click-to-Run Service is running" -Verbose
}
Else {

    Start-Service -Name "ClickToRunSvc" -Verbose
    Write-Verbose "Starting Microsoft Office Click-to-Run Service" -Verbose
}


#Updating Office to latest and version checking 
$InstalledVersion = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "VersionToReport"
$Channel = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "CDNBaseUrl" | Select-Object -Last 1
$CloudVersionInfo = Invoke-RestMethod 'https://clients.config.office.net/releases/v1.0/OfficeReleases'
$UsedChannel = $cloudVersioninfo | Where-Object { $_.OfficeVersions.cdnBaseURL -eq $channel }

if ($UsedChannel.latestVersion -eq $InstalledVersion) {
    Write-Verbose "Currently using the latest version of Microsoft 365 (M365) Apps in the $($UsedChannel.ChannelId) Channel | Version: $($InstalledVersion)" -Verbose
    exit 0
}
Elseif ($UsedChannel.latestVersion -ne $InstalledVersion) {
    Write-Verbose " Currently using the older version of in the $($UsedChannel.ChannelId) Channel | Version: $($InstalledVersion) and Microsoft 365 (M365) Apps are being updated to the latest version" -Verbose
    $updateprocess = Start-process $OfficeC2RClientPath -ArgumentList "/update user", displaylevel=false, forceappshutdown=true -Wait -PassThru
}

Start-Sleep 10

$processname = Get-process -Name "OfficeClickToRun"

While ($processname[0].HasExited -eq $false -and $processname[1].HasExited -eq $false ) {
        Write-Verbose " Microsoft 365 (M365) Apps are still being updated" -Verbose
        Start-Sleep 20
        $processname = Get-process -Name "OfficeClickToRun" 
}

Start-Sleep 2

#Get the Updated Version
$InstalledlatestVersion = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "VersionToReport"
Write-Verbose " Microsoft 365 (M365) Apps have been updated to $($UsedChannel.ChannelId) Channel | Version: $($InstalledlatestVersion)" -Verbose 

# End Logging
Stop-Transcript
$VerbosePreference=$SaveVerbosePreference

} catch {   
Write-Verbose "An error occurred:"
Write-Verbose $_
}