#requires -Version 5.1
<#
.SYNOPSIS
    Standalone PowerShell GUI system information console.

.DESCRIPTION
    TEKSysInfo presents a single-pane Windows Forms console for live CPU,
    memory, and disk metrics; system identity; installed software; hardware
    inventory; and Windows Update history/actions.

    Run:
        powershell.exe -ExecutionPolicy Bypass -File .\TEKSysInfo.ps1

    Notes:
      - Reading inventory does not require elevation.
      - Installing Windows updates may require Administrator rights and can
        require a reboot. The tool prompts before installation.
      - The Windows Update scan/install path uses the built-in COM API and does
        not install external PowerShell modules.
#>

param(
    [switch]$SelfTest,
    [switch]$SelfTestDiskDetails,
    [string]$SelfTestDrive
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

$Script:AppName = 'TEKSysInfo'
$Script:ProgramDataPath = Join-Path $env:ProgramData 'TEKSysInfo'
$Script:LogPath = Join-Path $Script:ProgramDataPath 'TEKSysInfo.log'
$Script:CpuHistory = New-Object System.Collections.ArrayList
$Script:MemoryHistory = New-Object System.Collections.ArrayList
$Script:DiskHistory = New-Object System.Collections.ArrayList
$Script:NetworkHistory = New-Object System.Collections.ArrayList
$Script:DiskIoHistory = New-Object System.Collections.ArrayList
$Script:TemperatureHistory = New-Object System.Collections.ArrayList
$Script:AvailableUpdates = @()
$Script:LastSnapshot = $null
$Script:LiveTimer = $null
$Script:ExtendedLiveTimer = $null
$Script:StartupInventoryTimer = $null
$Script:LiveMetricsInProgress = $false
$Script:ExtendedMetricsInProgress = $false
$Script:InventoryRefreshInProgress = $false
$Script:InventoryPendingTasks = 0
$Script:IsClosing = $false
$Script:Controls = @{}
$Script:BackgroundTaskId = 0

function Write-Log {
    param([string]$Message)

    try {
        if (-not (Test-Path -LiteralPath $Script:ProgramDataPath)) {
            New-Item -Path $Script:ProgramDataPath -ItemType Directory -Force | Out-Null
        }
        Add-Content -Path $Script:LogPath -Value ('{0} {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message) -Encoding UTF8
    } catch {
        # Logging is diagnostic only and must not block the interface.
    }
}

function Enable-ControlDoubleBuffering {
    param([System.Windows.Forms.Control]$Control)

    if (-not $Control) { return }
    try {
        $property = $Control.GetType().GetProperty('DoubleBuffered', [Reflection.BindingFlags]'Instance,NonPublic')
        if ($property) {
            $property.SetValue($Control, $true, $null)
        }
    } catch {
        # Some controls refuse reflection against DoubleBuffered.
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function ConvertTo-DisplayText {
    param([object]$Value, [string]$Fallback = 'Unknown')

    if ($null -eq $Value) { return $Fallback }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $Fallback }
    return $text
}

function ConvertTo-SizeText {
    param([Nullable[double]]$Bytes)

    if ($null -eq $Bytes -or $Bytes -lt 0) { return 'Unknown' }
    $units = @('B','KB','MB','GB','TB','PB')
    $value = [double]$Bytes
    $index = 0
    while ($value -ge 1024 -and $index -lt ($units.Count - 1)) {
        $value = $value / 1024
        $index++
    }
    return ('{0:N1} {1}' -f $value, $units[$index])
}

function ConvertTo-RateText {
    param([Nullable[double]]$BytesPerSecond)

    if ($null -eq $BytesPerSecond -or $BytesPerSecond -lt 0) { return 'Unknown' }
    return ('{0}/s' -f (ConvertTo-SizeText ([double]$BytesPerSecond)))
}

function ConvertTo-PercentText {
    param([Nullable[double]]$Value)

    if ($null -eq $Value -or [double]::IsNaN([double]$Value)) { return 'Unknown' }
    return ('{0:N0}%' -f [Math]::Max(0, [Math]::Min(100, [double]$Value)))
}

function ConvertFrom-CimDate {
    param([object]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [DateTime]) { return [DateTime]$Value }
    try { return [Management.ManagementDateTimeConverter]::ToDateTime([string]$Value) } catch { return $null }
}

function Get-ObjectPropertyValue {
    param(
        [object]$InputObject,
        [string]$Name,
        [object]$Default = $null
    )

    if ($null -eq $InputObject) { return $Default }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $Default }
    if ($null -eq $property.Value) { return $Default }
    return $property.Value
}

function Get-CimSafe {
    param(
        [Parameter(Mandatory)][string]$ClassName,
        [string]$Filter,
        [string]$Namespace = 'root/cimv2'
    )

    try {
        $params = @{ ClassName = $ClassName; Namespace = $Namespace; ErrorAction = 'Stop' }
        if ($Filter) { $params.Filter = $Filter }
        return Get-CimInstance @params
    } catch {
        Write-Log "CIM query failed for ${ClassName}: $($_.Exception.Message)"
        return @()
    }
}

function Add-HistoryValue {
    param(
        [System.Collections.ArrayList]$History,
        [double]$Value,
        [int]$MaxCount = 90
    )

    [void]$History.Add([Math]::Max(0, [Math]::Min(100, $Value)))
    while ($History.Count -gt $MaxCount) {
        $History.RemoveAt(0)
    }
}

function Get-PrimaryDiskUsage {
    $systemDrive = if ($env:SystemDrive) { $env:SystemDrive.TrimEnd('\') } else { 'C:' }
    $logical = Get-CimSafe -ClassName Win32_LogicalDisk -Filter "DeviceID='$systemDrive'"
    if (-not $logical) { $logical = Get-CimSafe -ClassName Win32_LogicalDisk -Filter 'DriveType=3' | Select-Object -First 1 }
    if (-not $logical -or -not $logical.Size) { return $null }
    return [pscustomobject]@{
        DeviceId = $logical.DeviceID
        UsedPercent = [Math]::Round((($logical.Size - $logical.FreeSpace) / $logical.Size) * 100, 1)
        FreePercent = [Math]::Round(($logical.FreeSpace / $logical.Size) * 100, 1)
        Size = [double]$logical.Size
        Free = [double]$logical.FreeSpace
    }
}

function ConvertFrom-KelvinTenths {
    param([Nullable[double]]$Value)

    if ($null -eq $Value -or $Value -le 0) { return $null }
    return [Math]::Round(([double]$Value / 10) - 273.15, 1)
}

function ConvertFrom-Kelvin {
    param([Nullable[double]]$Value)

    if ($null -eq $Value -or $Value -le 0) { return $null }
    return [Math]::Round([double]$Value - 273.15, 1)
}

function Get-ThermalReadings {
    $readings = New-Object System.Collections.Generic.List[object]

    foreach ($zone in (Get-CimSafe -Namespace 'root/cimv2' -ClassName 'Win32_PerfFormattedData_Counters_ThermalZoneInformation')) {
        $tempC = $null
        $raw = Get-ObjectPropertyValue -InputObject $zone -Name 'HighPrecisionTemperature'
        if ($raw) { $tempC = ConvertFrom-KelvinTenths ([double]$raw) }
        if ($null -eq $tempC) {
            $raw = Get-ObjectPropertyValue -InputObject $zone -Name 'Temperature'
            if ($raw) { $tempC = ConvertFrom-Kelvin ([double]$raw) }
        }
        if ($null -ne $tempC) {
            $readings.Add([pscustomobject]@{
                Name = ConvertTo-DisplayText $zone.Name 'Thermal Zone'
                TemperatureC = $tempC
                Source = 'Windows thermal zone'
                Detail = 'ACPI/platform thermal zone; may not be CPU package temperature.'
            })
        }
    }

    foreach ($zone in (Get-CimSafe -Namespace 'root/wmi' -ClassName 'MSAcpi_ThermalZoneTemperature')) {
        $tempC = ConvertFrom-KelvinTenths ([double]$zone.CurrentTemperature)
        if ($null -ne $tempC) {
            $readings.Add([pscustomobject]@{
                Name = ConvertTo-DisplayText $zone.InstanceName 'ACPI Thermal Zone'
                TemperatureC = $tempC
                Source = 'MSAcpi thermal zone'
                Detail = 'ACPI/platform thermal zone; may not be CPU package temperature.'
            })
        }
    }

    return @($readings | Sort-Object Name -Unique)
}

function Get-NetworkMetrics {
    $adapters = @(Get-CimSafe -ClassName Win32_PerfFormattedData_Tcpip_NetworkInterface |
        Where-Object {
            $_.Name -and
            $_.Name -notmatch 'Loopback|isatap|Teredo|Bluetooth Device|Pseudo-Interface'
        })

    $bytesIn = 0.0
    $bytesOut = 0.0
    foreach ($adapter in $adapters) {
        $bytesIn += [double](Get-ObjectPropertyValue -InputObject $adapter -Name 'BytesReceivedPersec' -Default 0)
        $bytesOut += [double](Get-ObjectPropertyValue -InputObject $adapter -Name 'BytesSentPersec' -Default 0)
    }

    $activeAdapter = $adapters | Sort-Object { [double]$_.BytesTotalPersec } -Descending | Select-Object -First 1
    $adapterName = if ($activeAdapter) { ConvertTo-DisplayText $activeAdapter.Name } else { 'No active adapter' }

    return [pscustomobject]@{
        BytesInPerSec = $bytesIn
        BytesOutPerSec = $bytesOut
        BytesTotalPerSec = $bytesIn + $bytesOut
        ActiveAdapter = $adapterName
    }
}

function Get-DiskActivityMetrics {
    $diskSample = Get-CimSafe -ClassName Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'" | Select-Object -First 1
    if (-not $diskSample) {
        return [pscustomobject]@{
            ReadBytesPerSec = 0.0
            WriteBytesPerSec = 0.0
            TotalBytesPerSec = 0.0
            QueueLength = 0.0
        }
    }

    $read = [double](Get-ObjectPropertyValue -InputObject $diskSample -Name 'DiskReadBytesPersec' -Default 0)
    $write = [double](Get-ObjectPropertyValue -InputObject $diskSample -Name 'DiskWriteBytesPersec' -Default 0)
    return [pscustomobject]@{
        ReadBytesPerSec = $read
        WriteBytesPerSec = $write
        TotalBytesPerSec = $read + $write
        QueueLength = [double](Get-ObjectPropertyValue -InputObject $diskSample -Name 'CurrentDiskQueueLength' -Default 0)
    }
}

function Get-TopProcessMetrics {
    $cpuProcess = Get-CimSafe -ClassName Win32_PerfFormattedData_PerfProc_Process |
        Where-Object { $_.Name -and $_.Name -notin @('_Total','Idle') } |
        Sort-Object PercentProcessorTime -Descending |
        Select-Object -First 1

    $memoryProcess = $null
    try {
        $memoryProcess = Get-Process -ErrorAction Stop | Sort-Object WorkingSet64 -Descending | Select-Object -First 1
    } catch { }

    return [pscustomobject]@{
        TopCpuName = if ($cpuProcess) { ConvertTo-DisplayText $cpuProcess.Name } else { 'Unknown' }
        TopCpuPercent = if ($cpuProcess) { [Math]::Round([double]$cpuProcess.PercentProcessorTime, 1) } else { 0 }
        TopMemoryName = if ($memoryProcess) { ConvertTo-DisplayText $memoryProcess.ProcessName } else { 'Unknown' }
        TopMemoryBytes = if ($memoryProcess) { [double](Get-ObjectPropertyValue -InputObject $memoryProcess -Name 'WorkingSet64' -Default 0) } else { 0 }
    }
}

function Get-LiveMetrics {
    param([switch]$Fast)

    $cpu = $null
    try {
        $cpuSample = Get-CimSafe -ClassName Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'" | Select-Object -First 1
        if ($cpuSample) { $cpu = [double]$cpuSample.PercentProcessorTime }
    } catch {
        $cpu = $null
    }

    $memoryPercent = $null
    $memoryUsed = $null
    $memoryTotal = $null
    $memoryFree = $null
    try {
        $computerInfo = New-Object Microsoft.VisualBasic.Devices.ComputerInfo
        $memoryTotal = [double]$computerInfo.TotalPhysicalMemory
        $memoryFree = [double]$computerInfo.AvailablePhysicalMemory
        if ($memoryTotal -gt 0) {
            $memoryUsed = $memoryTotal - $memoryFree
            $memoryPercent = [Math]::Round(($memoryUsed / $memoryTotal) * 100, 1)
        }
    } catch {
        Write-Log "Fast memory query failed: $($_.Exception.Message)"
    }

    $disk = Get-PrimaryDiskUsage
    $thermalReadings = @(Get-ThermalReadings)
    $highestTemperature = if ($thermalReadings.Count -gt 0) {
        ($thermalReadings | Sort-Object TemperatureC -Descending | Select-Object -First 1).TemperatureC
    } else {
        $null
    }

    return [pscustomobject]@{
        CpuPercent = if ($null -eq $cpu) { 0 } else { [Math]::Round($cpu, 1) }
        MemoryPercent = if ($null -eq $memoryPercent) { 0 } else { $memoryPercent }
        MemoryUsed = $memoryUsed
        MemoryTotal = $memoryTotal
        MemoryFree = $memoryFree
        DiskPercent = if ($disk) { $disk.UsedPercent } else { 0 }
        Disk = $disk
        ThermalReadings = $thermalReadings
        HighestTemperatureC = $highestTemperature
    }
}

function Get-ExtendedDashboardMetrics {
    return [pscustomobject]@{
        Network = Get-NetworkMetrics
        DiskActivity = Get-DiskActivityMetrics
        TopProcesses = Get-TopProcessMetrics
    }
}

function Get-SystemSnapshot {
    $computer = Get-CimSafe -ClassName Win32_ComputerSystem | Select-Object -First 1
    $os = Get-CimSafe -ClassName Win32_OperatingSystem | Select-Object -First 1
    $bios = Get-CimSafe -ClassName Win32_BIOS | Select-Object -First 1
    $baseBoard = Get-CimSafe -ClassName Win32_BaseBoard | Select-Object -First 1
    $processor = Get-CimSafe -ClassName Win32_Processor | Select-Object -First 1
    $enclosure = Get-CimSafe -ClassName Win32_SystemEnclosure | Select-Object -First 1
    $secureBoot = 'Unknown'

    try {
        $secureBoot = if (Confirm-SecureBootUEFI -ErrorAction Stop) { 'Enabled' } else { 'Disabled' }
    } catch {
        $secureBoot = 'Unavailable'
    }

    $installDate = $null
    $lastBoot = $null
    if ($os) {
        $installDate = ConvertFrom-CimDate $os.InstallDate
        $lastBoot = ConvertFrom-CimDate $os.LastBootUpTime
    }

    return [pscustomobject]@{
        ComputerName = ConvertTo-DisplayText $env:COMPUTERNAME
        UserName = ConvertTo-DisplayText ([Security.Principal.WindowsIdentity]::GetCurrent().Name)
        Manufacturer = ConvertTo-DisplayText $computer.Manufacturer
        Model = ConvertTo-DisplayText $computer.Model
        SerialNumber = ConvertTo-DisplayText $bios.SerialNumber
        ChassisSerial = ConvertTo-DisplayText $enclosure.SerialNumber
        BiosVersion = ConvertTo-DisplayText (($bios.SMBIOSBIOSVersion, $bios.Version | Where-Object { $_ }) -join ' / ')
        BiosDate = if ($bios.ReleaseDate) { (ConvertFrom-CimDate $bios.ReleaseDate).ToString('yyyy-MM-dd') } else { 'Unknown' }
        BaseBoard = ConvertTo-DisplayText ('{0} {1}' -f $baseBoard.Manufacturer, $baseBoard.Product)
        Processor = ConvertTo-DisplayText $processor.Name
        ProcessorCores = ConvertTo-DisplayText $processor.NumberOfCores
        ProcessorThreads = ConvertTo-DisplayText $processor.NumberOfLogicalProcessors
        TotalMemory = if ($computer.TotalPhysicalMemory) { ConvertTo-SizeText ([double]$computer.TotalPhysicalMemory) } else { 'Unknown' }
        WindowsCaption = ConvertTo-DisplayText $os.Caption
        WindowsVersion = ConvertTo-DisplayText ('{0} build {1}' -f $os.Version, $os.BuildNumber)
        WindowsArchitecture = ConvertTo-DisplayText $os.OSArchitecture
        InstallDate = if ($installDate) { $installDate.ToString('yyyy-MM-dd HH:mm') } else { 'Unknown' }
        LastBoot = if ($lastBoot) { $lastBoot.ToString('yyyy-MM-dd HH:mm') } else { 'Unknown' }
        Uptime = if ($lastBoot) { New-TimeSpan -Start $lastBoot -End (Get-Date) } else { $null }
        SecureBoot = $secureBoot
        IsAdministrator = Test-IsAdministrator
    }
}

function Test-PendingReboot {
    $reasons = New-Object System.Collections.Generic.List[string]
    $checks = @(
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'; Reason = 'Component servicing' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'; Reason = 'Windows Update' },
        @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'; Name = 'PendingFileRenameOperations'; Reason = 'Pending file rename' },
        @{ Path = 'HKLM:\SOFTWARE\Microsoft\Updates\UpdateExeVolatile'; Reason = 'Update executable' }
    )

    foreach ($check in $checks) {
        try {
            if ($check.ContainsKey('Name')) {
                $value = Get-ItemProperty -Path $check.Path -Name $check.Name -ErrorAction SilentlyContinue
                if ($value) { $reasons.Add($check.Reason) }
            } elseif (Test-Path -LiteralPath $check.Path) {
                $reasons.Add($check.Reason)
            }
        } catch { }
    }

    return [pscustomobject]@{
        IsPending = ($reasons.Count -gt 0)
        Reasons = @($reasons)
        Text = if ($reasons.Count -gt 0) { 'Pending: {0}' -f ($reasons -join ', ') } else { 'No reboot pending' }
    }
}

function Get-PowerSummary {
    $battery = Get-CimSafe -ClassName Win32_Battery | Select-Object -First 1
    $plan = ''
    try {
        $activePlan = powercfg /getactivescheme 2>$null
        $plan = [string]$activePlan
    } catch { }
    if ($battery) {
        return [pscustomobject]@{
            Status = if ([int]$battery.BatteryStatus -eq 2) { 'AC power' } else { 'Battery' }
            Detail = ('{0}% charge | {1}' -f $battery.EstimatedChargeRemaining, (ConvertTo-DisplayText $plan 'Power plan unknown'))
        }
    }

    return [pscustomobject]@{
        Status = 'Desktop / AC power'
        Detail = ConvertTo-DisplayText $plan 'Battery not detected'
    }
}

function Get-SecuritySummary {
    $secureBoot = 'Unknown'
    try { $secureBoot = if (Confirm-SecureBootUEFI -ErrorAction Stop) { 'On' } else { 'Off' } } catch { $secureBoot = 'Unavailable' }

    $firewallText = 'Unknown'
    try {
        $profiles = @(Get-NetFirewallProfile -ErrorAction Stop)
        $enabled = @($profiles | Where-Object { $_.Enabled }).Count
        $firewallText = ('Firewall {0}/{1} profiles on' -f $enabled, $profiles.Count)
    } catch { }

    $defenderText = 'Unknown'
    try {
        $defender = Get-MpComputerStatus -ErrorAction Stop
        $defenderText = if ($defender.AntivirusEnabled) { 'Defender on' } else { 'Defender off' }
    } catch { }

    return [pscustomobject]@{
        Status = ('Secure Boot {0}' -f $secureBoot)
        Detail = ('{0} | {1}' -f $defenderText, $firewallText)
    }
}

function Get-DriveHealthSummary {
    $physicalDisks = @()
    try { $physicalDisks = @(Get-PhysicalDisk -ErrorAction Stop) } catch { }
    $unhealthy = @($physicalDisks | Where-Object { [string]$_.HealthStatus -notin @('Healthy','') })
    $hottest = ''
    $wearText = ''
    try {
        $counters = foreach ($physical in $physicalDisks) {
            try { $physical | Get-StorageReliabilityCounter -ErrorAction Stop } catch { }
        }
        $hotCounter = @($counters | Where-Object { $null -ne $_.Temperature } | Sort-Object Temperature -Descending | Select-Object -First 1)
        if ($hotCounter) { $hottest = ('Hottest disk {0} C' -f $hotCounter.Temperature) }
        $wearCounter = @($counters | Where-Object { $null -ne $_.Wear -and [string]$_.Wear -ne '' } | Sort-Object Wear -Descending | Select-Object -First 1)
        if ($wearCounter) { $wearText = ('Worst Windows wear counter {0}%' -f $wearCounter.Wear) }
    } catch { }

    return [pscustomobject]@{
        Status = if ($unhealthy.Count -eq 0) { 'All drives healthy' } else { ('{0} drive warning(s)' -f $unhealthy.Count) }
        Detail = (@($hottest, $wearText) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join ' | '
    }
}

function Get-LogicalDisks {
    Get-CimSafe -ClassName Win32_LogicalDisk -Filter 'DriveType=3' |
        Sort-Object DeviceID |
        ForEach-Object {
            $used = if ($_.Size) { [double]$_.Size - [double]$_.FreeSpace } else { $null }
            $usedPercent = if ($_.Size) { [Math]::Round(($used / [double]$_.Size) * 100, 1) } else { $null }
            [pscustomobject]@{
                Drive = $_.DeviceID
                Label = ConvertTo-DisplayText $_.VolumeName ''
                FileSystem = ConvertTo-DisplayText $_.FileSystem
                Size = ConvertTo-SizeText ([double]$_.Size)
                Used = ConvertTo-SizeText $used
                Free = ConvertTo-SizeText ([double]$_.FreeSpace)
                UsedPercent = $usedPercent
            }
        }
}

function Format-NullableValue {
    param(
        [object]$Value,
        [string]$Suffix = ''
    )

    if ($null -eq $Value) { return 'Not reported' }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return 'Not reported' }
    if (-not [string]::IsNullOrWhiteSpace($Suffix)) { return "$text$Suffix" }
    return $text
}

function Format-HealthHours {
    param([object]$Hours)

    if ($null -eq $Hours) { return 'Not reported' }
    try {
        $value = [double]$Hours
        if ($value -lt 0) { return 'Not reported' }
        $days = $value / 24
        if ($days -ge 1) { return ('{0:N0} hr ({1:N1} days)' -f $value, $days) }
        return ('{0:N0} hr' -f $value)
    } catch {
        return 'Not reported'
    }
}

function Get-NumericPropertyValue {
    param(
        [object]$InputObject,
        [string]$Name
    )

    $value = Get-ObjectPropertyValue -InputObject $InputObject -Name $Name -Default $null
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return $null }
    try { return [double]$value } catch { return $null }
}

function Get-NormalizedHardwareId {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    return ([regex]::Replace($Value.ToUpperInvariant(), '[^A-Z0-9]', ''))
}

function Find-SmartctlPath {
    $candidates = New-Object System.Collections.Generic.List[string]
    $candidates.Add('smartctl.exe')
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $candidates.Add((Join-Path $PSScriptRoot 'smartctl.exe'))
        $candidates.Add((Join-Path $PSScriptRoot 'smartmontools\bin\smartctl.exe'))
    }
    $candidates.Add('C:\Program Files\smartmontools\bin\smartctl.exe')
    $candidates.Add('C:\Program Files (x86)\smartmontools\bin\smartctl.exe')

    foreach ($candidate in $candidates) {
        try {
            $command = Get-Command $candidate -ErrorAction SilentlyContinue
            if ($command) { return $command.Source }
        } catch {
            # Keep searching other common locations.
        }
    }

    return ''
}

function Format-NvmeDataUnits {
    param([object]$DataUnits)

    if ($null -eq $DataUnits -or [string]::IsNullOrWhiteSpace([string]$DataUnits)) { return 'Not reported' }
    try {
        $units = [double]$DataUnits
        if ($units -lt 0) { return 'Not reported' }
        return ConvertTo-SizeText ([double]($units * 512000))
    } catch {
        return 'Not reported'
    }
}

function Get-SmartctlReports {
    $smartctlPath = Find-SmartctlPath
    if ([string]::IsNullOrWhiteSpace($smartctlPath)) {
        Write-Log 'smartctl.exe was not found. Disk detail will use Windows storage counters only.'
        return @()
    }

    $devices = @()
    try {
        $scanJson = & $smartctlPath -j --scan-open 2>$null
        if ($LASTEXITCODE -gt 1 -or [string]::IsNullOrWhiteSpace(($scanJson -join [Environment]::NewLine))) {
            Write-Log "smartctl scan failed with exit code $LASTEXITCODE."
            return @()
        }
        $scan = ($scanJson -join [Environment]::NewLine) | ConvertFrom-Json
        $devices = @($scan.devices)
    } catch {
        Write-Log "smartctl scan failed: $($_.Exception.Message)"
        return @()
    }

    $reports = New-Object System.Collections.Generic.List[object]
    foreach ($device in $devices) {
        $name = [string](Get-ObjectPropertyValue -InputObject $device -Name 'name' -Default '')
        $type = [string](Get-ObjectPropertyValue -InputObject $device -Name 'type' -Default '')
        if ([string]::IsNullOrWhiteSpace($name)) { continue }

        try {
            $args = @('-a', '-j', $name)
            if (-not [string]::IsNullOrWhiteSpace($type)) { $args += @('-d', $type) }
            $jsonLines = & $smartctlPath @args 2>$null
            if ([string]::IsNullOrWhiteSpace(($jsonLines -join [Environment]::NewLine))) { continue }
            $report = ($jsonLines -join [Environment]::NewLine) | ConvertFrom-Json
            $report | Add-Member -NotePropertyName 'SmartctlDeviceName' -NotePropertyValue $name -Force
            $report | Add-Member -NotePropertyName 'SmartctlDeviceType' -NotePropertyValue $type -Force
            [void]$reports.Add($report)
        } catch {
            Write-Log "smartctl details failed for ${name}: $($_.Exception.Message)"
        }
    }

    return $reports.ToArray()
}

function Find-SmartctlReportForDisk {
    param(
        [object[]]$Reports,
        [object]$DiskDrive,
        [object]$PhysicalDisk
    )

    $diskSerials = @(
        [string](Get-ObjectPropertyValue -InputObject $DiskDrive -Name 'SerialNumber' -Default ''),
        [string](Get-ObjectPropertyValue -InputObject $PhysicalDisk -Name 'SerialNumber' -Default '')
    ) | ForEach-Object { Get-NormalizedHardwareId -Value $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    foreach ($report in @($Reports)) {
        $reportSerial = Get-NormalizedHardwareId -Value ([string](Get-ObjectPropertyValue -InputObject $report -Name 'serial_number' -Default ''))
        if ([string]::IsNullOrWhiteSpace($reportSerial)) { continue }
        foreach ($diskSerial in $diskSerials) {
            if ($diskSerial -eq $reportSerial -or $diskSerial.Contains($reportSerial) -or $reportSerial.Contains($diskSerial)) {
                return $report
            }
        }
    }

    $diskModel = Get-NormalizedHardwareId -Value ([string](Get-ObjectPropertyValue -InputObject $DiskDrive -Name 'Model' -Default (Get-ObjectPropertyValue -InputObject $PhysicalDisk -Name 'FriendlyName' -Default '')))
    $diskSize = [UInt64](Get-ObjectPropertyValue -InputObject $DiskDrive -Name 'Size' -Default 0)
    foreach ($report in @($Reports)) {
        $reportModel = Get-NormalizedHardwareId -Value ([string](Get-ObjectPropertyValue -InputObject $report -Name 'model_name' -Default (Get-ObjectPropertyValue -InputObject $report -Name 'model_family' -Default '')))
        $userCapacity = Get-ObjectPropertyValue -InputObject $report -Name 'user_capacity' -Default $null
        $reportSize = [UInt64](Get-ObjectPropertyValue -InputObject $userCapacity -Name 'bytes' -Default 0)
        if (-not [string]::IsNullOrWhiteSpace($diskModel) -and -not [string]::IsNullOrWhiteSpace($reportModel) -and ($diskModel.Contains($reportModel) -or $reportModel.Contains($diskModel))) {
            if ($diskSize -eq 0 -or $reportSize -eq 0 -or ([Math]::Abs([double]$diskSize - [double]$reportSize) -lt 104857600)) {
                return $report
            }
        }
    }

    return $null
}

function Get-SmartctlHealthValues {
    param([object]$SmartctlReport)

    if ($null -eq $SmartctlReport) { return $null }

    $status = Get-ObjectPropertyValue -InputObject $SmartctlReport -Name 'smart_status' -Default $null
    $passed = Get-ObjectPropertyValue -InputObject $status -Name 'passed' -Default $null
    $smartText = 'Not reported'
    if ($null -ne $passed) { $smartText = if ([bool]$passed) { 'SMART passed' } else { 'SMART FAILED' } }

    $nvmeLog = Get-ObjectPropertyValue -InputObject $SmartctlReport -Name 'nvme_smart_health_information_log' -Default $null
    if ($nvmeLog) {
        $percentageUsed = Get-NumericPropertyValue -InputObject $nvmeLog -Name 'percentage_used'
        $estimatedHealth = 'Not reported'
        if ($null -ne $percentageUsed) {
            $estimatedHealth = ('{0:N0}% remaining ({1:N0}% used)' -f ([Math]::Max(0, 100 - $percentageUsed)), $percentageUsed)
        }

        return [pscustomobject]@{
            EstimatedHealth = $estimatedHealth
            PercentageUsed = if ($null -ne $percentageUsed) { ('{0:N0}%' -f $percentageUsed) } else { 'Not reported' }
            SmartPrediction = $smartText
            Temperature = Format-NullableValue -Value (Get-ObjectPropertyValue -InputObject $nvmeLog -Name 'temperature' -Default (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $SmartctlReport -Name 'temperature' -Default $null) -Name 'current' -Default $null)) -Suffix ' C'
            PowerOnHours = Format-HealthHours -Hours (Get-ObjectPropertyValue -InputObject $nvmeLog -Name 'power_on_hours' -Default (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $SmartctlReport -Name 'power_on_time' -Default $null) -Name 'hours' -Default $null))
            PowerCycleCount = Format-NullableValue -Value (Get-ObjectPropertyValue -InputObject $nvmeLog -Name 'power_cycles' -Default $null)
            TotalBytesRead = Format-NvmeDataUnits -DataUnits (Get-ObjectPropertyValue -InputObject $nvmeLog -Name 'data_units_read' -Default $null)
            TotalBytesWritten = Format-NvmeDataUnits -DataUnits (Get-ObjectPropertyValue -InputObject $nvmeLog -Name 'data_units_written' -Default $null)
            MediaErrors = Format-NullableValue -Value (Get-ObjectPropertyValue -InputObject $nvmeLog -Name 'media_errors' -Default $null)
            ErrorLogEntries = Format-NullableValue -Value (Get-ObjectPropertyValue -InputObject $nvmeLog -Name 'num_err_log_entries' -Default $null)
            Source = "smartctl $([string](Get-ObjectPropertyValue -InputObject $SmartctlReport -Name 'SmartctlDeviceName' -Default ''))"
        }
    }

    return [pscustomobject]@{
        EstimatedHealth = 'Not reported'
        PercentageUsed = 'Not reported'
        SmartPrediction = $smartText
        Temperature = Format-NullableValue -Value (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $SmartctlReport -Name 'temperature' -Default $null) -Name 'current' -Default $null) -Suffix ' C'
        PowerOnHours = Format-HealthHours -Hours (Get-ObjectPropertyValue -InputObject (Get-ObjectPropertyValue -InputObject $SmartctlReport -Name 'power_on_time' -Default $null) -Name 'hours' -Default $null)
        PowerCycleCount = 'Not reported'
        TotalBytesRead = 'Not reported'
        TotalBytesWritten = 'Not reported'
        MediaErrors = 'Not reported'
        ErrorLogEntries = 'Not reported'
        Source = "smartctl $([string](Get-ObjectPropertyValue -InputObject $SmartctlReport -Name 'SmartctlDeviceName' -Default ''))"
    }
}

function Get-DiskDetail {
    param([string]$DriveLetter)

    $driveId = if ($DriveLetter) { $DriveLetter.TrimEnd('\') } else { $env:SystemDrive.TrimEnd('\') }
    $logical = Get-CimSafe -ClassName Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $driveId.Replace("'", "''")) | Select-Object -First 1
    $partition = $null
    $diskDrive = $null
    if ($logical) {
        try {
            $escapedDevice = $logical.DeviceID.Replace('\', '\\')
            $partition = Get-CimInstance -Query "ASSOCIATORS OF {Win32_LogicalDisk.DeviceID='$escapedDevice'} WHERE AssocClass=Win32_LogicalDiskToPartition" -ErrorAction Stop | Select-Object -First 1
            if ($partition) {
                $diskDrive = Get-CimInstance -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($partition.DeviceID)'} WHERE AssocClass=Win32_DiskDriveToDiskPartition" -ErrorAction Stop | Select-Object -First 1
            }
        } catch {
            Write-Log "Disk association lookup failed for ${driveId}: $($_.Exception.Message)"
        }
    }

    $physicalDisk = $null
    if ($diskDrive) {
        try {
            $physicalDisk = Get-PhysicalDisk -ErrorAction Stop |
                Where-Object {
                    $_.FriendlyName -eq $diskDrive.Model -or
                    $_.SerialNumber -eq $diskDrive.SerialNumber -or
                    $_.DeviceId -eq $diskDrive.Index
                } |
                Select-Object -First 1
        } catch {
            Write-Log "Get-PhysicalDisk failed: $($_.Exception.Message)"
        }
    }

    $storageReliability = $null
    if ($physicalDisk) {
        try {
            $storageReliability = $physicalDisk | Get-StorageReliabilityCounter -ErrorAction Stop
        } catch {
            Write-Log "Get-StorageReliabilityCounter failed for $($physicalDisk.FriendlyName): $($_.Exception.Message)"
        }
    }

    $smartctlValues = $null
    try {
        $smartctlReport = Find-SmartctlReportForDisk -Reports @(Get-SmartctlReports) -DiskDrive $diskDrive -PhysicalDisk $physicalDisk
        $smartctlValues = Get-SmartctlHealthValues -SmartctlReport $smartctlReport
    } catch {
        Write-Log "smartctl health lookup failed for ${driveId}: $($_.Exception.Message)"
    }

    $volume = $null
    try {
        $volume = Get-Volume -DriveLetter $driveId.TrimEnd(':') -ErrorAction Stop
    } catch {
        Write-Log "Get-Volume failed for ${driveId}: $($_.Exception.Message)"
    }

    $rows = New-Object System.Collections.Generic.List[object]
    $addRow = {
        param([string]$Category, [string]$Property, [object]$Value)
        $rows.Add([pscustomobject]@{
            Category = $Category
            Property = $Property
            Value = ConvertTo-DisplayText $Value ''
        })
    }

    if ($smartctlValues) {
        & $addRow 'Health Summary' 'Drive health remaining' $smartctlValues.EstimatedHealth
        & $addRow 'Health Summary' 'SMART status' $smartctlValues.SmartPrediction
        & $addRow 'Health Summary' 'Temperature' $smartctlValues.Temperature
        & $addRow 'Health Summary' 'Health data source' $smartctlValues.Source
    } elseif ($storageReliability) {
        & $addRow 'Health Summary' 'Drive health remaining' 'Not reported by Windows storage provider'
        & $addRow 'Health Summary' 'Windows health status' $physicalDisk.HealthStatus
        & $addRow 'Health Summary' 'Temperature' (Format-NullableValue -Value $storageReliability.Temperature -Suffix ' C')
        & $addRow 'Health Summary' 'Health data source' 'Windows storage reliability counters'
    } else {
        & $addRow 'Health Summary' 'Drive health remaining' 'Not reported'
        & $addRow 'Health Summary' 'Health data source' 'No SMART or Windows reliability counters available'
    }

    & $addRow 'Volume Usage' 'Drive' $driveId
    & $addRow 'Volume' 'Label' $logical.VolumeName
    & $addRow 'Volume' 'File system' $logical.FileSystem
    & $addRow 'Volume Usage' 'Volume size' (ConvertTo-SizeText ([double]$logical.Size))
    & $addRow 'Volume Usage' 'Volume free space' (ConvertTo-SizeText ([double]$logical.FreeSpace))
    if ($logical.Size) {
        $usedBytes = [double]$logical.Size - [double]$logical.FreeSpace
        & $addRow 'Volume Usage' 'Volume used space' (ConvertTo-SizeText $usedBytes)
        & $addRow 'Volume Usage' 'Volume space used' (ConvertTo-PercentText ($usedBytes / [double]$logical.Size * 100))
    }
    if ($volume) {
        & $addRow 'Volume' 'Volume health status' $volume.HealthStatus
        & $addRow 'Volume' 'Operational status' (($volume.OperationalStatus | ForEach-Object { $_ }) -join ', ')
        & $addRow 'Volume' 'Allocation unit size' $volume.AllocationUnitSize
    }

    if ($partition) {
        & $addRow 'Partition' 'Name' $partition.Name
        & $addRow 'Partition' 'Type' $partition.Type
        & $addRow 'Partition' 'Boot partition' $partition.BootPartition
        & $addRow 'Partition' 'Primary partition' $partition.PrimaryPartition
        & $addRow 'Partition' 'Size' (ConvertTo-SizeText ([double]$partition.Size))
        & $addRow 'Partition' 'Starting offset' $partition.StartingOffset
    }

    if ($diskDrive) {
        & $addRow 'Physical Disk' 'Model' $diskDrive.Model
        & $addRow 'Physical Disk' 'Manufacturer' $diskDrive.Manufacturer
        & $addRow 'Physical Disk' 'Serial number' $diskDrive.SerialNumber
        & $addRow 'Physical Disk' 'Firmware revision' $diskDrive.FirmwareRevision
        & $addRow 'Physical Disk' 'Interface' $diskDrive.InterfaceType
        & $addRow 'Physical Disk' 'Media type' $diskDrive.MediaType
        & $addRow 'Physical Disk' 'Status' $diskDrive.Status
        & $addRow 'Physical Disk' 'Partitions' $diskDrive.Partitions
        & $addRow 'Physical Disk' 'Bytes per sector' $diskDrive.BytesPerSector
        & $addRow 'Physical Disk' 'Total size' (ConvertTo-SizeText ([double]$diskDrive.Size))
    }

    if ($physicalDisk) {
        & $addRow 'Storage Health' 'Friendly name' $physicalDisk.FriendlyName
        & $addRow 'Storage Health' 'Health status' $physicalDisk.HealthStatus
        & $addRow 'Storage Health' 'Operational status' (($physicalDisk.OperationalStatus | ForEach-Object { $_ }) -join ', ')
        & $addRow 'Storage Health' 'Media type' $physicalDisk.MediaType
        & $addRow 'Storage Health' 'Bus type' $physicalDisk.BusType
        & $addRow 'Storage Health' 'Spindle speed' $physicalDisk.SpindleSpeed
        & $addRow 'Storage Health' 'Cannot pool reason' (($physicalDisk.CannotPoolReason | ForEach-Object { $_ }) -join ', ')
    }

    if ($smartctlValues) {
        & $addRow 'SMART Health' 'Estimated health' $smartctlValues.EstimatedHealth
        & $addRow 'SMART Health' 'SMART status' $smartctlValues.SmartPrediction
        & $addRow 'SMART Health' 'NVMe percentage used' $smartctlValues.PercentageUsed
        & $addRow 'SMART Health' 'Temperature' $smartctlValues.Temperature
        & $addRow 'SMART Health' 'Power on hours' $smartctlValues.PowerOnHours
        & $addRow 'SMART Health' 'Power cycles' $smartctlValues.PowerCycleCount
        & $addRow 'SMART Health' 'Total bytes read' $smartctlValues.TotalBytesRead
        & $addRow 'SMART Health' 'Total bytes written' $smartctlValues.TotalBytesWritten
        & $addRow 'SMART Health' 'Media errors' $smartctlValues.MediaErrors
        & $addRow 'SMART Health' 'Error log entries' $smartctlValues.ErrorLogEntries
        & $addRow 'SMART Health' 'Source' $smartctlValues.Source
    } elseif ([string]::IsNullOrWhiteSpace((Find-SmartctlPath))) {
        & $addRow 'SMART Health' 'smartctl' 'Unavailable. Install smartmontools or place smartctl.exe next to TEKSysInfo.ps1 for NVMe/SATA SMART details.'
    } else {
        & $addRow 'SMART Health' 'smartctl' 'Installed, but no matching SMART device report was found for this disk.'
    }

    if ($storageReliability) {
        & $addRow 'Reliability' 'Temperature' $storageReliability.Temperature
        & $addRow 'Reliability' 'Temperature max' $storageReliability.TemperatureMax
        & $addRow 'Reliability' 'Windows wear counter' $storageReliability.Wear
        & $addRow 'Reliability' 'Power on hours' $storageReliability.PowerOnHours
        & $addRow 'Reliability' 'Read errors total' $storageReliability.ReadErrorsTotal
        & $addRow 'Reliability' 'Write errors total' $storageReliability.WriteErrorsTotal
        & $addRow 'Reliability' 'Read latency max' $storageReliability.ReadLatencyMax
        & $addRow 'Reliability' 'Write latency max' $storageReliability.WriteLatencyMax
    } else {
        & $addRow 'Reliability' 'SMART/lifecycle counters' 'Unavailable from Windows storage provider for this disk.'
    }

    return $rows
}

function Get-InstalledSoftware {
    $paths = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    $items = foreach ($path in $paths) {
        try {
            Get-ItemProperty -Path $path -ErrorAction Stop |
                Where-Object {
                    (Get-ObjectPropertyValue -InputObject $_ -Name 'DisplayName' -Default '') -and
                    (Get-ObjectPropertyValue -InputObject $_ -Name 'SystemComponent' -Default 0) -ne 1
                } |
                ForEach-Object {
                    $displayName = Get-ObjectPropertyValue -InputObject $_ -Name 'DisplayName' -Default ''
                    $displayVersion = Get-ObjectPropertyValue -InputObject $_ -Name 'DisplayVersion' -Default ''
                    $publisher = Get-ObjectPropertyValue -InputObject $_ -Name 'Publisher' -Default ''
                    $rawInstallDate = Get-ObjectPropertyValue -InputObject $_ -Name 'InstallDate' -Default ''
                    $date = ''
                    if ($rawInstallDate -match '^\d{8}$') {
                        try { $date = [DateTime]::ParseExact($rawInstallDate, 'yyyyMMdd', $null).ToString('yyyy-MM-dd') } catch { $date = [string]$rawInstallDate }
                    } elseif ($rawInstallDate) {
                        $date = [string]$rawInstallDate
                    }

                    [pscustomobject]@{
                        Name = [string]$displayName
                        Version = ConvertTo-DisplayText $displayVersion ''
                        Publisher = ConvertTo-DisplayText $publisher ''
                        InstallDate = $date
                        Source = if ($path -like 'HKCU:*') { 'Current user' } elseif ($path -like '*WOW6432Node*') { 'Machine x86' } else { 'Machine' }
                    }
                }
        } catch {
            Write-Log "Software registry read failed for ${path}: $($_.Exception.Message)"
        }
    }

    $items |
        Sort-Object Name, Version -Unique |
        Sort-Object Name
}

function Get-HardwareInventory {
    $rows = New-Object System.Collections.Generic.List[object]

    foreach ($cpu in (Get-CimSafe -ClassName Win32_Processor)) {
        $rows.Add([pscustomobject]@{ Category = 'Processor'; Name = ConvertTo-DisplayText $cpu.Name; Manufacturer = ConvertTo-DisplayText $cpu.Manufacturer; Status = ConvertTo-DisplayText $cpu.Status; Detail = ('{0} cores / {1} threads / {2} MHz' -f $cpu.NumberOfCores, $cpu.NumberOfLogicalProcessors, $cpu.MaxClockSpeed) })
    }
    foreach ($mem in (Get-CimSafe -ClassName Win32_PhysicalMemory)) {
        $rows.Add([pscustomobject]@{ Category = 'Memory Module'; Name = ConvertTo-DisplayText $mem.PartNumber; Manufacturer = ConvertTo-DisplayText $mem.Manufacturer; Status = ConvertTo-DisplayText $mem.BankLabel; Detail = ('{0} {1} MHz {2}' -f (ConvertTo-SizeText ([double]$mem.Capacity)), $mem.Speed, $mem.DeviceLocator) })
    }
    foreach ($disk in (Get-CimSafe -ClassName Win32_DiskDrive)) {
        $rows.Add([pscustomobject]@{ Category = 'Physical Disk'; Name = ConvertTo-DisplayText $disk.Model; Manufacturer = ConvertTo-DisplayText $disk.Manufacturer; Status = ConvertTo-DisplayText $disk.Status; Detail = ('{0} | {1} | {2}' -f (ConvertTo-SizeText ([double]$disk.Size)), $disk.InterfaceType, $disk.SerialNumber) })
    }
    foreach ($video in (Get-CimSafe -ClassName Win32_VideoController)) {
        $rows.Add([pscustomobject]@{ Category = 'Video'; Name = ConvertTo-DisplayText $video.Name; Manufacturer = ConvertTo-DisplayText $video.AdapterCompatibility; Status = ConvertTo-DisplayText $video.Status; Detail = ('Driver {0} | RAM {1}' -f $video.DriverVersion, (ConvertTo-SizeText ([double]$video.AdapterRAM))) })
    }
    foreach ($net in (Get-CimSafe -ClassName Win32_NetworkAdapter -Filter 'PhysicalAdapter=True')) {
        $rows.Add([pscustomobject]@{ Category = 'Network Adapter'; Name = ConvertTo-DisplayText $net.Name; Manufacturer = ConvertTo-DisplayText $net.Manufacturer; Status = ConvertTo-DisplayText $net.NetConnectionStatus; Detail = ('MAC {0} | Speed {1}' -f (ConvertTo-DisplayText $net.MACAddress ''), (ConvertTo-SizeText ([double]$net.Speed))) })
    }
    foreach ($device in (Get-CimSafe -ClassName Win32_PnPEntity | Where-Object { $_.Name } | Sort-Object PNPClass, Name)) {
        $rows.Add([pscustomobject]@{ Category = ConvertTo-DisplayText $device.PNPClass 'Device'; Name = ConvertTo-DisplayText $device.Name; Manufacturer = ConvertTo-DisplayText $device.Manufacturer ''; Status = ConvertTo-DisplayText $device.Status; Detail = ConvertTo-DisplayText $device.DeviceID '' })
    }

    return $rows
}

function Get-RecentHotFixes {
    try {
        Get-HotFix -ErrorAction Stop |
            Sort-Object InstalledOn -Descending |
            Select-Object -First 80 |
            ForEach-Object {
                [pscustomobject]@{
                    HotFixId = $_.HotFixID
                    Description = $_.Description
                    InstalledOn = if ($_.InstalledOn) { ([DateTime]$_.InstalledOn).ToString('yyyy-MM-dd') } else { '' }
                    InstalledBy = $_.InstalledBy
                }
            }
    } catch {
        Write-Log "Get-HotFix failed: $($_.Exception.Message)"
        return @()
    }
}

function Search-WindowsUpdates {
    $session = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()
    $result = $searcher.Search("IsInstalled=0 and IsHidden=0")
    $updates = @()
    for ($i = 0; $i -lt $result.Updates.Count; $i++) {
        $update = $result.Updates.Item($i)
        $size = 0
        try { $size = [double]$update.MaxDownloadSize } catch { $size = 0 }
        $updates += [pscustomobject]@{
            Index = $i
            Title = [string]$update.Title
            KB = (($update.KBArticleIDs | ForEach-Object { "KB$_" }) -join ', ')
            Severity = ConvertTo-DisplayText $update.MsrcSeverity ''
            RebootRequired = [bool]$update.RebootRequired
            Size = ConvertTo-SizeText $size
            RawUpdate = $update
        }
    }
    return $updates
}

function Install-WindowsUpdates {
    param([object[]]$Updates)

    $updateList = @($Updates)
    if ($updateList.Count -eq 0) {
        throw 'No updates were selected.'
    }

    $session = New-Object -ComObject Microsoft.Update.Session
    $collection = New-Object -ComObject Microsoft.Update.UpdateColl

    foreach ($item in $updateList) {
        $update = $item.RawUpdate
        if (-not $update.EulaAccepted) {
            $update.AcceptEula()
        }
        [void]$collection.Add($update)
    }

    $downloader = $session.CreateUpdateDownloader()
    $downloader.Updates = $collection
    $downloadResult = $downloader.Download()

    $installer = $session.CreateUpdateInstaller()
    $installer.Updates = $collection
    $installResult = $installer.Install()

    return [pscustomobject]@{
        DownloadResult = $downloadResult.ResultCode
        InstallResult = $installResult.ResultCode
        RebootRequired = [bool]$installResult.RebootRequired
    }
}

function Invoke-Ui {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Control]$Control,
        [Parameter(Mandatory)][scriptblock]$Action
    )

    if ($Control.InvokeRequired) {
        [void]$Control.BeginInvoke([Action]$Action)
    } else {
        & $Action
    }
}

function Invoke-Background {
    param(
        [Parameter(Mandatory)][scriptblock]$Work,
        [Parameter(Mandatory)][scriptblock]$Done,
        [object[]]$ArgumentList = @(),
        [object]$State = $null,
        [scriptblock]$Failed
    )

    $Script:BackgroundTaskId++
    $taskName = "TEKSysInfoTask$($Script:BackgroundTaskId)"
    $workText = $Work.ToString()
    $functionNames = @(
        'Write-Log',
        'Test-IsAdministrator',
        'ConvertTo-DisplayText',
        'ConvertTo-SizeText',
        'ConvertTo-RateText',
        'ConvertTo-PercentText',
        'ConvertFrom-CimDate',
        'Get-ObjectPropertyValue',
        'Get-CimSafe',
        'Get-PrimaryDiskUsage',
        'ConvertFrom-KelvinTenths',
        'ConvertFrom-Kelvin',
        'Get-ThermalReadings',
        'Get-NetworkMetrics',
        'Get-DiskActivityMetrics',
        'Get-TopProcessMetrics',
        'Get-LiveMetrics',
        'Get-ExtendedDashboardMetrics',
        'Get-SystemSnapshot',
        'Test-PendingReboot',
        'Get-PowerSummary',
        'Get-SecuritySummary',
        'Get-DriveHealthSummary',
        'Get-LogicalDisks',
        'Format-NullableValue',
        'Format-HealthHours',
        'Get-NumericPropertyValue',
        'Get-NormalizedHardwareId',
        'Find-SmartctlPath',
        'Format-NvmeDataUnits',
        'Get-SmartctlReports',
        'Find-SmartctlReportForDisk',
        'Get-SmartctlHealthValues',
        'Get-DiskDetail',
        'Get-InstalledSoftware',
        'Get-HardwareInventory',
        'Get-RecentHotFixes',
        'Get-SystemReportData',
        'New-ReportOptions',
        'Get-ReportBrowserPath',
        'ConvertTo-HtmlCell',
        'Add-HtmlTable',
        'ConvertTo-SystemReportHtml',
        'Export-SystemReportCsv',
        'Export-SystemReportPdf',
        'Search-WindowsUpdates',
        'Install-WindowsUpdates'
    )
    $functionText = ($functionNames | ForEach-Object {
        $command = Get-Command -Name $_ -CommandType Function -ErrorAction Stop
        "function $($_) {`r`n$($command.Definition)`r`n}"
    }) -join "`r`n`r`n"

    $bootstrap = @"
Set-StrictMode -Version 2.0
`$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
`$Script:AppName = '$($Script:AppName.Replace("'", "''"))'
`$Script:ProgramDataPath = '$($Script:ProgramDataPath.Replace("'", "''"))'
`$Script:LogPath = '$($Script:LogPath.Replace("'", "''"))'
`$Script:Controls = @{}
`$Script:LastSnapshot = `$null
"@

    $script = @"
$bootstrap
$functionText
`$taskArguments = `$args
& ([scriptblock]::Create(@'
$workText
'@)) @taskArguments
"@

    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()

    $powerShell = [PowerShell]::Create()
    $powerShell.Runspace = $runspace
    [void]$powerShell.AddScript($script)
    foreach ($argument in @($ArgumentList)) {
        [void]$powerShell.AddArgument($argument)
    }
    $asyncResult = $powerShell.BeginInvoke()

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 250
    $timer.Tag = @{
        Name = $taskName
        PowerShell = $powerShell
        Runspace = $runspace
        AsyncResult = $asyncResult
        Done = $Done
        Failed = if ($PSBoundParameters.ContainsKey('Failed')) { $Failed } else { $null }
        State = $State
    }
    $timer.add_Tick({
        param($sender, $eventArgs)
        $task = $sender.Tag
        if (-not $task.AsyncResult.IsCompleted) { return }

        $sender.Stop()
        $sender.Dispose()
        try {
            $result = @($task.PowerShell.EndInvoke($task.AsyncResult))
            if ($task.PowerShell.Streams.Error.Count -gt 0) {
                $message = (($task.PowerShell.Streams.Error | ForEach-Object { $_.ToString() }) -join "`r`n")
                throw $message
            }
            & $task.Done $result $task.State
        } catch {
            if ($null -ne $task.Failed) {
                try {
                    & $task.Failed $_.Exception $task.State
                } catch {
                    Set-Status "Task error handler failed: $($_.Exception.Message)"
                    Write-Log "Task error handler failed: $($_.Exception.Message)"
                    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $Script:AppName, 'OK', 'Error') | Out-Null
                }
            } else {
                Set-Status "Task failed: $($_.Exception.Message)"
                Write-Log "Task failed: $($_.Exception.Message)"
                [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $Script:AppName, 'OK', 'Error') | Out-Null
            }
        } finally {
            try {
                $task.PowerShell.Dispose()
            } catch { }
            try {
                $task.Runspace.Close()
                $task.Runspace.Dispose()
            } catch { }
        }
    })
    $timer.Start()
}

function New-Label {
    param(
        [string]$Text,
        [int]$Size = 9,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::FromArgb(38, 50, 56)
    )

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.AutoSize = $true
    $label.Font = New-Object System.Drawing.Font('Segoe UI', $Size, $Style)
    $label.ForeColor = $Color
    $label.Margin = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
    return $label
}

function New-MetricCard {
    param(
        [string]$Key,
        [string]$Title,
        [System.Drawing.Color]$Accent
    )

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Fill'
    $panel.BackColor = [System.Drawing.Color]::White
    $panel.Margin = New-Object System.Windows.Forms.Padding(8)
    $panel.Padding = New-Object System.Windows.Forms.Padding(14)
    Enable-ControlDoubleBuffering $panel

    $titleLabel = New-Label -Text $Title -Size 10 -Style Bold -Color ([System.Drawing.Color]::FromArgb(73, 80, 87))
    $titleLabel.Dock = 'Top'

    $valueLabel = New-Label -Text 'Loading...' -Size 22 -Style Bold -Color $Accent
    $valueLabel.Dock = 'Top'
    $valueLabel.Height = 48
    $valueLabel.AutoSize = $false
    $valueLabel.TextAlign = 'MiddleLeft'

    $detailLabel = New-Label -Text '' -Size 9 -Color ([System.Drawing.Color]::FromArgb(83, 96, 107))
    $detailLabel.Dock = 'Top'
    $detailLabel.Height = 28
    $detailLabel.AutoSize = $false

    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Dock = 'Top'
    $bar.Height = 14
    $bar.Style = 'Continuous'
    $bar.Margin = New-Object System.Windows.Forms.Padding(8)

    $graph = New-Object System.Windows.Forms.Panel
    $graph.Dock = 'Fill'
    $graph.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $graph.Tag = @{ History = $null; Accent = $Accent }
    Enable-ControlDoubleBuffering $graph
    $graph.add_Paint({
        param($sender, $eventArgs)
        $data = $sender.Tag
        Draw-HistoryGraph -Graphics $eventArgs.Graphics -Bounds $sender.ClientRectangle -History $data.History -Accent $data.Accent
    })

    $panel.Controls.Add($graph)
    $panel.Controls.Add($bar)
    $panel.Controls.Add($detailLabel)
    $panel.Controls.Add($valueLabel)
    $panel.Controls.Add($titleLabel)

    $Script:Controls["${Key}Value"] = $valueLabel
    $Script:Controls["${Key}Detail"] = $detailLabel
    $Script:Controls["${Key}Bar"] = $bar
    $Script:Controls["${Key}Graph"] = $graph

    return $panel
}

function New-DashboardTile {
    param(
        [string]$Key,
        [string]$Title,
        [System.Drawing.Color]$Accent
    )

    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Fill'
    $panel.BackColor = [System.Drawing.Color]::White
    $panel.Margin = New-Object System.Windows.Forms.Padding(8)
    $panel.Padding = New-Object System.Windows.Forms.Padding(14)
    $panel.BorderStyle = 'FixedSingle'
    Enable-ControlDoubleBuffering $panel

    $titleLabel = New-Label -Text $Title -Size 9 -Style Bold -Color ([System.Drawing.Color]::FromArgb(73, 80, 87))
    $titleLabel.Dock = 'Top'

    $valueLabel = New-Label -Text 'Loading...' -Size 18 -Style Bold -Color $Accent
    $valueLabel.Dock = 'Top'
    $valueLabel.Height = 38
    $valueLabel.AutoSize = $false
    $valueLabel.TextAlign = 'MiddleLeft'

    $detailLabel = New-Label -Text '' -Size 9 -Color ([System.Drawing.Color]::FromArgb(83, 96, 107))
    $detailLabel.Dock = 'Fill'
    $detailLabel.AutoSize = $false

    $panel.Controls.Add($detailLabel)
    $panel.Controls.Add($valueLabel)
    $panel.Controls.Add($titleLabel)

    $Script:Controls["${Key}TileValue"] = $valueLabel
    $Script:Controls["${Key}TileDetail"] = $detailLabel

    return $panel
}

function Draw-HistoryGraph {
    param(
        [System.Drawing.Graphics]$Graphics,
        [System.Drawing.Rectangle]$Bounds,
        [System.Collections.ArrayList]$History,
        [System.Drawing.Color]$Accent
    )

    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $Graphics.Clear([System.Drawing.Color]::FromArgb(248, 250, 252))
    if ($Bounds.Width -lt 10 -or $Bounds.Height -lt 10) { return }

    $gridPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(226, 232, 240), 1)
    try {
        for ($i = 1; $i -lt 4; $i++) {
            $y = $Bounds.Top + (($Bounds.Height / 4) * $i)
            $Graphics.DrawLine($gridPen, $Bounds.Left, [int]$y, $Bounds.Right, [int]$y)
        }
    } finally {
        $gridPen.Dispose()
    }

    if (-not $History -or $History.Count -lt 2) { return }

    $points = New-Object System.Collections.Generic.List[System.Drawing.PointF]
    $max = [Math]::Max(2, $History.Count - 1)
    for ($i = 0; $i -lt $History.Count; $i++) {
        $x = $Bounds.Left + (($Bounds.Width - 2) * ($i / $max))
        $y = $Bounds.Bottom - 2 - (($Bounds.Height - 4) * ([double]$History[$i] / 100))
        $points.Add((New-Object System.Drawing.PointF([float]$x, [float]$y)))
    }

    $pen = New-Object System.Drawing.Pen($Accent, 2)
    try {
        $Graphics.DrawLines($pen, $points.ToArray())
    } finally {
        $pen.Dispose()
    }
}

function New-ListView {
    param([string[]]$Columns)

    $list = New-Object System.Windows.Forms.ListView
    $list.View = 'Details'
    $list.FullRowSelect = $true
    $list.GridLines = $true
    $list.Dock = 'Fill'
    $list.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    foreach ($column in $Columns) {
        [void]$list.Columns.Add($column, 150)
    }
    Enable-ControlDoubleBuffering $list
    return $list
}

function Add-ListViewRow {
    param(
        [System.Windows.Forms.ListView]$ListView,
        [object[]]$Values,
        [object]$Tag
    )

    $item = New-Object System.Windows.Forms.ListViewItem([string]$Values[0])
    for ($i = 1; $i -lt $Values.Count; $i++) {
        [void]$item.SubItems.Add([string]$Values[$i])
    }
    $item.Tag = $Tag
    [void]$ListView.Items.Add($item)
}

function Resize-ListColumns {
    param([System.Windows.Forms.ListView]$ListView)

    foreach ($column in $ListView.Columns) {
        try { $column.Width = -2 } catch { }
    }
}

function Invoke-UiYield {
    try {
        [System.Windows.Forms.Application]::DoEvents()
    } catch {
        # DoEvents is best-effort to keep the form responsive during large fills.
    }
}

function Set-Status {
    param([string]$Text)
    if ($Script:Controls.StatusLabel) {
        $Script:Controls.StatusLabel.Text = $Text
    }
}

function Populate-Details {
    param([object]$Snapshot)

    $list = $Script:Controls.DetailsList
    $list.BeginUpdate()
    try {
        $list.Items.Clear()
        $properties = @(
            'ComputerName','UserName','Manufacturer','Model','SerialNumber','ChassisSerial',
            'BiosVersion','BiosDate','BaseBoard','Processor','ProcessorCores','ProcessorThreads',
            'TotalMemory','WindowsCaption','WindowsVersion','WindowsArchitecture','InstallDate',
            'LastBoot','SecureBoot','IsAdministrator'
        )
        foreach ($name in $properties) {
            $value = $Snapshot.$name
            if ($name -eq 'IsAdministrator') { $value = if ($value) { 'Yes' } else { 'No' } }
            Add-ListViewRow -ListView $list -Values @($name, $value) -Tag $null
        }
        if ($Snapshot.Uptime) {
            Add-ListViewRow -ListView $list -Values @('Uptime', ('{0}d {1}h {2}m' -f $Snapshot.Uptime.Days, $Snapshot.Uptime.Hours, $Snapshot.Uptime.Minutes)) -Tag $null
        }
        Resize-ListColumns $list
    } finally {
        $list.EndUpdate()
    }
}

function Populate-Disks {
    param([object[]]$Disks)

    $list = $Script:Controls.DiskList
    $list.BeginUpdate()
    try {
        $list.Items.Clear()
        foreach ($disk in $Disks) {
            Add-ListViewRow -ListView $list -Values @($disk.Drive, $disk.Label, $disk.FileSystem, $disk.Size, $disk.Used, $disk.Free, (ConvertTo-PercentText $disk.UsedPercent)) -Tag $disk
        }
        Resize-ListColumns $list
    } finally {
        $list.EndUpdate()
    }
}

function Populate-Software {
    param([object[]]$Software)

    $Script:Controls.SoftwareItems = $Software
    Apply-SoftwareFilter
}

function Apply-SoftwareFilter {
    $software = @($Script:Controls.SoftwareItems)
    $query = ''
    if ($Script:Controls.SoftwareSearch) { $query = $Script:Controls.SoftwareSearch.Text.Trim() }
    if ($query) {
        $software = $software | Where-Object {
            $_.Name -like "*$query*" -or $_.Publisher -like "*$query*" -or $_.Version -like "*$query*"
        }
    }

    $list = $Script:Controls.SoftwareList
    $list.BeginUpdate()
    try {
        $list.Items.Clear()
        $index = 0
        foreach ($app in $software) {
            Add-ListViewRow -ListView $list -Values @($app.Name, $app.Version, $app.Publisher, $app.InstallDate, $app.Source) -Tag $app
            $index++
            if (($index % 150) -eq 0) { Invoke-UiYield }
        }
        Resize-ListColumns $list
        $Script:Controls.SoftwareCountLabel.Text = ('{0:N0} programs' -f $software.Count)
    } finally {
        $list.EndUpdate()
    }
}

function Populate-Hardware {
    param([object[]]$Hardware)

    $Script:Controls.HardwareItems = $Hardware
    Apply-HardwareFilter
}

function Apply-HardwareFilter {
    $hardware = @($Script:Controls.HardwareItems)
    $query = ''
    if ($Script:Controls.HardwareSearch) { $query = $Script:Controls.HardwareSearch.Text.Trim() }
    if ($query) {
        $hardware = $hardware | Where-Object {
            $_.Category -like "*$query*" -or $_.Name -like "*$query*" -or $_.Manufacturer -like "*$query*" -or $_.Detail -like "*$query*"
        }
    }

    $list = $Script:Controls.HardwareList
    $list.BeginUpdate()
    try {
        $list.Items.Clear()
        $index = 0
        foreach ($item in $hardware) {
            Add-ListViewRow -ListView $list -Values @($item.Category, $item.Name, $item.Manufacturer, $item.Status, $item.Detail) -Tag $item
            $index++
            if (($index % 150) -eq 0) { Invoke-UiYield }
        }
        Resize-ListColumns $list
        $Script:Controls.HardwareCountLabel.Text = ('{0:N0} devices' -f $hardware.Count)
    } finally {
        $list.EndUpdate()
    }
}

function Populate-HotFixes {
    param([object[]]$HotFixes)

    $list = $Script:Controls.HotFixList
    $list.BeginUpdate()
    try {
        $list.Items.Clear()
        foreach ($hotfix in $HotFixes) {
            Add-ListViewRow -ListView $list -Values @($hotfix.HotFixId, $hotfix.Description, $hotfix.InstalledOn, $hotfix.InstalledBy) -Tag $hotfix
        }
        Resize-ListColumns $list
    } finally {
        $list.EndUpdate()
    }
}

function Show-DiskDetailWindow {
    param([object]$Disk)

    if (-not $Disk -or -not $Disk.Drive) {
        [System.Windows.Forms.MessageBox]::Show('Select a fixed disk first.', $Script:AppName, 'OK', 'Information') | Out-Null
        return
    }

    $detailForm = New-Object System.Windows.Forms.Form
    $detailForm.Text = "Disk Details - $($Disk.Drive)"
    $detailForm.StartPosition = 'CenterParent'
    $detailForm.Size = New-Object System.Drawing.Size(820, 620)
    $detailForm.MinimumSize = New-Object System.Drawing.Size(680, 480)
    $detailForm.BackColor = [System.Drawing.Color]::White
    $detailForm.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.RowCount = 3
    $layout.ColumnCount = 1
    $layout.Padding = New-Object System.Windows.Forms.Padding(12)
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 54)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))
    $detailForm.Controls.Add($layout)

    $heading = New-Label -Text ("{0} {1} - {2}" -f $Disk.Drive, $Disk.Label, $Disk.Size) -Size 14 -Style Bold -Color ([System.Drawing.Color]::FromArgb(15, 23, 42))
    $heading.Dock = 'Fill'
    $heading.AutoSize = $false
    $heading.TextAlign = 'MiddleLeft'
    $layout.Controls.Add($heading, 0, 0)

    $detailList = New-ListView -Columns @('Category','Property','Value')
    $layout.Controls.Add($detailList, 0, 1)

    $footer = New-Object System.Windows.Forms.FlowLayoutPanel
    $footer.Dock = 'Fill'
    $footer.FlowDirection = 'RightToLeft'
    $closeButton = New-ToolbarButton -Text 'Close'
    $closeButton.Tag = $detailForm
    $closeButton.add_Click({
        param($sender, $eventArgs)
        if ($sender.Tag) { $sender.Tag.Close() }
    })
    $footer.Controls.Add($closeButton)
    $layout.Controls.Add($footer, 0, 2)

    $detailList.Items.Add((New-Object System.Windows.Forms.ListViewItem('Loading disk details...'))) | Out-Null
    $detailForm.Show($Script:Controls.MainForm)

    $detailState = [pscustomobject]@{
        Form = $detailForm
        List = $detailList
        Drive = $Disk.Drive
    }

    Invoke-Background -Work {
        param($drive)
        @(Get-DiskDetail -DriveLetter $drive)
    } -Done {
        param($rows, $state)
        $list = $state.List
        if (-not $list -or $list.IsDisposed) { return }
        $list.BeginUpdate()
        try {
            $list.Items.Clear()
            foreach ($row in @($rows)) {
                Add-ListViewRow -ListView $list -Values @($row.Category, $row.Property, $row.Value) -Tag $row
            }
            Resize-ListColumns $list
        } finally {
            $list.EndUpdate()
        }
        Set-Status "Loaded disk details for $($state.Drive)"
    } -ArgumentList @($Disk.Drive) -State $detailState -Failed {
        param($errorRecord, $state)
        $list = $state.List
        if ($list -and -not $list.IsDisposed) {
            $list.Items.Clear()
            Add-ListViewRow -ListView $list -Values @('Error', 'Disk details failed', $errorRecord.Message) -Tag $null
            Resize-ListColumns $list
        }
        Set-Status "Disk details failed: $($errorRecord.Message)"
    }
}

function Show-SelectedDiskDetails {
    $list = $Script:Controls.DiskList
    if (-not $list -or $list.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Select a fixed disk first.', $Script:AppName, 'OK', 'Information') | Out-Null
        return
    }
    Show-DiskDetailWindow -Disk $list.SelectedItems[0].Tag
}

function Populate-AvailableUpdates {
    param([object[]]$Updates)

    $Script:AvailableUpdates = @($Updates)
    $updateCount = @($Script:AvailableUpdates).Count
    $list = $Script:Controls.AvailableUpdateList
    $list.BeginUpdate()
    try {
        $list.Items.Clear()
        foreach ($update in $Script:AvailableUpdates) {
            $rebootRequired = if ($update.RebootRequired) { 'Yes' } else { 'No' }
            Add-ListViewRow -ListView $list -Values @($update.Title, $update.KB, $update.Severity, $update.Size, $rebootRequired) -Tag $update
        }
        Resize-ListColumns $list
        $Script:Controls.UpdateSummaryLabel.Text = if ($updateCount -eq 0) { 'No applicable updates found.' } else { ('{0:N0} applicable updates found.' -f $updateCount) }
    } finally {
        $list.EndUpdate()
    }
}

function Populate-DashboardOverview {
    param(
        [object]$Snapshot,
        [object]$PendingReboot,
        [object]$Power,
        [object]$Security,
        [object]$DriveHealth
    )

    if (-not $Script:Controls.ContainsKey('DashboardOverviewList')) { return }
    $list = $Script:Controls.DashboardOverviewList
    $list.BeginUpdate()
    try {
        $list.Items.Clear()
        $uptimeText = if ($Snapshot.Uptime) { 'Uptime {0}d {1}h {2}m' -f $Snapshot.Uptime.Days, $Snapshot.Uptime.Hours, $Snapshot.Uptime.Minutes } else { 'Uptime unknown' }
        Add-ListViewRow -ListView $list -Values @('Computer', ('{0} - {1}' -f $Snapshot.Manufacturer, $Snapshot.Model), $Snapshot.ComputerName) -Tag $null
        Add-ListViewRow -ListView $list -Values @('Operating System', $Snapshot.WindowsCaption, $Snapshot.WindowsVersion) -Tag $null
        Add-ListViewRow -ListView $list -Values @('Processor', $Snapshot.Processor, ('{0} cores / {1} threads' -f $Snapshot.ProcessorCores, $Snapshot.ProcessorThreads)) -Tag $null
        Add-ListViewRow -ListView $list -Values @('Memory', $Snapshot.TotalMemory, 'Physical memory') -Tag $null
        Add-ListViewRow -ListView $list -Values @('Boot', $Snapshot.LastBoot, $uptimeText) -Tag $null
        Add-ListViewRow -ListView $list -Values @('Security', ('Secure Boot {0}' -f $Snapshot.SecureBoot), ('Administrator: {0}' -f $(if ($Snapshot.IsAdministrator) { 'Yes' } else { 'No' }))) -Tag $null
        if ($PendingReboot) { Add-ListViewRow -ListView $list -Values @('Pending Reboot', $PendingReboot.Text, '') -Tag $PendingReboot }
        if ($Power) { Add-ListViewRow -ListView $list -Values @('Power', $Power.Status, $Power.Detail) -Tag $Power }
        if ($Security) { Add-ListViewRow -ListView $list -Values @('Security Detail', $Security.Status, $Security.Detail) -Tag $Security }
        if ($DriveHealth) { Add-ListViewRow -ListView $list -Values @('Drive Health', $DriveHealth.Status, $DriveHealth.Detail) -Tag $DriveHealth }
        Resize-ListColumns $list
    } finally {
        $list.EndUpdate()
    }

    if ($PendingReboot -and $Script:Controls.ContainsKey('RebootTileValue')) {
        $Script:Controls.RebootTileValue.Text = if ($PendingReboot.IsPending) { 'Pending' } else { 'Clear' }
        $Script:Controls.RebootTileDetail.Text = $PendingReboot.Text
    }
    if ($DriveHealth -and $Script:Controls.ContainsKey('DriveHealthTileValue')) {
        $Script:Controls.DriveHealthTileValue.Text = $DriveHealth.Status
        $Script:Controls.DriveHealthTileDetail.Text = ConvertTo-DisplayText $DriveHealth.Detail 'Windows drive health summary'
    }
    if ($Power -and $Script:Controls.ContainsKey('PowerTileValue')) {
        $Script:Controls.PowerTileValue.Text = $Power.Status
        $Script:Controls.PowerTileDetail.Text = $Power.Detail
    }
    if ($Security -and $Script:Controls.ContainsKey('SecurityTileValue')) {
        $Script:Controls.SecurityTileValue.Text = $Security.Status
        $Script:Controls.SecurityTileDetail.Text = $Security.Detail
    }
}

function Populate-LiveProcessList {
    param([object]$TopProcesses)

    if (-not $Script:Controls.ContainsKey('ProcessList')) { return }
    $list = $Script:Controls.ProcessList
    $list.BeginUpdate()
    try {
        $list.Items.Clear()
        Add-ListViewRow -ListView $list -Values @('Top CPU', $TopProcesses.TopCpuName, ('{0:N1}% CPU' -f [double]$TopProcesses.TopCpuPercent)) -Tag $TopProcesses
        Add-ListViewRow -ListView $list -Values @('Top Memory', $TopProcesses.TopMemoryName, (ConvertTo-SizeText $TopProcesses.TopMemoryBytes)) -Tag $TopProcesses
        Resize-ListColumns $list
    } finally {
        $list.EndUpdate()
    }
}

function Populate-LiveIoList {
    param(
        [object]$Network,
        [object]$DiskActivity
    )

    if (-not $Script:Controls.ContainsKey('IoList')) { return }
    $list = $Script:Controls.IoList
    $list.BeginUpdate()
    try {
        $list.Items.Clear()
        Add-ListViewRow -ListView $list -Values @('Network down', (ConvertTo-RateText $Network.BytesInPerSec), $Network.ActiveAdapter) -Tag $Network
        Add-ListViewRow -ListView $list -Values @('Network up', (ConvertTo-RateText $Network.BytesOutPerSec), $Network.ActiveAdapter) -Tag $Network
        Add-ListViewRow -ListView $list -Values @('Disk reads', (ConvertTo-RateText $DiskActivity.ReadBytesPerSec), 'All physical disks') -Tag $DiskActivity
        Add-ListViewRow -ListView $list -Values @('Disk writes', (ConvertTo-RateText $DiskActivity.WriteBytesPerSec), ('Queue {0:N0}' -f [double]$DiskActivity.QueueLength)) -Tag $DiskActivity
        Resize-ListColumns $list
    } finally {
        $list.EndUpdate()
    }
}

function Populate-ThermalReadings {
    param([object[]]$Readings)

    if (-not $Script:Controls.ContainsKey('ThermalList')) { return }
    $list = $Script:Controls.ThermalList
    $list.BeginUpdate()
    try {
        $list.Items.Clear()
        $thermalReadings = @($Readings)
        if ($thermalReadings.Count -eq 0) {
            Add-ListViewRow -ListView $list -Values @('Unavailable', '', 'No thermal sensors exposed by Windows', '') -Tag $null
        } else {
            foreach ($reading in $thermalReadings) {
                Add-ListViewRow -ListView $list -Values @($reading.Name, ('{0:N1} C' -f [double]$reading.TemperatureC), $reading.Source, $reading.Detail) -Tag $reading
            }
        }
        Resize-ListColumns $list
    } finally {
        $list.EndUpdate()
    }
}

function Update-LiveDashboardFromMetrics {
    param([Parameter(Mandatory)][object]$Metrics)

    if ($Script:IsClosing) { return }

    Add-HistoryValue -History $Script:CpuHistory -Value $Metrics.CpuPercent
    Add-HistoryValue -History $Script:MemoryHistory -Value $Metrics.MemoryPercent
    Add-HistoryValue -History $Script:DiskHistory -Value $Metrics.DiskPercent
    if ($null -ne $Metrics.HighestTemperatureC) { Add-HistoryValue -History $Script:TemperatureHistory -Value ([Math]::Min(100, [double]$Metrics.HighestTemperatureC)) }

    if ($Script:Controls.ContainsKey('CpuValue')) {
        $Script:Controls.CpuValue.Text = ConvertTo-PercentText $Metrics.CpuPercent
        $Script:Controls.CpuDetail.Text = 'Total processor load'
        $Script:Controls.CpuBar.Value = [int][Math]::Max(0, [Math]::Min(100, $Metrics.CpuPercent))
        $Script:Controls.CpuGraph.Tag.History = $Script:CpuHistory
        $Script:Controls.CpuGraph.Invalidate()
    }

    if ($Script:Controls.ContainsKey('MemoryValue')) {
        $Script:Controls.MemoryValue.Text = ConvertTo-PercentText $Metrics.MemoryPercent
        $Script:Controls.MemoryDetail.Text = ('{0} used / {1} total' -f (ConvertTo-SizeText $Metrics.MemoryUsed), (ConvertTo-SizeText $Metrics.MemoryTotal))
        $Script:Controls.MemoryBar.Value = [int][Math]::Max(0, [Math]::Min(100, $Metrics.MemoryPercent))
        $Script:Controls.MemoryGraph.Tag.History = $Script:MemoryHistory
        $Script:Controls.MemoryGraph.Invalidate()
    }

    if ($Script:Controls.ContainsKey('DiskValue')) {
        $Script:Controls.DiskValue.Text = ConvertTo-PercentText $Metrics.DiskPercent
        if ($Metrics.Disk) {
            $Script:Controls.DiskDetail.Text = ('{0} used, {1} free on {2}' -f (ConvertTo-SizeText ($Metrics.Disk.Size - $Metrics.Disk.Free)), (ConvertTo-SizeText $Metrics.Disk.Free), $Metrics.Disk.DeviceId)
        } else {
            $Script:Controls.DiskDetail.Text = 'Primary fixed disk unavailable'
        }
        $Script:Controls.DiskBar.Value = [int][Math]::Max(0, [Math]::Min(100, $Metrics.DiskPercent))
        $Script:Controls.DiskGraph.Tag.History = $Script:DiskHistory
        $Script:Controls.DiskGraph.Invalidate()
    }

    if ($Script:Controls.ContainsKey('CpuTileValue')) {
        $Script:Controls.CpuTileValue.Text = ConvertTo-PercentText $Metrics.CpuPercent
        $Script:Controls.CpuTileDetail.Text = 'Total processor load'
    }
    if ($Script:Controls.ContainsKey('MemoryTileValue')) {
        $Script:Controls.MemoryTileValue.Text = ConvertTo-PercentText $Metrics.MemoryPercent
        $Script:Controls.MemoryTileDetail.Text = ('{0} used / {1} total' -f (ConvertTo-SizeText $Metrics.MemoryUsed), (ConvertTo-SizeText $Metrics.MemoryTotal))
    }
    if ($Script:Controls.ContainsKey('DiskTileValue')) {
        $Script:Controls.DiskTileValue.Text = ConvertTo-PercentText $Metrics.DiskPercent
        $Script:Controls.DiskTileDetail.Text = if ($Metrics.Disk) { ('{0} free on {1}' -f (ConvertTo-SizeText $Metrics.Disk.Free), $Metrics.Disk.DeviceId) } else { 'Primary fixed disk unavailable' }
    }
    if ($Script:Controls.ContainsKey('ThermalTileValue')) {
        if ($null -ne $Metrics.HighestTemperatureC) {
            $Script:Controls.ThermalTileValue.Text = ('{0:N1} C' -f [double]$Metrics.HighestTemperatureC)
            $Script:Controls.ThermalTileDetail.Text = 'Highest Windows thermal-zone reading'
        } else {
            $Script:Controls.ThermalTileValue.Text = 'Unavailable'
            $Script:Controls.ThermalTileDetail.Text = 'No thermal sensors exposed by Windows'
        }
    }
    Populate-ThermalReadings -Readings $Metrics.ThermalReadings
}

function Update-ExtendedDashboardFromMetrics {
    param([Parameter(Mandatory)][object]$Metrics)

    if ($Script:IsClosing) { return }
    Add-HistoryValue -History $Script:NetworkHistory -Value ([Math]::Min(100, ([double]$Metrics.Network.BytesTotalPerSec / 1MB) * 10))
    Add-HistoryValue -History $Script:DiskIoHistory -Value ([Math]::Min(100, ([double]$Metrics.DiskActivity.TotalBytesPerSec / 1MB) * 10))

    if ($Script:Controls.ContainsKey('NetworkTileValue')) {
        $Script:Controls.NetworkTileValue.Text = ConvertTo-RateText $Metrics.Network.BytesTotalPerSec
        $Script:Controls.NetworkTileDetail.Text = ('Down {0} / Up {1}' -f (ConvertTo-RateText $Metrics.Network.BytesInPerSec), (ConvertTo-RateText $Metrics.Network.BytesOutPerSec))
    }
    if ($Script:Controls.ContainsKey('DiskIoTileValue')) {
        $Script:Controls.DiskIoTileValue.Text = ConvertTo-RateText $Metrics.DiskActivity.TotalBytesPerSec
        $Script:Controls.DiskIoTileDetail.Text = ('Read {0} / Write {1}' -f (ConvertTo-RateText $Metrics.DiskActivity.ReadBytesPerSec), (ConvertTo-RateText $Metrics.DiskActivity.WriteBytesPerSec))
    }
    if ($Script:Controls.ContainsKey('TopCpuTileValue')) {
        $Script:Controls.TopCpuTileValue.Text = $Metrics.TopProcesses.TopCpuName
        $Script:Controls.TopCpuTileDetail.Text = ('{0:N1}% CPU' -f [double]$Metrics.TopProcesses.TopCpuPercent)
    }
    if ($Script:Controls.ContainsKey('TopMemTileValue')) {
        $Script:Controls.TopMemTileValue.Text = $Metrics.TopProcesses.TopMemoryName
        $Script:Controls.TopMemTileDetail.Text = ConvertTo-SizeText $Metrics.TopProcesses.TopMemoryBytes
    }
    Populate-LiveProcessList -TopProcesses $Metrics.TopProcesses
    Populate-LiveIoList -Network $Metrics.Network -DiskActivity $Metrics.DiskActivity
}

function Start-LiveDashboardRefresh {
    if ($Script:IsClosing -or $Script:LiveMetricsInProgress) { return }
    $Script:LiveMetricsInProgress = $true

    Invoke-Background -Work {
        Get-LiveMetrics -Fast
    } -Done {
        param($result)
        try {
            $metrics = @($result) | Select-Object -First 1
            if ($metrics) { Update-LiveDashboardFromMetrics -Metrics $metrics }
        } finally {
            $Script:LiveMetricsInProgress = $false
        }
    } -Failed {
        param($errorRecord)
        $Script:LiveMetricsInProgress = $false
        Set-Status "Live metric refresh failed: $($errorRecord.Message)"
        Write-Log "Live metric refresh failed: $($errorRecord.Message)"
    }
}

function Start-ExtendedDashboardRefresh {
    if ($Script:IsClosing -or $Script:ExtendedMetricsInProgress) { return }
    $Script:ExtendedMetricsInProgress = $true

    Invoke-Background -Work {
        Get-ExtendedDashboardMetrics
    } -Done {
        param($result)
        try {
            $metrics = @($result) | Select-Object -First 1
            if ($metrics) { Update-ExtendedDashboardFromMetrics -Metrics $metrics }
        } finally {
            $Script:ExtendedMetricsInProgress = $false
        }
    } -Failed {
        param($errorRecord)
        $Script:ExtendedMetricsInProgress = $false
        Set-Status "Extended metric refresh failed: $($errorRecord.Message)"
        Write-Log "Extended metric refresh failed: $($errorRecord.Message)"
    }
}

function Update-LiveDashboard {
    try {
        $metrics = Get-LiveMetrics
        Update-LiveDashboardFromMetrics -Metrics $metrics
    } catch {
        Set-Status "Live metric refresh failed: $($_.Exception.Message)"
        Write-Log "Live metric refresh failed: $($_.Exception.Message)"
    }
}

function Complete-InventoryTask {
    param([string]$Message)

    $Script:InventoryPendingTasks = [Math]::Max(0, $Script:InventoryPendingTasks - 1)
    if ($Script:InventoryPendingTasks -eq 0) {
        $Script:InventoryRefreshInProgress = $false
        if ($Script:Controls.ContainsKey('RefreshButton')) { $Script:Controls.RefreshButton.Enabled = $true }
        Set-Status ('Inventory refreshed at {0}' -f (Get-Date -Format 'HH:mm:ss'))
    } elseif (-not [string]::IsNullOrWhiteSpace($Message)) {
        Set-Status $Message
    }
}

function Refresh-Inventory {
    if ($Script:IsClosing -or $Script:InventoryRefreshInProgress) { return }
    $Script:InventoryRefreshInProgress = $true
    $Script:InventoryPendingTasks = 4
    Set-Status 'Collecting system inventory...'
    $Script:Controls.RefreshButton.Enabled = $false

    Invoke-Background -Work {
        [pscustomobject]@{
            Snapshot = Get-SystemSnapshot
            Disks = @(Get-LogicalDisks)
            PendingReboot = Test-PendingReboot
            Power = Get-PowerSummary
            Security = Get-SecuritySummary
            DriveHealth = Get-DriveHealthSummary
        }
    } -Done {
        param($result)
        if ($Script:IsClosing) { Complete-InventoryTask; return }
        $Script:LastSnapshot = $result.Snapshot
        Populate-DashboardOverview -Snapshot $result.Snapshot -PendingReboot $result.PendingReboot -Power $result.Power -Security $result.Security -DriveHealth $result.DriveHealth
        Populate-Details -Snapshot $result.Snapshot
        Populate-Disks -Disks $result.Disks
        Complete-InventoryTask -Message 'Core inventory loaded; loading software, hardware, and updates...'
    } -Failed {
        param($errorRecord)
        if (-not $Script:IsClosing) {
            Set-Status "Core inventory failed: $($errorRecord.Message)"
            Write-Log "Core inventory failed: $($errorRecord.Message)"
        }
        Complete-InventoryTask
    }

    Invoke-Background -Work {
        @(Get-InstalledSoftware)
    } -Done {
        param($software)
        if (-not $Script:IsClosing) { Populate-Software -Software $software }
        Complete-InventoryTask -Message 'Software inventory loaded; still collecting remaining sections...'
    } -Failed {
        param($errorRecord)
        if (-not $Script:IsClosing) {
            Set-Status "Software inventory failed: $($errorRecord.Message)"
            Write-Log "Software inventory failed: $($errorRecord.Message)"
        }
        Complete-InventoryTask
    }

    Invoke-Background -Work {
        @(Get-HardwareInventory)
    } -Done {
        param($hardware)
        if (-not $Script:IsClosing) { Populate-Hardware -Hardware $hardware }
        Complete-InventoryTask -Message 'Hardware inventory loaded; still collecting remaining sections...'
    } -Failed {
        param($errorRecord)
        if (-not $Script:IsClosing) {
            Set-Status "Hardware inventory failed: $($errorRecord.Message)"
            Write-Log "Hardware inventory failed: $($errorRecord.Message)"
        }
        Complete-InventoryTask
    }

    Invoke-Background -Work {
        @(Get-RecentHotFixes)
    } -Done {
        param($hotfixes)
        if (-not $Script:IsClosing) { Populate-HotFixes -HotFixes $hotfixes }
        Complete-InventoryTask -Message 'Update history loaded; still collecting remaining sections...'
    } -Failed {
        param($errorRecord)
        if (-not $Script:IsClosing) {
            Set-Status "Update history failed: $($errorRecord.Message)"
            Write-Log "Update history failed: $($errorRecord.Message)"
        }
        Complete-InventoryTask
    }
}

function Start-UpdateScan {
    $Script:Controls.ScanUpdatesButton.Enabled = $false
    $Script:Controls.InstallUpdatesButton.Enabled = $false
    $Script:Controls.UpdateSummaryLabel.Text = 'Scanning Windows Update...'
    Set-Status 'Scanning Windows Update...'

    Invoke-Background -Work {
        @(Search-WindowsUpdates)
    } -Done {
        param($updates)
        $updateList = @($updates)
        Populate-AvailableUpdates -Updates $updateList
        $Script:Controls.ScanUpdatesButton.Enabled = $true
        $Script:Controls.InstallUpdatesButton.Enabled = ($updateList.Count -gt 0)
        Set-Status ('Windows Update scan completed at {0}' -f (Get-Date -Format 'HH:mm:ss'))
    } -Failed {
        param($errorRecord)
        $Script:Controls.ScanUpdatesButton.Enabled = $true
        $Script:Controls.InstallUpdatesButton.Enabled = (@($Script:AvailableUpdates).Count -gt 0)
        $Script:Controls.UpdateSummaryLabel.Text = 'Windows Update scan failed.'
        Set-Status "Windows Update scan failed: $($errorRecord.Message)"
        [System.Windows.Forms.MessageBox]::Show($errorRecord.Message, $Script:AppName, 'OK', 'Error') | Out-Null
    }
}

function Start-UpdateInstall {
    $selected = @()
    foreach ($item in $Script:Controls.AvailableUpdateList.SelectedItems) {
        $selected += $item.Tag
    }
    $selected = @($selected)
    if ($selected.Count -eq 0) { $selected = @($Script:AvailableUpdates) }
    $selectedCount = @($selected).Count
    if ($selectedCount -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No updates are available to install.', $Script:AppName, 'OK', 'Information') | Out-Null
        return
    }

    $adminText = if (Test-IsAdministrator) { '' } else { "`r`n`r`nThis session is not elevated. Windows may reject installation or prompt for elevation." }
    $message = "Install $selectedCount selected/applicable update(s)?`r`n`r`nThis can take several minutes and may require a reboot.$adminText"
    $choice = [System.Windows.Forms.MessageBox]::Show($message, $Script:AppName, 'YesNo', 'Warning')
    if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $Script:Controls.ScanUpdatesButton.Enabled = $false
    $Script:Controls.InstallUpdatesButton.Enabled = $false
    $Script:Controls.UpdateSummaryLabel.Text = 'Downloading and installing updates...'
    Set-Status 'Installing Windows updates...'

    Invoke-Background -Work {
        param($updatesToInstall)
        Install-WindowsUpdates -Updates $updatesToInstall
    } -Done {
        param($result)
        $Script:Controls.ScanUpdatesButton.Enabled = $true
        $Script:Controls.InstallUpdatesButton.Enabled = $false
        $Script:Controls.UpdateSummaryLabel.Text = ('Install result: {0}. Reboot required: {1}' -f $result.InstallResult, $result.RebootRequired)
        Set-Status 'Windows Update installation completed.'
        [System.Windows.Forms.MessageBox]::Show(('Install result: {0}. Reboot required: {1}' -f $result.InstallResult, $result.RebootRequired), $Script:AppName, 'OK', 'Information') | Out-Null
        Start-UpdateScan
    } -ArgumentList (,$selected) -Failed {
        param($errorRecord)
        $Script:Controls.ScanUpdatesButton.Enabled = $true
        $Script:Controls.InstallUpdatesButton.Enabled = (@($Script:AvailableUpdates).Count -gt 0)
        $Script:Controls.UpdateSummaryLabel.Text = 'Windows Update installation failed.'
        Set-Status "Windows Update installation failed: $($errorRecord.Message)"
        [System.Windows.Forms.MessageBox]::Show($errorRecord.Message, $Script:AppName, 'OK', 'Error') | Out-Null
    }
}

function Get-SystemReportData {
    $snapshot = if ($Script:LastSnapshot) { $Script:LastSnapshot } else { Get-SystemSnapshot }
    [object[]]$software = @()
    if ($Script:Controls.ContainsKey('SoftwareItems')) { $software = @($Script:Controls.SoftwareItems) }
    if ($software.Count -eq 0) { $software = @(Get-InstalledSoftware) }
    [object[]]$hardware = @()
    if ($Script:Controls.ContainsKey('HardwareItems')) { $hardware = @($Script:Controls.HardwareItems) }
    if ($hardware.Count -eq 0) { $hardware = @(Get-HardwareInventory) }

    return [pscustomobject]@{
        GeneratedAt = Get-Date
        Snapshot = $snapshot
        Metrics = Get-LiveMetrics
        Disks = @(Get-LogicalDisks)
        Software = $software
        Hardware = $hardware
        HotFixes = @(Get-RecentHotFixes)
    }
}

function New-ReportOptions {
    param([string]$Level = 'Standard')

    switch ($Level) {
        'Summary' {
            return [pscustomobject]@{
                Level = 'Summary'
                IncludeDashboard = $true
                IncludeSystem = $true
                IncludeDisks = $false
                IncludeHardware = $false
                IncludeUpdates = $false
                IncludeSoftware = $false
            }
        }
        'Full' {
            return [pscustomobject]@{
                Level = 'Full'
                IncludeDashboard = $true
                IncludeSystem = $true
                IncludeDisks = $true
                IncludeHardware = $true
                IncludeUpdates = $true
                IncludeSoftware = $true
            }
        }
        default {
            return [pscustomobject]@{
                Level = 'Standard'
                IncludeDashboard = $true
                IncludeSystem = $true
                IncludeDisks = $true
                IncludeHardware = $true
                IncludeUpdates = $true
                IncludeSoftware = $false
            }
        }
    }
}

function Show-ReportOptionsDialog {
    param([string]$Title = 'Report Options')

    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = $Title
    $dialog.StartPosition = 'CenterParent'
    $dialog.FormBorderStyle = 'FixedDialog'
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false
    $dialog.ClientSize = New-Object System.Drawing.Size(560, 430)
    $dialog.MinimumSize = New-Object System.Drawing.Size(560, 430)
    $dialog.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $dialog.BackColor = [System.Drawing.Color]::White

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.RowCount = 5
    $layout.ColumnCount = 1
    $layout.Padding = New-Object System.Windows.Forms.Padding(14)
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 56)))
    $dialog.Controls.Add($layout)

    $layout.Controls.Add((New-Label -Text 'Choose report detail level' -Size 11 -Style Bold), 0, 0)

    $levelPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $levelPanel.Dock = 'Fill'
    $levelPanel.FlowDirection = 'TopDown'
    $levelPanel.WrapContents = $false
    $summaryRadio = New-Object System.Windows.Forms.RadioButton
    $summaryRadio.Text = 'Simple system summary'
    $summaryRadio.Width = 360
    $summaryRadio.Tag = 'Summary'
    $standardRadio = New-Object System.Windows.Forms.RadioButton
    $standardRadio.Text = 'Standard report'
    $standardRadio.Width = 360
    $standardRadio.Tag = 'Standard'
    $standardRadio.Checked = $true
    $fullRadio = New-Object System.Windows.Forms.RadioButton
    $fullRadio.Text = 'Full report'
    $fullRadio.Width = 360
    $fullRadio.Tag = 'Full'
    $levelPanel.Controls.Add($summaryRadio)
    $levelPanel.Controls.Add($standardRadio)
    $levelPanel.Controls.Add($fullRadio)
    $layout.Controls.Add($levelPanel, 0, 1)

    $layout.Controls.Add((New-Label -Text 'Sections' -Size 10 -Style Bold), 0, 2)
    $sectionsPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $sectionsPanel.Dock = 'Fill'
    $sectionsPanel.ColumnCount = 2
    $sectionsPanel.RowCount = 3
    [void]$sectionsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$sectionsPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$sectionsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
    [void]$sectionsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
    [void]$sectionsPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.34)))

    $sectionChecks = @{}
    $sectionIndex = 0
    foreach ($section in @(
        @{ Key = 'IncludeDashboard'; Text = 'Dashboard metrics'; Checked = $true },
        @{ Key = 'IncludeSystem'; Text = 'System/OS summary'; Checked = $true },
        @{ Key = 'IncludeDisks'; Text = 'Logical disks'; Checked = $true },
        @{ Key = 'IncludeHardware'; Text = 'Hardware'; Checked = $true },
        @{ Key = 'IncludeUpdates'; Text = 'Updates'; Checked = $true },
        @{ Key = 'IncludeSoftware'; Text = 'Installed software'; Checked = $false }
    )) {
        $check = New-Object System.Windows.Forms.CheckBox
        $check.Text = $section.Text
        $check.Dock = 'Fill'
        $check.Margin = New-Object System.Windows.Forms.Padding(4, 8, 4, 4)
        $check.Checked = $section.Checked
        $sectionChecks[$section.Key] = $check
        $sectionsPanel.Controls.Add($check, ($sectionIndex % 2), [Math]::Floor($sectionIndex / 2))
        $sectionIndex++
    }
    $layout.Controls.Add($sectionsPanel, 0, 3)

    $applyLevel = {
        $level = if ($summaryRadio.Checked) { 'Summary' } elseif ($fullRadio.Checked) { 'Full' } else { 'Standard' }
        $options = New-ReportOptions -Level $level
        foreach ($property in $options.PSObject.Properties) {
            if ($sectionChecks.ContainsKey($property.Name)) { $sectionChecks[$property.Name].Checked = [bool]$property.Value }
        }
    }
    $summaryRadio.add_CheckedChanged({ if ($summaryRadio.Checked) { & $applyLevel } })
    $standardRadio.add_CheckedChanged({ if ($standardRadio.Checked) { & $applyLevel } })
    $fullRadio.add_CheckedChanged({ if ($fullRadio.Checked) { & $applyLevel } })

    $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttons.Dock = 'Fill'
    $buttons.FlowDirection = 'RightToLeft'
    $okButton = New-ToolbarButton -Text 'OK'
    $cancelButton = New-ToolbarButton -Text 'Cancel'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $buttons.Controls.Add($okButton)
    $buttons.Controls.Add($cancelButton)
    $layout.Controls.Add($buttons, 0, 4)
    $dialog.AcceptButton = $okButton
    $dialog.CancelButton = $cancelButton

    try {
        if ($dialog.ShowDialog($Script:Controls.MainForm) -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
        $level = if ($summaryRadio.Checked) { 'Summary' } elseif ($fullRadio.Checked) { 'Full' } else { 'Standard' }
        $options = New-ReportOptions -Level $level
        foreach ($key in $sectionChecks.Keys) {
            $options.$key = [bool]$sectionChecks[$key].Checked
        }
        return $options
    } finally {
        $dialog.Dispose()
    }
}

function Get-ReportBrowserPath {
    $candidates = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) { return $candidate }
    }

    foreach ($name in @('msedge.exe', 'chrome.exe')) {
        $command = Get-Command -Name $name -ErrorAction SilentlyContinue
        if ($command) { return $command.Source }
    }

    return $null
}

function ConvertTo-HtmlCell {
    param([object]$Value)
    return [Net.WebUtility]::HtmlEncode([string](ConvertTo-DisplayText $Value ''))
}

function Add-HtmlTable {
    param(
        [Parameter(Mandatory)][System.Text.StringBuilder]$Builder,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][object[]]$Items,
        [string[]]$Properties
    )

    [void]$Builder.AppendLine(('<h2>{0}</h2>' -f (ConvertTo-HtmlCell $Title)))
    $rows = @($Items)
    if (-not $Properties -or $Properties.Count -eq 0) {
        if ($rows.Count -gt 0) {
            $Properties = @($rows[0].PSObject.Properties.Name | Where-Object { $_ -ne 'RawUpdate' })
        }
    }
    if ($rows.Count -eq 0 -or -not $Properties -or $Properties.Count -eq 0) {
        [void]$Builder.AppendLine('<p class="muted">No records found.</p>')
        return
    }

    [void]$Builder.AppendLine('<table><thead><tr>')
    foreach ($property in $Properties) { [void]$Builder.AppendLine(('<th>{0}</th>' -f (ConvertTo-HtmlCell $property))) }
    [void]$Builder.AppendLine('</tr></thead><tbody>')
    foreach ($row in $rows) {
        [void]$Builder.AppendLine('<tr>')
        foreach ($property in $Properties) {
            [void]$Builder.AppendLine(('<td>{0}</td>' -f (ConvertTo-HtmlCell $row.$property)))
        }
        [void]$Builder.AppendLine('</tr>')
    }
    [void]$Builder.AppendLine('</tbody></table>')
}

function ConvertTo-SystemReportHtml {
    param(
        [Parameter(Mandatory)][object]$ReportData,
        [object]$Options = (New-ReportOptions -Level 'Standard')
    )

    $snapshot = $ReportData.Snapshot
    $metrics = $ReportData.Metrics
    $uptime = if ($snapshot.Uptime) { '{0}d {1}h {2}m' -f $snapshot.Uptime.Days, $snapshot.Uptime.Hours, $snapshot.Uptime.Minutes } else { 'Unknown' }
    $admin = if ($snapshot.IsAdministrator) { 'Yes' } else { 'No' }
    $diskLabel = if ($metrics.Disk) { $metrics.Disk.DeviceId } else { 'System disk' }
    $diskDetail = if ($metrics.Disk) { '{0} free of {1}' -f (ConvertTo-SizeText $metrics.Disk.Free), (ConvertTo-SizeText $metrics.Disk.Size) } else { 'Unavailable' }
    $isSummaryReport = ([string]$Options.Level -eq 'Summary')
    $bodyClass = if ($isSummaryReport) { ' class="compact"' } else { '' }

    $html = New-Object System.Text.StringBuilder
    [void]$html.AppendLine('<!doctype html><html><head><meta charset="utf-8"><title>TEKSysInfo Printable Report</title>')
    [void]$html.AppendLine('<style>')
    [void]$html.AppendLine('body{font-family:Segoe UI,Arial,sans-serif;margin:28px;color:#17212b;background:#fff}header{border-bottom:3px solid #2563eb;padding-bottom:14px;margin-bottom:18px}h1{margin:0 0 4px 0;font-size:28px}h2{font-size:17px;border-bottom:1px solid #d8dee9;padding-bottom:6px;margin:24px 0 10px}.muted{color:#5d6b78}.summary{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin:16px 0 20px}.card{border:1px solid #d8dee9;border-radius:6px;padding:10px 12px;break-inside:avoid}.card .label{font-size:11px;text-transform:uppercase;color:#5d6b78;font-weight:700}.card .value{font-size:20px;font-weight:700;margin-top:4px}.card .detail{font-size:12px;color:#5d6b78;margin-top:2px}table{border-collapse:collapse;width:100%;font-size:12px;margin-bottom:12px;break-inside:auto}tr{break-inside:avoid}td,th{border:1px solid #d8dee9;padding:6px;text-align:left;vertical-align:top}th{background:#eef2f7;font-weight:700}.two-col{display:grid;grid-template-columns:1fr 1fr;gap:18px}.section{break-inside:avoid}.footer{margin-top:28px;font-size:11px;color:#5d6b78;border-top:1px solid #d8dee9;padding-top:8px}body.compact{margin:18px}body.compact header{padding-bottom:8px;margin-bottom:10px;border-bottom-width:2px}body.compact h1{font-size:22px}body.compact h2{font-size:14px;margin:12px 0 6px;padding-bottom:4px}body.compact .summary{gap:6px;margin:8px 0 10px}body.compact .card{padding:6px 8px;border-radius:4px}body.compact .card .label{font-size:9px}body.compact .card .value{font-size:16px;margin-top:2px}body.compact .card .detail{font-size:10px}body.compact table{font-size:10.5px;margin-bottom:6px}body.compact td,body.compact th{padding:3px 5px}body.compact .footer{margin-top:10px;font-size:9px;padding-top:5px}@media print{body{margin:0.35in}.compact{margin:0.22in}.no-print{display:none}.summary{grid-template-columns:repeat(4,1fr)}h2{page-break-after:avoid}.page-break{page-break-before:always}}@media(max-width:900px){.summary,.two-col{grid-template-columns:1fr}}')
    [void]$html.AppendLine("</style></head><body$bodyClass>")
    [void]$html.AppendLine(('<header><h1>TEKSysInfo System Summary</h1><div class="muted">{0} - generated {1}</div></header>' -f (ConvertTo-HtmlCell $snapshot.ComputerName), (ConvertTo-HtmlCell $ReportData.GeneratedAt)))
    if ($Options.IncludeDashboard) {
        [void]$html.AppendLine('<section class="summary">')
        foreach ($card in @(
            @{ Label = 'Processor Load'; Value = ConvertTo-PercentText $metrics.CpuPercent; Detail = 'Total processor load' },
            @{ Label = 'Memory Usage'; Value = ConvertTo-PercentText $metrics.MemoryPercent; Detail = ('{0} used / {1} total' -f (ConvertTo-SizeText $metrics.MemoryUsed), (ConvertTo-SizeText $metrics.MemoryTotal)) },
            @{ Label = 'Disk Usage'; Value = ConvertTo-PercentText $metrics.DiskPercent; Detail = "$diskDetail on $diskLabel" },
            @{ Label = 'Uptime'; Value = $uptime; Detail = ('Last boot {0}' -f $snapshot.LastBoot) }
        )) {
            [void]$html.AppendLine(('<div class="card"><div class="label">{0}</div><div class="value">{1}</div><div class="detail">{2}</div></div>' -f (ConvertTo-HtmlCell $card.Label), (ConvertTo-HtmlCell $card.Value), (ConvertTo-HtmlCell $card.Detail)))
        }
        [void]$html.AppendLine('</section>')
    }

    if ($Options.IncludeSystem -and $isSummaryReport) {
        [void]$html.AppendLine('<section class="section"><h2>System and OS</h2><table><tbody>')
        foreach ($row in @(
            @{ Name = 'Computer'; Value = $snapshot.ComputerName },
            @{ Name = 'User'; Value = $snapshot.UserName },
            @{ Name = 'Manufacturer / Model'; Value = ('{0} {1}' -f $snapshot.Manufacturer, $snapshot.Model) },
            @{ Name = 'Serial Number'; Value = $snapshot.SerialNumber },
            @{ Name = 'Processor'; Value = ('{0} ({1} cores / {2} threads)' -f $snapshot.Processor, $snapshot.ProcessorCores, $snapshot.ProcessorThreads) },
            @{ Name = 'Memory'; Value = $snapshot.TotalMemory },
            @{ Name = 'Windows'; Value = ('{0} - {1}' -f $snapshot.WindowsCaption, $snapshot.WindowsVersion) },
            @{ Name = 'Architecture'; Value = $snapshot.WindowsArchitecture },
            @{ Name = 'Last Boot / Uptime'; Value = ('{0} / {1}' -f $snapshot.LastBoot, $uptime) },
            @{ Name = 'Secure Boot / Admin'; Value = ('{0} / {1}' -f $snapshot.SecureBoot, $admin) }
        )) {
            [void]$html.AppendLine(('<tr><th>{0}</th><td>{1}</td></tr>' -f (ConvertTo-HtmlCell $row.Name), (ConvertTo-HtmlCell $row.Value)))
        }
        [void]$html.AppendLine('</tbody></table></section>')
    } elseif ($Options.IncludeSystem) {
        [void]$html.AppendLine('<section class="two-col">')
        [void]$html.AppendLine('<div class="section"><h2>System</h2><table><tbody>')
        foreach ($row in @(
            @{ Name = 'Computer'; Value = $snapshot.ComputerName },
            @{ Name = 'User'; Value = $snapshot.UserName },
            @{ Name = 'Manufacturer'; Value = $snapshot.Manufacturer },
            @{ Name = 'Model'; Value = $snapshot.Model },
            @{ Name = 'Serial Number'; Value = $snapshot.SerialNumber },
            @{ Name = 'BaseBoard'; Value = $snapshot.BaseBoard },
            @{ Name = 'BIOS'; Value = ('{0} ({1})' -f $snapshot.BiosVersion, $snapshot.BiosDate) },
            @{ Name = 'Administrator'; Value = $admin }
        )) {
            [void]$html.AppendLine(('<tr><th>{0}</th><td>{1}</td></tr>' -f (ConvertTo-HtmlCell $row.Name), (ConvertTo-HtmlCell $row.Value)))
        }
        [void]$html.AppendLine('</tbody></table></div>')

        [void]$html.AppendLine('<div class="section"><h2>Operating System</h2><table><tbody>')
        foreach ($row in @(
            @{ Name = 'Windows'; Value = $snapshot.WindowsCaption },
            @{ Name = 'Version'; Value = $snapshot.WindowsVersion },
            @{ Name = 'Architecture'; Value = $snapshot.WindowsArchitecture },
            @{ Name = 'Install Date'; Value = $snapshot.InstallDate },
            @{ Name = 'Last Boot'; Value = $snapshot.LastBoot },
            @{ Name = 'Secure Boot'; Value = $snapshot.SecureBoot }
        )) {
            [void]$html.AppendLine(('<tr><th>{0}</th><td>{1}</td></tr>' -f (ConvertTo-HtmlCell $row.Name), (ConvertTo-HtmlCell $row.Value)))
        }
        [void]$html.AppendLine('</tbody></table></div></section>')
    }

    if ($Options.IncludeDisks) { Add-HtmlTable -Builder $html -Title 'Logical Disks' -Items $ReportData.Disks -Properties @('Drive','Label','FileSystem','Size','Used','Free','UsedPercent') }
    if ($Options.IncludeHardware) { Add-HtmlTable -Builder $html -Title 'Hardware Inventory' -Items $ReportData.Hardware -Properties @('Category','Name','Manufacturer','Status','Detail') }
    if ($Options.IncludeUpdates) { Add-HtmlTable -Builder $html -Title 'Recent Installed Updates' -Items $ReportData.HotFixes -Properties @('HotFixId','Description','InstalledOn','InstalledBy') }
    if ($Options.IncludeSoftware) { Add-HtmlTable -Builder $html -Title 'Installed Software' -Items $ReportData.Software -Properties @('Name','Version','Publisher','InstallDate','Source') }

    [void]$html.AppendLine(('<div class="footer">Generated by TEKSysInfo on {0}.</div>' -f (ConvertTo-HtmlCell $ReportData.GeneratedAt)))
    [void]$html.AppendLine('</body></html>')
    return $html.ToString()
}

function Export-SystemReportCsv {
    param(
        [Parameter(Mandatory)][object]$ReportData,
        [Parameter(Mandatory)][string]$Path,
        [object]$Options = (New-ReportOptions -Level 'Standard')
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $add = {
        param([string]$Section, [string]$Item, [string]$Property, [object]$Value)
        $rows.Add([pscustomobject]@{
            Section = $Section
            Item = $Item
            Property = $Property
            Value = [string](ConvertTo-DisplayText $Value '')
        })
    }

    foreach ($property in $ReportData.Snapshot.PSObject.Properties) {
        if ($property.Name -ne 'Uptime') { & $add 'System Summary' $ReportData.Snapshot.ComputerName $property.Name $property.Value }
    }
    foreach ($property in $ReportData.Metrics.PSObject.Properties) {
        if ($property.Name -ne 'Disk') { & $add 'Live Metrics' $ReportData.Snapshot.ComputerName $property.Name $property.Value }
    }
    $sections = New-Object System.Collections.Generic.List[object]
    if ($Options.IncludeDisks) { $sections.Add(@{ Name = 'Logical Disks'; Items = $ReportData.Disks; ItemProperty = 'Drive' }) }
    if ($Options.IncludeUpdates) { $sections.Add(@{ Name = 'Recent Installed Updates'; Items = $ReportData.HotFixes; ItemProperty = 'HotFixId' }) }
    if ($Options.IncludeSoftware) { $sections.Add(@{ Name = 'Installed Software'; Items = $ReportData.Software; ItemProperty = 'Name' }) }
    if ($Options.IncludeHardware) { $sections.Add(@{ Name = 'Hardware Inventory'; Items = $ReportData.Hardware; ItemProperty = 'Name' }) }

    foreach ($section in $sections) {
        foreach ($item in @($section.Items)) {
            $itemName = [string]$item.($section.ItemProperty)
            foreach ($property in $item.PSObject.Properties) {
                if ($property.Name -ne 'RawUpdate') { & $add $section.Name $itemName $property.Name $property.Value }
            }
        }
    }

    $rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

function Export-SystemReportPdf {
    param(
        [Parameter(Mandatory)][string]$Html,
        [Parameter(Mandatory)][string]$Path
    )

    $browser = Get-ReportBrowserPath
    if (-not $browser) {
        throw 'Microsoft Edge or Google Chrome is required for PDF export. Export HTML instead, then print to PDF from your browser.'
    }

    $tempHtml = Join-Path $env:TEMP ('TEKSysInfo_Print_{0}.html' -f ([Guid]::NewGuid().ToString('N')))
    [IO.File]::WriteAllText($tempHtml, $Html, [Text.Encoding]::UTF8)
    try {
        $arguments = @(
            '--headless',
            '--disable-gpu',
            '--no-pdf-header-footer',
            ('--print-to-pdf="{0}"' -f $Path),
            ('"file:///{0}"' -f ($tempHtml -replace '\\','/'))
        )
        $process = Start-Process -FilePath $browser -ArgumentList $arguments -Wait -PassThru -WindowStyle Hidden
        if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $Path)) {
            throw "PDF export failed with browser exit code $($process.ExitCode)."
        }
    } finally {
        Remove-Item -LiteralPath $tempHtml -Force -ErrorAction SilentlyContinue
    }
}

function Export-SystemReport {
    $options = Show-ReportOptionsDialog -Title 'Export Report Options'
    if (-not $options) { return }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = 'Export TEKSysInfo report'
    $dialog.Filter = 'PDF report (*.pdf)|*.pdf|Printable HTML report (*.html)|*.html|CSV data (*.csv)|*.csv'
    $dialog.FileName = ('TEKSysInfo_{0}_{1}.pdf' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        $dialog.Dispose()
        return
    }

    $targetPath = $dialog.FileName
    $dialog.Dispose()

    Set-Status 'Exporting report...'
    if ($Script:Controls.ContainsKey('ExportButton')) { $Script:Controls.ExportButton.Enabled = $false }
    Invoke-Background -Work {
        param($path, $reportOptions)
        $report = Get-SystemReportData
        $extension = [IO.Path]::GetExtension($path).ToLowerInvariant()
        if ($extension -eq '.csv') {
            Export-SystemReportCsv -ReportData $report -Path $path -Options $reportOptions
        } else {
            $html = ConvertTo-SystemReportHtml -ReportData $report -Options $reportOptions
            if ($extension -eq '.pdf') {
                Export-SystemReportPdf -Html $html -Path $path
            } else {
                [IO.File]::WriteAllText($path, $html, [Text.Encoding]::UTF8)
            }
        }
        return $path
    } -Done {
        param($result)
        $path = @($result) | Select-Object -First 1
        if ($Script:Controls.ContainsKey('ExportButton')) { $Script:Controls.ExportButton.Enabled = $true }
        Set-Status "Report exported to $path"
        [System.Windows.Forms.MessageBox]::Show("Report exported:`r`n$path", $Script:AppName, 'OK', 'Information') | Out-Null
    } -ArgumentList @($targetPath, $options) -Failed {
        param($errorRecord)
        if ($Script:Controls.ContainsKey('ExportButton')) { $Script:Controls.ExportButton.Enabled = $true }
        Set-Status "Report export failed: $($errorRecord.Message)"
        [System.Windows.Forms.MessageBox]::Show($errorRecord.Message, $Script:AppName, 'OK', 'Error') | Out-Null
    }
}

function Open-PrintableSystemReport {
    $options = Show-ReportOptionsDialog -Title 'Print Report Options'
    if (-not $options) { return }

    Set-Status 'Preparing printable report...'
    if ($Script:Controls.ContainsKey('PrintButton')) { $Script:Controls.PrintButton.Enabled = $false }
    Invoke-Background -Work {
        param($reportOptions)
        $report = Get-SystemReportData
        $html = ConvertTo-SystemReportHtml -ReportData $report -Options $reportOptions
        $path = Join-Path $env:TEMP ('TEKSysInfo_Printable_{0}_{1}.html' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd_HHmmss'))
        [IO.File]::WriteAllText($path, $html, [Text.Encoding]::UTF8)
        return $path
    } -Done {
        param($result)
        $path = @($result) | Select-Object -First 1
        if ($Script:Controls.ContainsKey('PrintButton')) { $Script:Controls.PrintButton.Enabled = $true }
        Start-Process $path
        Set-Status "Printable report opened: $path"
    } -ArgumentList @($options) -Failed {
        param($errorRecord)
        if ($Script:Controls.ContainsKey('PrintButton')) { $Script:Controls.PrintButton.Enabled = $true }
        Set-Status "Printable report failed: $($errorRecord.Message)"
        [System.Windows.Forms.MessageBox]::Show($errorRecord.Message, $Script:AppName, 'OK', 'Error') | Out-Null
    }
}

function New-ToolbarButton {
    param([string]$Text)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.AutoSize = $true
    $button.Height = 32
    $button.Margin = New-Object System.Windows.Forms.Padding(4)
    $button.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    return $button
}

function Build-Interface {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "$Script:AppName - System Information Console"
    $form.StartPosition = 'CenterScreen'
    $form.MinimumSize = New-Object System.Drawing.Size(1120, 720)
    $form.Size = New-Object System.Drawing.Size(1320, 840)
    $form.BackColor = [System.Drawing.Color]::FromArgb(241, 245, 249)
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $Script:Controls.MainForm = $form

    $main = New-Object System.Windows.Forms.TableLayoutPanel
    $main.Dock = 'Fill'
    $main.RowCount = 3
    $main.ColumnCount = 1
    $main.Padding = New-Object System.Windows.Forms.Padding(14)
    [void]$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 88)))
    [void]$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26)))
    $form.Controls.Add($main)

    $header = New-Object System.Windows.Forms.TableLayoutPanel
    $header.Dock = 'Fill'
    $header.ColumnCount = 2
    [void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))

    $titlePanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $titlePanel.FlowDirection = 'TopDown'
    $titlePanel.Dock = 'Fill'
    $titlePanel.WrapContents = $false
    $titlePanel.Padding = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
    $title = New-Label -Text 'TEKSysInfo' -Size 22 -Style Bold -Color ([System.Drawing.Color]::FromArgb(15, 23, 42))
    $title.Margin = New-Object System.Windows.Forms.Padding(8, 2, 8, 0)
    $titlePanel.Controls.Add($title)
    $subtitle = New-Label -Text 'Live system health, inventory, and Windows Update console' -Size 10 -Color ([System.Drawing.Color]::FromArgb(71, 85, 105))
    $subtitle.Margin = New-Object System.Windows.Forms.Padding(8, 2, 8, 4)
    $titlePanel.Controls.Add($subtitle)

    $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttons.FlowDirection = 'LeftToRight'
    $buttons.Dock = 'Fill'
    $buttons.AutoSize = $false
    $buttons.Width = 620
    $buttons.Height = 42
    $buttons.MinimumSize = New-Object System.Drawing.Size(620, 42)
    $buttons.WrapContents = $false
    $buttons.Anchor = 'Right'

    $refreshButton = New-ToolbarButton -Text 'Refresh Inventory'
    $exportButton = New-ToolbarButton -Text 'Export Report'
    $printButton = New-ToolbarButton -Text 'Print Report'
    $settingsButton = New-ToolbarButton -Text 'Windows Update Settings'
    $refreshButton.add_Click({ Refresh-Inventory })
    $exportButton.add_Click({ Export-SystemReport })
    $printButton.add_Click({ Open-PrintableSystemReport })
    $settingsButton.add_Click({ Start-Process 'ms-settings:windowsupdate' })
    $Script:Controls.RefreshButton = $refreshButton
    $Script:Controls.ExportButton = $exportButton
    $Script:Controls.PrintButton = $printButton
    $buttons.Controls.Add($refreshButton)
    $buttons.Controls.Add($exportButton)
    $buttons.Controls.Add($printButton)
    $buttons.Controls.Add($settingsButton)

    $header.Controls.Add($titlePanel, 0, 0)
    $header.Controls.Add($buttons, 1, 0)
    $main.Controls.Add($header, 0, 0)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = 'Fill'
    $tabs.Font = New-Object System.Drawing.Font('Segoe UI', 9)
    $tabs.Alignment = 'Top'
    $main.Controls.Add($tabs, 0, 1)

    $dashboardTab = New-Object System.Windows.Forms.TabPage
    $dashboardTab.Text = 'Dashboard'
    $dashboardTab.BackColor = [System.Drawing.Color]::White
    $dashboardLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $dashboardLayout.Dock = 'Fill'
    $dashboardLayout.RowCount = 3
    $dashboardLayout.ColumnCount = 1
    $dashboardLayout.Padding = New-Object System.Windows.Forms.Padding(8)
    [void]$dashboardLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 300)))
    [void]$dashboardLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36)))
    [void]$dashboardLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

    $dashboardTiles = New-Object System.Windows.Forms.TableLayoutPanel
    $dashboardTiles.Dock = 'Fill'
    $dashboardTiles.ColumnCount = 4
    $dashboardTiles.RowCount = 3
    [void]$dashboardTiles.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$dashboardTiles.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$dashboardTiles.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$dashboardTiles.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
    [void]$dashboardTiles.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
    [void]$dashboardTiles.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.33)))
    [void]$dashboardTiles.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33.34)))
    $dashboardTiles.Controls.Add((New-DashboardTile -Key 'Cpu' -Title 'Processor Load' -Accent ([System.Drawing.Color]::FromArgb(37, 99, 235))), 0, 0)
    $dashboardTiles.Controls.Add((New-DashboardTile -Key 'Memory' -Title 'Memory Usage' -Accent ([System.Drawing.Color]::FromArgb(5, 150, 105))), 1, 0)
    $dashboardTiles.Controls.Add((New-DashboardTile -Key 'Disk' -Title 'System Disk' -Accent ([System.Drawing.Color]::FromArgb(217, 119, 6))), 2, 0)
    $dashboardTiles.Controls.Add((New-DashboardTile -Key 'Thermal' -Title 'Temperature' -Accent ([System.Drawing.Color]::FromArgb(220, 38, 38))), 3, 0)
    $dashboardTiles.Controls.Add((New-DashboardTile -Key 'Network' -Title 'Network Throughput' -Accent ([System.Drawing.Color]::FromArgb(8, 145, 178))), 0, 1)
    $dashboardTiles.Controls.Add((New-DashboardTile -Key 'DiskIo' -Title 'Disk Activity' -Accent ([System.Drawing.Color]::FromArgb(124, 58, 237))), 1, 1)
    $dashboardTiles.Controls.Add((New-DashboardTile -Key 'TopCpu' -Title 'Top CPU Process' -Accent ([System.Drawing.Color]::FromArgb(37, 99, 235))), 2, 1)
    $dashboardTiles.Controls.Add((New-DashboardTile -Key 'TopMem' -Title 'Top Memory Process' -Accent ([System.Drawing.Color]::FromArgb(5, 150, 105))), 3, 1)
    $dashboardTiles.Controls.Add((New-DashboardTile -Key 'Reboot' -Title 'Pending Reboot' -Accent ([System.Drawing.Color]::FromArgb(217, 119, 6))), 0, 2)
    $dashboardTiles.Controls.Add((New-DashboardTile -Key 'DriveHealth' -Title 'Drive Health' -Accent ([System.Drawing.Color]::FromArgb(22, 163, 74))), 1, 2)
    $dashboardTiles.Controls.Add((New-DashboardTile -Key 'Power' -Title 'Power' -Accent ([System.Drawing.Color]::FromArgb(79, 70, 229))), 2, 2)
    $dashboardTiles.Controls.Add((New-DashboardTile -Key 'Security' -Title 'Security' -Accent ([System.Drawing.Color]::FromArgb(15, 118, 110))), 3, 2)
    $dashboardLayout.Controls.Add($dashboardTiles, 0, 0)

    $dashboardLayout.Controls.Add((New-Label -Text 'System Overview and Live Thermals' -Size 11 -Style Bold), 0, 1)
    $dashboardSplit = New-Object System.Windows.Forms.SplitContainer
    $dashboardSplit.Dock = 'Fill'
    $dashboardSplit.Orientation = 'Vertical'
    $dashboardSplit.SplitterDistance = 560
    $overviewPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $overviewPanel.Dock = 'Fill'
    $overviewPanel.RowCount = 2
    [void]$overviewPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 32)))
    [void]$overviewPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $overviewPanel.Controls.Add((New-Label -Text 'Overview' -Size 10 -Style Bold), 0, 0)
    $overviewList = New-ListView -Columns @('Area','Status','Detail')
    $Script:Controls.DashboardOverviewList = $overviewList
    $overviewPanel.Controls.Add($overviewList, 0, 1)
    $livePanel = New-Object System.Windows.Forms.TableLayoutPanel
    $livePanel.Dock = 'Fill'
    $livePanel.RowCount = 6
    [void]$livePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
    [void]$livePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 34)))
    [void]$livePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
    [void]$livePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    [void]$livePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30)))
    [void]$livePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 33)))
    $livePanel.Controls.Add((New-Label -Text 'Thermal Readings' -Size 10 -Style Bold), 0, 0)
    $thermalList = New-ListView -Columns @('Sensor','Temperature','Source','Detail')
    $Script:Controls.ThermalList = $thermalList
    $livePanel.Controls.Add($thermalList, 0, 1)
    $livePanel.Controls.Add((New-Label -Text 'Top Processes' -Size 10 -Style Bold), 0, 2)
    $processList = New-ListView -Columns @('Metric','Process','Value')
    $Script:Controls.ProcessList = $processList
    $livePanel.Controls.Add($processList, 0, 3)
    $livePanel.Controls.Add((New-Label -Text 'Network and Disk Activity' -Size 10 -Style Bold), 0, 4)
    $ioList = New-ListView -Columns @('Metric','Value','Detail')
    $Script:Controls.IoList = $ioList
    $livePanel.Controls.Add($ioList, 0, 5)
    $dashboardSplit.Panel1.Controls.Add($overviewPanel)
    $dashboardSplit.Panel2.Controls.Add($livePanel)
    $dashboardLayout.Controls.Add($dashboardSplit, 0, 2)
    $dashboardTab.Controls.Add($dashboardLayout)
    $tabs.TabPages.Add($dashboardTab)

    $disksTab = New-Object System.Windows.Forms.TabPage
    $disksTab.Text = 'Fixed Disks'
    $disksTab.BackColor = [System.Drawing.Color]::White
    $diskLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $diskLayout.Dock = 'Fill'
    $diskLayout.RowCount = 2
    [void]$diskLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))
    [void]$diskLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $diskLayout.Controls.Add((New-Label -Text 'Fixed Disks' -Size 11 -Style Bold), 0, 0)
    $diskList = New-ListView -Columns @('Drive','Label','File System','Size','Used','Free','Used %')
    $diskMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $diskDetailMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $diskDetailMenuItem.Text = 'View Disk Health and Details'
    $diskDetailMenuItem.add_Click({ Show-SelectedDiskDetails })
    [void]$diskMenu.Items.Add($diskDetailMenuItem)
    $diskList.ContextMenuStrip = $diskMenu
    $diskList.add_MouseDoubleClick({ Show-SelectedDiskDetails })
    $Script:Controls.DiskList = $diskList
    $diskLayout.Controls.Add($diskList, 0, 1)
    $disksTab.Controls.Add($diskLayout)
    $tabs.TabPages.Add($disksTab)

    $detailsTab = New-Object System.Windows.Forms.TabPage
    $detailsTab.Text = 'System Details'
    $detailsTab.BackColor = [System.Drawing.Color]::White
    $detailsList = New-ListView -Columns @('Property','Value')
    $Script:Controls.DetailsList = $detailsList
    $detailsTab.Controls.Add($detailsList)
    $tabs.TabPages.Add($detailsTab)

    $softwareTab = New-Object System.Windows.Forms.TabPage
    $softwareTab.Text = 'Installed Software'
    $softwareTab.BackColor = [System.Drawing.Color]::White
    $softwareLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $softwareLayout.Dock = 'Fill'
    $softwareLayout.RowCount = 2
    [void]$softwareLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))
    [void]$softwareLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $softwareToolbar = New-Object System.Windows.Forms.FlowLayoutPanel
    $softwareToolbar.Dock = 'Fill'
    $softwareToolbar.FlowDirection = 'LeftToRight'
    $softwareToolbar.Controls.Add((New-Label -Text 'Search' -Size 9 -Style Bold))
    $softwareSearch = New-Object System.Windows.Forms.TextBox
    $softwareSearch.Width = 320
    $softwareSearch.Margin = New-Object System.Windows.Forms.Padding(4, 8, 16, 4)
    $softwareCount = New-Label -Text '0 programs'
    $softwareSearch.add_TextChanged({ Apply-SoftwareFilter })
    $Script:Controls.SoftwareSearch = $softwareSearch
    $Script:Controls.SoftwareCountLabel = $softwareCount
    $softwareToolbar.Controls.Add($softwareSearch)
    $softwareToolbar.Controls.Add($softwareCount)
    $softwareList = New-ListView -Columns @('Name','Version','Publisher','Install Date','Source')
    $Script:Controls.SoftwareList = $softwareList
    $softwareLayout.Controls.Add($softwareToolbar, 0, 0)
    $softwareLayout.Controls.Add($softwareList, 0, 1)
    $softwareTab.Controls.Add($softwareLayout)
    $tabs.TabPages.Add($softwareTab)

    $hardwareTab = New-Object System.Windows.Forms.TabPage
    $hardwareTab.Text = 'Hardware'
    $hardwareTab.BackColor = [System.Drawing.Color]::White
    $hardwareLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $hardwareLayout.Dock = 'Fill'
    $hardwareLayout.RowCount = 2
    [void]$hardwareLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))
    [void]$hardwareLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $hardwareToolbar = New-Object System.Windows.Forms.FlowLayoutPanel
    $hardwareToolbar.Dock = 'Fill'
    $hardwareToolbar.FlowDirection = 'LeftToRight'
    $hardwareToolbar.Controls.Add((New-Label -Text 'Search' -Size 9 -Style Bold))
    $hardwareSearch = New-Object System.Windows.Forms.TextBox
    $hardwareSearch.Width = 320
    $hardwareSearch.Margin = New-Object System.Windows.Forms.Padding(4, 8, 16, 4)
    $hardwareCount = New-Label -Text '0 devices'
    $hardwareSearch.add_TextChanged({ Apply-HardwareFilter })
    $Script:Controls.HardwareSearch = $hardwareSearch
    $Script:Controls.HardwareCountLabel = $hardwareCount
    $hardwareToolbar.Controls.Add($hardwareSearch)
    $hardwareToolbar.Controls.Add($hardwareCount)
    $hardwareList = New-ListView -Columns @('Category','Name','Manufacturer','Status','Detail')
    $Script:Controls.HardwareList = $hardwareList
    $hardwareLayout.Controls.Add($hardwareToolbar, 0, 0)
    $hardwareLayout.Controls.Add($hardwareList, 0, 1)
    $hardwareTab.Controls.Add($hardwareLayout)
    $tabs.TabPages.Add($hardwareTab)

    $updatesTab = New-Object System.Windows.Forms.TabPage
    $updatesTab.Text = 'Windows Update'
    $updatesTab.BackColor = [System.Drawing.Color]::White
    $updatesSplit = New-Object System.Windows.Forms.SplitContainer
    $updatesSplit.Dock = 'Fill'
    $updatesSplit.Orientation = 'Horizontal'
    $updatesSplit.SplitterDistance = 230

    $availablePanel = New-Object System.Windows.Forms.TableLayoutPanel
    $availablePanel.Dock = 'Fill'
    $availablePanel.RowCount = 2
    [void]$availablePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))
    [void]$availablePanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $updateToolbar = New-Object System.Windows.Forms.FlowLayoutPanel
    $updateToolbar.Dock = 'Fill'
    $scanButton = New-ToolbarButton -Text 'Scan Available Updates'
    $installButton = New-ToolbarButton -Text 'Install Selected / All'
    $installButton.Enabled = $false
    $updateSummary = New-Label -Text 'Scan has not run yet.'
    $scanButton.add_Click({ Start-UpdateScan })
    $installButton.add_Click({ Start-UpdateInstall })
    $Script:Controls.ScanUpdatesButton = $scanButton
    $Script:Controls.InstallUpdatesButton = $installButton
    $Script:Controls.UpdateSummaryLabel = $updateSummary
    $updateToolbar.Controls.Add($scanButton)
    $updateToolbar.Controls.Add($installButton)
    $updateToolbar.Controls.Add($updateSummary)
    $availableList = New-ListView -Columns @('Title','KB','Severity','Size','Reboot')
    $availableList.MultiSelect = $true
    $Script:Controls.AvailableUpdateList = $availableList
    $availablePanel.Controls.Add($updateToolbar, 0, 0)
    $availablePanel.Controls.Add($availableList, 0, 1)
    $updatesSplit.Panel1.Controls.Add($availablePanel)

    $hotfixPanel = New-Object System.Windows.Forms.TableLayoutPanel
    $hotfixPanel.Dock = 'Fill'
    $hotfixPanel.RowCount = 2
    [void]$hotfixPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 36)))
    [void]$hotfixPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $hotfixPanel.Controls.Add((New-Label -Text 'Recent Installed Updates' -Size 10 -Style Bold), 0, 0)
    $hotfixList = New-ListView -Columns @('HotFix ID','Description','Installed On','Installed By')
    $Script:Controls.HotFixList = $hotfixList
    $hotfixPanel.Controls.Add($hotfixList, 0, 1)
    $updatesSplit.Panel2.Controls.Add($hotfixPanel)
    $updatesTab.Controls.Add($updatesSplit)
    $tabs.TabPages.Add($updatesTab)

    $status = New-Object System.Windows.Forms.StatusStrip
    $statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
    $statusLabel.Text = 'Ready'
    $statusLabel.Spring = $true
    $statusLabel.TextAlign = 'MiddleLeft'
    [void]$status.Items.Add($statusLabel)
    $Script:Controls.StatusLabel = $statusLabel
    $main.Controls.Add($status, 0, 2)

    $form.add_Shown({
        Start-LiveDashboardRefresh
        Start-ExtendedDashboardRefresh
        $Script:StartupInventoryTimer = New-Object System.Windows.Forms.Timer
        $Script:StartupInventoryTimer.Interval = 750
        $Script:StartupInventoryTimer.add_Tick({
            param($sender, $eventArgs)
            $sender.Stop()
            $sender.Dispose()
            $Script:StartupInventoryTimer = $null
            Refresh-Inventory
        })
        $Script:StartupInventoryTimer.Start()
        $Script:LiveTimer = New-Object System.Windows.Forms.Timer
        $Script:LiveTimer.Interval = 2000
        $Script:LiveTimer.add_Tick({ Start-LiveDashboardRefresh })
        $Script:LiveTimer.Start()
        $Script:ExtendedLiveTimer = New-Object System.Windows.Forms.Timer
        $Script:ExtendedLiveTimer.Interval = 10000
        $Script:ExtendedLiveTimer.add_Tick({ Start-ExtendedDashboardRefresh })
        $Script:ExtendedLiveTimer.Start()
    })

    $form.add_FormClosing({
        $Script:IsClosing = $true
    })

    $form.add_FormClosed({
        $Script:IsClosing = $true
        if ($Script:LiveTimer) {
            $Script:LiveTimer.Stop()
            $Script:LiveTimer.Dispose()
        }
        if ($Script:ExtendedLiveTimer) {
            $Script:ExtendedLiveTimer.Stop()
            $Script:ExtendedLiveTimer.Dispose()
        }
        if ($Script:StartupInventoryTimer) {
            $Script:StartupInventoryTimer.Stop()
            $Script:StartupInventoryTimer.Dispose()
        }
    })

    return $form
}

if ($SelfTest) {
    $snapshot = Get-SystemSnapshot
    $metrics = Get-LiveMetrics
    $disks = @(Get-LogicalDisks)
    $software = @(Get-InstalledSoftware | Select-Object -First 5)
    $hardware = @(Get-HardwareInventory | Select-Object -First 5)
    $hotfixes = @(Get-RecentHotFixes | Select-Object -First 5)

    [pscustomobject]@{
        ComputerName = $snapshot.ComputerName
        Windows = $snapshot.WindowsVersion
        CpuPercent = $metrics.CpuPercent
        MemoryPercent = $metrics.MemoryPercent
        DiskCount = $disks.Count
        SoftwareSampleCount = $software.Count
        HardwareSampleCount = $hardware.Count
        HotFixSampleCount = $hotfixes.Count
    } | Format-List
    return
}

if ($SelfTestDiskDetails) {
    $disk = if ($SelfTestDrive) {
        Get-LogicalDisks | Where-Object { $_.Drive -eq $SelfTestDrive.TrimEnd('\') } | Select-Object -First 1
    } else {
        Get-LogicalDisks | Select-Object -First 1
    }
    if (-not $disk) {
        throw 'No fixed disks were found for disk detail self-test.'
    }
    $details = @(Get-DiskDetail -DriveLetter $disk.Drive)
    [pscustomobject]@{
        Drive = $disk.Drive
        DetailRows = $details.Count
        FirstRows = ($details | Select-Object -First 8 | ForEach-Object { '{0} / {1} = {2}' -f $_.Category, $_.Property, $_.Value }) -join "`n"
    } | Format-List
    return
}

try {
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
    $form = Build-Interface
    [void][System.Windows.Forms.Application]::Run($form)
} catch {
    Write-Log "Fatal error: $($_.Exception.Message)"
    [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, $Script:AppName, 'OK', 'Error') | Out-Null
    throw
}
