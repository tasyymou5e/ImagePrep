# ImagePrep - VDI Image Optimization Tool
# Combines best features from imageprep2.0 and WDOT

[CmdletBinding(DefaultParameterSetName="Interactive")]
Param (
    [Parameter(ParameterSetName="Cmdlets")]
    [ArgumentCompleter( { Get-ChildItem $PSScriptRoot -Directory | Select-Object -ExpandProperty Name } )]
    [System.String]$WindowsVersion = (Get-ItemProperty "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\").ReleaseId,

    [Parameter(ParameterSetName="Cmdlets")]
    [ValidateSet('All','Safe','Medium','Full','Services','ScheduledTasks','AppxPackages','DefaultUser','Autologgers','Network','Policy','DiskCleanup','Edge','Privacy','WindowsUpdate','UIPersonalization','ServicesStartType','SoundsNotifications','ExplorerAutoplay')]
    [String[]]$Optimizations,

    [Parameter(ParameterSetName="Cmdlets")]
    [Switch]$DryRun,

    [Parameter(ParameterSetName="Cmdlets")]
    [Switch]$Verbose,

    [Parameter(ParameterSetName="Cmdlets")]
    [Switch]$Restart,

    [Parameter(ParameterSetName="Cmdlets")]
    [Switch]$AcceptEULA
)

$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$ConfigPath = Join-Path $ScriptDir $WindowsVersion
$JsonFolder = Join-Path $ConfigPath "ConfigurationFiles"

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole("Administrator")) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force

$LogPath = Join-Path (Get-Location) "ImagePrep_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
function Write-Log { param([string]$Msg, [string]$Level="INFO")
    "$((Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) [$Level] $Msg" | Add-Content $LogPath
    if ($Verbose -or $VerboseLogging) { Write-Host "$Level`: $Msg" }
}

function Load-JsonConfig {
    param([string]$FileName)
    $path = Join-Path $JsonFolder $FileName
    if (Test-Path $path) {
        try { return Get-Content $path -Raw | ConvertFrom-Json -EA Stop } catch {}
    }
    return $null
}

$Script:Configs = @{
    Services           = Load-JsonConfig "Services.json"
    ScheduledTasks     = Load-JsonConfig "ScheduledTasks.json"
    AppxPackages      = Load-JsonConfig "AppxPackages.json"
    Autologgers       = Load-JsonConfig "Autologgers.Json"
    LanManWorkstation = Load-JsonConfig "LanManWorkstation.json"
    PolicyRegSettings = Load-JsonConfig "PolicyRegSettings.json"
    DefaultUser       = Load-JsonConfig "DefaultUserSettings.json"
    EdgeSettings      = Load-JsonConfig "EdgeSettings.json"
    PrivacyTelemetry  = Load-JsonConfig "PrivacyTelemetry.json"
    WindowsUpdate     = Load-JsonConfig "WindowsUpdate.json"
    UIPersonalization = Load-JsonConfig "UIPersonalization.json"
    ServicesStartType = Load-JsonConfig "ServicesStartType.json"
    SoundsNotifications = Load-JsonConfig "SoundsNotifications.json"
    ExplorerAutoplay  = Load-JsonConfig "ExplorerAutoplay.json"
}

function Apply-RegistrySettings {
    param([array]$Settings, [string]$CategoryName)
    if (-not $Settings) { return }
    $count = @($Settings).Count
    Write-Log "Processing $count $CategoryName registry settings" "INFO"
    foreach ($item in $Settings) {
        try {
            if (-not $Script:DryRun) {
                if (Get-ItemProperty -Path $item.Path -Name $item.Name -EA SilentlyContinue) {
                    Set-ItemProperty -Path $item.Path -Name $item.Name -Value $item.Value -Type $item.Type -Force -EA SilentlyContinue
                } else {
                    if (Test-Path $item.Path) {
                        New-ItemProperty -Path $item.Path -Name $item.Name -PropertyType $item.Type -Value $item.Value -Force -EA SilentlyContinue | Out-Null
                    } else {
                        New-Item -Path $item.Path -Force -EA SilentlyContinue | Out-Null
                        New-ItemProperty -Path $item.Path -Name $item.Name -PropertyType $item.Type -Value $item.Value -Force -EA SilentlyContinue | Out-Null
                    }
                }
            }
            Write-Log "Registry: $($item.Name)" "INFO"
        } catch { Write-Log "Failed: $($item.Name)" "WARN" }
    }
}

function Apply-ServiceStartType {
    param([array]$Services, [string]$CategoryName)
    if (-not $Services) { return }
    $count = @($Services).Count
    Write-Log "Processing $count service start types" "INFO"
    foreach ($svc in $Services) {
        try {
            if (-not $Script:DryRun) {
                Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\$($svc.Name)" -Name "Start" -Value $svc.StartType -Type DWord -Force -EA SilentlyContinue
            }
            Write-Log "Service: $($svc.Name) -> Start=$($svc.StartType)" "INFO"
        } catch { Write-Log "Failed: $($svc.Name)" "WARN" }
    }
}

if ($PSCmdlet.ParameterSetName -eq "Interactive") {
    Clear-Host
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "     ImagePrep - VDI Optimization Tool" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. Safe Mode       -> Scheduled Tasks" -ForegroundColor Yellow
    Write-Host "2. Medium Mode     -> Safe + Services + Autologgers" -ForegroundColor Yellow
    Write-Host "3. Full Mode       -> Everything" -ForegroundColor Yellow
    Write-Host "4. Registry Only   -> Privacy + WindowsUpdate + UI + Sounds + Explorer" -ForegroundColor Yellow
    Write-Host "5. Custom          -> Select categories" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "6. Dry Run Preview -> See what would be changed" -ForegroundColor Magenta
    Write-Host "Q. Quit" -ForegroundColor Gray
    Write-Host ""

    $Script:DryRun = (Read-Host "Dry run mode? (Y/N)") -match '^Y$'
    $Script:VerboseLogging = (Read-Host "Verbose output? (Y/N)") -match '^Y$'

    $choice = Read-Host "Select (1-6, Q)"
    if ($choice -eq 'Q' -or $choice -eq 'q') { exit }

    $Script:RunLevel = switch ($choice) {
        "1" { "Safe" }
        "2" { "Medium" }
        "3" { "Full" }
        "4" { "Registry" }
        "5" { "Custom" }
        "6" { 
            Write-Log "=== DRY RUN PREVIEW MODE ===" "INFO"
            $Script:DryRun = $true
            $Script:VerboseLogging = $true
            "Full" 
        }
        default { exit }
    }

    if ($Script:RunLevel -eq "Custom") {
        Write-Host ""
        Write-Host "Available categories:" -ForegroundColor Cyan
        Write-Host "S=Services, T=ScheduledTasks, A=AppxPackages, U=DefaultUser"
        Write-Host "L=Autologgers, N=Network, P=Policy, D=DiskCleanup, E=Edge"
        Write-Host "R=Privacy, W=WindowsUpdate, I=UI Personalization"
        Write-Host "O=Sounds/Notifications, X=Explorer/Autoplay"
        $cats = Read-Host "Enter letters (e.g. STRLA)"
        $Script:CustomCategories = @()
        if ($cats -match 'S') { $Script:CustomCategories += "Services" }
        if ($cats -match 'T') { $Script:CustomCategories += "ScheduledTasks" }
        if ($cats -match 'A') { $Script:CustomCategories += "AppxPackages" }
        if ($cats -match 'U') { $Script:CustomCategories += "DefaultUser" }
        if ($cats -match 'L') { $Script:CustomCategories += "Autologgers" }
        if ($cats -match 'N') { $Script:CustomCategories += "Network" }
        if ($cats -match 'P') { $Script:CustomCategories += "Policy" }
        if ($cats -match 'D') { $Script:CustomCategories += "DiskCleanup" }
        if ($cats -match 'E') { $Script:CustomCategories += "Edge" }
        if ($cats -match 'R') { $Script:CustomCategories += "Privacy" }
        if ($cats -match 'W') { $Script:CustomCategories += "WindowsUpdate" }
        if ($cats -match 'I') { $Script:CustomCategories += "UIPersonalization" }
        if ($cats -match 'O') { $Script:CustomCategories += "SoundsNotifications" }
        if ($cats -match 'X') { $Script:CustomCategories += "ExplorerAutoplay" }
    }

    if (-not $Script:DryRun -and (Read-Host "`nApply changes? (Y/N)") -notmatch '^Y$') {
        Write-Host "Cancelled." -ForegroundColor Red
        exit
    }
} else {
    $Script:DryRun = $DryRun
    $Script:VerboseLogging = $Verbose
    if ($Optimizations -contains "All") { $Script:RunLevel = "Full" }
    elseif ($Optimizations -contains "Medium") { $Script:RunLevel = "Medium" }
    elseif ($Optimizations -contains "Safe") { $Script:RunLevel = "Safe" }
    elseif ($Optimizations -contains "Registry") { $Script:RunLevel = "Registry" }
    else { $Script:RunLevel = "Custom"; $Script:CustomCategories = $Optimizations }
}

Write-Log "Starting ImagePrep - Mode: $Script:RunLevel, DryRun: $Script:DryRun" "INFO"

function Invoke-SafeMode {
    Write-Log "Running Safe Mode optimizations" "INFO"
    if ($Script:Configs.ScheduledTasks) {
        $tasks = $Script:Configs.ScheduledTasks | Where-Object { $_.VDIState -eq 'Disabled' }
        Write-Log "Processing $($tasks.Count) scheduled tasks" "INFO"
        foreach ($task in $tasks) {
            try {
                $taskObj = Get-ScheduledTask $task.ScheduledTask -EA SilentlyContinue
                if ($taskObj -and $taskObj.State -ne 'Disabled') {
                    if (-not $Script:DryRun) { Disable-ScheduledTask -InputObject $taskObj | Out-Null }
                    Write-Log "Disabled: $($task.ScheduledTask)" "INFO"
                }
            } catch { Write-Log "Failed: $($task.ScheduledTask) - $($_.Exception.Message)" "WARN" }
        }
    }
}

function Invoke-MediumMode {
    Write-Log "Running Medium Mode optimizations" "INFO"
    Invoke-SafeMode

    if ($Script:Configs.Services) {
        $services = $Script:Configs.Services | Where-Object { $_.VDIState -eq 'Disabled' }
        Write-Log "Processing $($services.Count) services" "INFO"
        foreach ($svc in $services) {
            try {
                if (-not $Script:DryRun) { Set-Service -Name $svc.Name -StartupType Disabled -EA SilentlyContinue }
                Write-Log "Disabled service: $($svc.Name)" "INFO"
            } catch { Write-Log "Failed: $($svc.Name) - $($_.Exception.Message)" "WARN" }
        }
    }

    if ($Script:Configs.ServicesStartType) {
        Apply-ServiceStartType -Services $Script:Configs.ServicesStartType -CategoryName "ServicesStartType"
    }

    if ($Script:Configs.Autologgers) {
        $loggers = $Script:Configs.Autologgers | Where-Object { $_.Disabled -eq 'True' }
        Write-Log "Processing $($loggers.Count) autologgers" "INFO"
        foreach ($logger in $loggers) {
            try {
                if (-not $Script:DryRun) {
                    New-ItemProperty -Path $logger.KeyName -Name "Start" -PropertyType DWORD -Value 0 -Force -EA SilentlyContinue | Out-Null
                }
                Write-Log "Disabled autologger: $($logger.KeyName)" "INFO"
            } catch { Write-Log "Failed: $($logger.KeyName)" "WARN" }
        }
    }
}

function Invoke-AppxCleanup {
    Write-Log "Processing Appx Packages" "INFO"
    if (-not $Script:Configs.AppxPackages) { return }
    $apps = $Script:Configs.AppxPackages | Where-Object { $_.VDIState -eq 'Disabled' }
    Write-Log "Removing $($apps.Count) provisioned packages" "INFO"
    foreach ($app in $apps) {
        try {
            if (-not $Script:DryRun) {
                Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -like "*$($app.AppxPackage)*" } | Remove-AppxProvisionedPackage -Online -AllUsers -EA SilentlyContinue | Out-Null
                Get-AppxPackage -AllUsers -Name "*$($app.AppxPackage)*" | Remove-AppxPackage -AllUsers -EA SilentlyContinue | Out-Null
            }
            Write-Log "Removed Appx: $($app.AppxPackage)" "INFO"
        } catch { Write-Log "Failed: $($app.AppxPackage)" "WARN" }
    }
}

function Invoke-DefaultUserSettings {
    Write-Log "Processing Default User Settings" "INFO"
    if (-not $Script:Configs.DefaultUser) { return }
    $settings = $Script:Configs.DefaultUser | Where-Object { $_.SetProperty -eq $true }
    Write-Log "Processing $($settings.Count) default user settings" "INFO"
    try {
        $null = Start-Process reg -ArgumentList "LOAD HKLM\IP_TEMP C:\Users\Default\NTUSER.DAT" -PassThru -Wait -WindowStyle Hidden
        foreach ($item in $settings) {
            try {
                if ($item.PropertyType -eq "BINARY") { $value = [byte[]]($item.PropertyValue.Split(",")) }
                else { $value = $item.PropertyValue }
                if (-not $Script:DryRun) {
                    if (Get-ItemProperty -Path $item.HivePath -Name $item.KeyName -EA SilentlyContinue) {
                        Set-ItemProperty -Path $item.HivePath -Name $item.KeyName -Value $value -Type $item.PropertyType -Force
                    } else {
                        New-ItemProperty -Path $item.HivePath -Name $item.KeyName -PropertyType $item.PropertyType -Value $value -Force | Out-Null
                    }
                }
                Write-Log "Applied: $($item.KeyName)" "INFO"
            } catch { Write-Log "Failed: $($item.KeyName)" "WARN" }
        }
        $null = Start-Process reg -ArgumentList "UNLOAD HKLM\IP_TEMP" -PassThru -Wait -WindowStyle Hidden
    } catch { Write-Log "Failed to load default user hive" "ERROR" }
}

function Invoke-NetworkOptimizations {
    Write-Log "Processing Network Optimizations" "INFO"
    if (-not $Script:Configs.LanManWorkstation) { return }
    foreach ($hive in $Script:Configs.LanManWorkstation) {
        $keys = $hive.Keys | Where-Object { $_.SetProperty -eq $true }
        foreach ($key in $keys) {
            try {
                if (-not $Script:DryRun) {
                    if (Get-ItemProperty -Path $hive.HivePath -Name $key.Name -EA SilentlyContinue) {
                        Set-ItemProperty -Path $hive.HivePath -Name $key.Name -Value $key.PropertyValue -Force
                    } else {
                        New-ItemProperty -Path $hive.HivePath -Name $key.Name -PropertyType $key.PropertyType -Value $key.PropertyValue -Force | Out-Null
                    }
                }
                Write-Log "Network: Set $($key.Name)" "INFO"
            } catch { Write-Log "Failed: $($key.Name)" "WARN" }
        }
    }
}

function Invoke-PolicySettings {
    Write-Log "Processing Policy Settings" "INFO"
    if (-not $Script:Configs.PolicyRegSettings) { return }
    $policies = $Script:Configs.PolicyRegSettings | Where-Object { $_.VDIState -eq 'Enabled' }
    foreach ($policy in $policies) {
        try {
            if (-not $Script:DryRun) {
                if (Get-ItemProperty -Path $policy.RegItemPath -Name $policy.RegItemValueName -EA SilentlyContinue) {
                    Set-ItemProperty -Path $policy.RegItemPath -Name $policy.RegItemValueName -Value $policy.RegItemValue -Force
                } else {
                    if (Test-Path $policy.RegItemPath) {
                        New-ItemProperty -Path $policy.RegItemPath -Name $policy.RegItemValueName -PropertyType $policy.RegItemValueType -Value $policy.RegItemValue -Force | Out-Null
                    } else {
                        New-Item -Path $policy.RegItemPath -Force | New-ItemProperty -Name $policy.RegItemValueName -PropertyType $policy.RegItemValueType -Value $policy.RegItemValue -Force | Out-Null
                    }
                }
            }
            Write-Log "Policy: $($policy.RegItemValueName)" "INFO"
        } catch { Write-Log "Failed: $($policy.RegItemValueName)" "WARN" }
    }
}

function Invoke-EdgeSettings {
    Write-Log "Processing Edge Settings" "INFO"
    if (-not $Script:Configs.EdgeSettings) { return }
    $edge = $Script:Configs.EdgeSettings | Where-Object { $_.VDIState -eq 'Enabled' }
    foreach ($setting in $edge) {
        try {
            if (-not $Script:DryRun) {
                if ($setting.RegItemValueName -eq 'DefaultAssociationsConfiguration') {
                    Copy-Item (Join-Path $JsonFolder "DefaultAssociationsConfiguration.xml") $setting.RegItemValue -Force -EA SilentlyContinue
                }
                if (Get-ItemProperty -Path $setting.RegItemPath -Name $setting.RegItemValueName -EA SilentlyContinue) {
                    Set-ItemProperty -Path $setting.RegItemPath -Name $setting.RegItemValueName -Value $setting.RegItemValue -Force
                } else {
                    New-ItemProperty -Path $setting.RegItemPath -Name $setting.RegItemValueName -PropertyType $setting.RegItemValueType -Value $setting.RegItemValue -Force | Out-Null
                }
            }
            Write-Log "Edge: $($setting.RegItemValueName)" "INFO"
        } catch { Write-Log "Failed: $($setting.RegItemValueName)" "WARN" }
    }
}

function Invoke-PrivacyTelemetry {
    Write-Log "Processing Privacy & Telemetry" "INFO"
    Apply-RegistrySettings -Settings $Script:Configs.PrivacyTelemetry -CategoryName "PrivacyTelemetry"
}

function Invoke-WindowsUpdate {
    Write-Log "Processing Windows Update Settings" "INFO"
    Apply-RegistrySettings -Settings $Script:Configs.WindowsUpdate -CategoryName "WindowsUpdate"
}

function Invoke-UIPersonalization {
    Write-Log "Processing UI Personalization" "INFO"
    Apply-RegistrySettings -Settings $Script:Configs.UIPersonalization -CategoryName "UIPersonalization"
}

function Invoke-SoundsNotifications {
    Write-Log "Processing Sounds & Notifications" "INFO"
    Apply-RegistrySettings -Settings $Script:Configs.SoundsNotifications -CategoryName "SoundsNotifications"
}

function Invoke-ExplorerAutoplay {
    Write-Log "Processing Explorer & Autoplay" "INFO"
    Apply-RegistrySettings -Settings $Script:Configs.ExplorerAutoplay -CategoryName "ExplorerAutoplay"
}

function Invoke-DiskCleanup {
    Write-Log "Running Disk Cleanup" "INFO"
    if ($Script:DryRun) {
        Write-Log "[DRY RUN] Would clean temp files, logs, caches" "INFO"
        return
    }
    try {
        Get-ChildItem -Path c:\ -Include *.tmp,*.dmp,*.etl,*.evtx,thumbcache*.db,*.log -File -Recurse -Force -EA SilentlyContinue | Remove-Item -EA SilentlyContinue
        Get-ChildItem -Path $env:windir\Temp\* -Recurse -Force -EA SilentlyContinue | Remove-Item -Recurse -Force -EA SilentlyContinue -Exclude packer*.ps1
        Get-ChildItem -Path $env:TEMP\* -Recurse -Force -EA SilentlyContinue | Remove-Item -Recurse -Force -EA SilentlyContinue -Exclude packer*.ps1
        Get-ChildItem -Path $env:ProgramData\Microsoft\Windows\WER\Temp\* -Recurse -Force -EA SilentlyContinue | Remove-Item -Recurse -Force -EA SilentlyContinue
        Clear-RecycleBin -Force -EA SilentlyContinue
        Write-Log "Disk cleanup completed" "INFO"
    } catch { Write-Log "Disk cleanup error: $($_.Exception.Message)" "WARN" }
}

function Invoke-Category {
    param([string]$Category)
    switch ($Category) {
        "Services"           { Invoke-MediumMode }
        "ScheduledTasks"    { Invoke-SafeMode }
        "AppxPackages"     { Invoke-AppxCleanup }
        "DefaultUser"      { Invoke-DefaultUserSettings }
        "Autologgers"      { Invoke-MediumMode }
        "Network"          { Invoke-NetworkOptimizations }
        "Policy"           { Invoke-PolicySettings }
        "DiskCleanup"      { Invoke-DiskCleanup }
        "Edge"             { Invoke-EdgeSettings }
        "Privacy"          { Invoke-PrivacyTelemetry }
        "WindowsUpdate"    { Invoke-WindowsUpdate }
        "UIPersonalization" { Invoke-UIPersonalization }
        "ServicesStartType" { Invoke-MediumMode }
        "SoundsNotifications" { Invoke-SoundsNotifications }
        "ExplorerAutoplay" { Invoke-ExplorerAutoplay }
    }
}

try {
    if ($Script:RunLevel -eq "Safe") {
        Invoke-SafeMode
    } elseif ($Script:RunLevel -eq "Medium") {
        Invoke-MediumMode
    } elseif ($Script:RunLevel -eq "Registry") {
        Invoke-PrivacyTelemetry
        Invoke-WindowsUpdate
        Invoke-UIPersonalization
        Invoke-SoundsNotifications
        Invoke-ExplorerAutoplay
    } elseif ($Script:RunLevel -eq "Full") {
        Invoke-SafeMode
        Invoke-MediumMode
        Invoke-AppxCleanup
        Invoke-DefaultUserSettings
        Invoke-NetworkOptimizations
        Invoke-PolicySettings
        Invoke-PrivacyTelemetry
        Invoke-WindowsUpdate
        Invoke-UIPersonalization
        Invoke-SoundsNotifications
        Invoke-ExplorerAutoplay
        Invoke-DiskCleanup
        Invoke-EdgeSettings
    } elseif ($Script:RunLevel -eq "Custom" -and $Script:CustomCategories) {
        foreach ($cat in $Script:CustomCategories) {
            Invoke-Category $cat
        }
    }

    Write-Log "ImagePrep completed. Log: $LogPath" "INFO"
    Write-Host "`n============================================" -ForegroundColor Green
    Write-Host "Done! Log saved to: $LogPath" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green

    if ($Restart -and -not $Script:DryRun) {
        Write-Host "Restart required. Rebooting in 10 seconds..." -ForegroundColor Yellow
        Start-Sleep 5
        Restart-Computer -Force
    } elseif (-not $Script:DryRun) {
        Write-Host "`nA reboot may be required for all changes to take effect." -ForegroundColor Yellow
    }
} catch {
    Write-Log "Critical error: $($_.Exception.Message)" "ERROR"
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
