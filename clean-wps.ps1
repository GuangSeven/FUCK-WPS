#Requires -Version 7
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    清理 WPS 残留注册表项工具 - 优化版

.DESCRIPTION
    基于 .NET RegistryKey API 的高性能注册表清理脚本
    支持配置文件、进度条、静默模式、演练模式、自动备份、日志记录、图标缓存刷新

.PARAMETER ConfigFile
    配置文件路径 (JSON)，默认脚本同目录 config.json

.PARAMETER WhatIf
    演练模式：仅显示将要执行的操作，不实际修改注册表

.PARAMETER Backup
    自动导出 .reg 备份到脚本目录 (默认开启)

.PARAMETER NoBackup
    禁用自动备份

.PARAMETER LogFile
    日志文件路径，默认脚本目录 Clean-WPS-YYYYMMDD-HHMMSS.log

.PARAMETER Verbose
    显示详细操作日志

.PARAMETER Quiet
    静默模式：仅显示进度条和最终汇总 (默认)

.PARAMETER NoRefreshIcons
    跳过图标缓存刷新确认提示

.PARAMETER NoProgress
    禁用进度条

.EXAMPLE
    .\Clean-WPS.ps1
    # 默认静默模式，自动备份，自动刷新图标缓存

.EXAMPLE
    .\Clean-WPS.ps1 -Verbose -WhatIf
    # 详细演练模式

.EXAMPLE
    .\Clean-WPS.ps1 -ConfigFile "D:\Config\my-wps.json" -LogFile "D:\Logs\clean.log" -NoBackup
    # 自定义配置、日志，禁用备份
#>

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    [Parameter(Position = 0)]
    [string] $ConfigFile = (Join-Path $PSScriptRoot 'config.json'),

    [switch] $WhatIf,
    [switch] $Backup,
    [switch] $NoBackup,
    [string] $LogFile,
    [switch] $Verbose,
    [switch] $Quiet,
    [switch] $NoRefreshIcons,
    [switch] $NoProgress
)

$ErrorActionPreference = 'Stop'
$global:CancelToken = $false

Set-StrictMode -Version Latest

# ──────────────────────────────────────────────────────────────
# 配置加载
# ──────────────────────────────────────────────────────────────
if (-not (Test-Path $ConfigFile)) {
    throw "配置文件不存在: $ConfigFile"
}
$config = Get-Content $ConfigFile -Raw | ConvertFrom-Json -Depth 10

$keywords = $config.keywords
$excludePatterns = $config.excludePatterns
$registryPaths = $config.registryPaths
$priorityPaths = $config.priorityPaths

# 将优先路径转换为 HKLM 哈希表格式
$priorityPathsTable = @{ 'HKLM' = @() }
foreach ($path in $priorityPaths) {
    if ($path -match '^HKLM:\\') {
        $priorityPathsTable['HKLM'] += $path.Substring(6)
    } elseif ($path -match '^HKCU:\\') {
        $priorityPathsTable['HKCU'] += $path.Substring(6)
    } else {
        # 默认归类到 HKLM
        $priorityPathsTable['HKLM'] += $path
    }
}
$settings = $config.settings

# 将优先路径转换为 hashtable 格式
$priorityPathsTable = @{ 'HKLM' = @() }
foreach ($path in $priorityPaths) {
    if ($path.StartsWith('HKLM:')) {
        $priorityPathsTable['HKLM'] += $path.Substring('HKLM:\'.Length)
    } elseif ($path.StartsWith('HKCU:')) {
        $priorityPathsTable['HKCU'] += $path.Substring('HKCU:\'.Length)
    }
}

$kwRegex = [System.Text.RegularExpressions.Regex]::new(
    '(' + ($keywords -join '|') + ')',
    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
)
$excludeRegex = if ($excludePatterns.Count -gt 0) {
    [System.Text.RegularExpressions.Regex]::new(
        '(' + ($excludePatterns -join '|') + ')',
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
} else {
    $null
}

# ──────────────────────────────────────────────────────────────
# 全局状态
# ──────────────────────────────────────────────────────────────
$script:stats = @{
    ScannedKeys   = 0
    ScannedValues = 0
    DeletedKeys   = 0
    DeletedValues = 0
    ModifiedValues = 0
    Skipped       = 0
    Errors        = 0
}
$script:logEntries = @()
$script:backupPath = $null
$script:startTime = Get-Date

# ──────────────────────────────────────────────────────────────
# 日志系统
# ──────────────────────────────────────────────────────────────
function Initialize-Log {
    if (-not $LogFile) {
        $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $LogFile = Join-Path $PSScriptRoot "Clean-WPS-$timestamp.log"
    }
    $script:logFilePath = $LogFile
    $header = @(
        "=== Clean-WPS Log ===",
        "Start Time: $($script:startTime.ToString('yyyy-MM-dd HH:mm:ss'))",
        "Script: $PSScriptRoot\Clean-WPS.ps1",
        "Config: $ConfigFile",
        "WhatIf: $WhatIf",
        "Backup: $((-not $NoBackup) -and $settings.autoBackup)",
        "Verbose: $Verbose",
        "Quiet: $Quiet",
        "==========================================",
        ""
    )
    [System.IO.File]::WriteAllLines($script:logFilePath, $header, [System.Text.Encoding]::UTF8)
}

function Write-Log {
    param(
        [string] $Level,
        [string] $Message,
        [string] $Path = ''
    )
    $entry = "[{0:yyyy-MM-dd HH:mm:ss.fff}] [{1,-7}] {2} {3}" -f (Get-Date), $Level, $Message, $Path
    $script:logEntries += $entry
    if ($script:logEntries.Count -ge 100) {
        Flush-Log
    }
    if ($Verbose -or $Level -in @('ERROR', 'WARN')) {
        $color = switch ($Level) {
            'ERROR' { 'Red' }
            'WARN'  { 'Yellow' }
            'OK'    { 'Green' }
            'INFO'  { 'Cyan' }
            default { 'Gray' }
        }
        Write-Host $entry -ForegroundColor $color
    }
}

function Flush-Log {
    if ($script:logEntries.Count -gt 0) {
        [System.IO.File]::AppendAllLines($script:logFilePath, $script:logEntries, [System.Text.Encoding]::UTF8)
        $script:logEntries.Clear()
    }
}

function Close-Log {
    Flush-Log
    $footer = @(
        "",
        "=== Summary ===",
        "End Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "Duration: $((Get-Date) - $script:startTime)",
        "Scanned Keys: $($script:stats.ScannedKeys)",
        "Scanned Values: $($script:stats.ScannedValues)",
        "Deleted Keys: $($script:stats.DeletedKeys)",
        "Deleted Values: $($script:stats.DeletedValues)",
        "Modified Values: $($script:stats.ModifiedValues)",
        "Skipped: $($script:stats.Skipped)",
        "Errors: $($script:stats.Errors)",
        "Log File: $script:logFilePath"
    )
    [System.IO.File]::AppendAllLines($script:logFilePath, $footer, [System.Text.Encoding]::UTF8)
}

# ──────────────────────────────────────────────────────────────
# 进度条与中断处理
# ──────────────────────────────────────────────────────────────
function Update-Progress {
    param(
        [string] $Status,
        [int] $PercentComplete = -1
    )
    if ($NoProgress -or $Quiet) { return }
    if ($PercentComplete -ge 0) {
        Write-Progress -Activity 'Clean-WPS 注册表清理' -Status $Status -PercentComplete $PercentComplete
    } else {
        Write-Progress -Activity 'Clean-WPS 注册表清理' -Status $Status
    }
}

function Show-Summary {
    $duration = (Get-Date) - $script:startTime
    $summary = @"
========================================
Clean-WPS 清理完成
========================================
耗时: $($duration.ToString('hh\:mm\:ss'))
扫描键值: $($script:stats.ScannedKeys)
扫描值:   $($script:stats.ScannedValues)
已删除键: $($script:stats.DeletedKeys)
已删除值: $($script:stats.DeletedValues)
已修改值: $($script:stats.ModifiedValues)
已跳过:   $($script:stats.Skipped)
错误数:   $($script:stats.Errors)
日志文件: $script:logFilePath
"@
    if (-not $Quiet) {
        Write-Host $summary -ForegroundColor Magenta
    }
    Write-Log 'INFO' $summary.Replace("`n", " | ")
}

trap {
    $global:CancelToken = $true
    Write-Warning "`n检测到中断信号 (Ctrl+C)，正在汇总..."
    Show-Summary
    Close-Log
    exit 1
}

# ──────────────────────────────────────────────────────────────
# 注册表备份
# ──────────────────────────────────────────────────────────────
function Export-RegistryBackup {
    if ($NoBackup -or -not $settings.autoBackup -or $WhatIf) { return }
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $script:backupPath = Join-Path $PSScriptRoot "WPS-RegBackup-$timestamp.reg"
    Write-Log 'INFO' "导出注册表备份: $script:backupPath"
    try {
        $keysToExport = @()
        foreach ($hive in @('HKLM', 'HKCU')) {
            if ($registryPaths.$hive) {
                foreach ($path in $registryPaths.$hive) {
                    $keysToExport += "$hive\$path"
                }
            }
        }
        foreach ($hive in $priorityPathsTable.Keys) {
            if ($priorityPathsTable[$hive]) {
                foreach ($path in $priorityPathsTable[$hive]) {
                    $keysToExport += "$hive\$path"
                }
            }
        }
        $uniqueKeys = $keysToExport | Sort-Object -Unique
        $regArgs = $uniqueKeys | ForEach-Object { "/e `"$script:backupPath`" $_" } -join ' '
        & regedit.exe $regArgs
        Write-Log 'OK' "备份完成: $script:backupPath"
    } catch {
        Write-Log 'ERROR' "备份失败: $_"
        $script:stats.Errors++
    }
}

# ──────────────────────────────────────────────────────────────
# 核心扫描与清理逻辑 (使用 .NET RegistryKey API)
# ──────────────────────────────────────────────────────────────
function Get-RegistryHive {
    param([string] $HiveName)
    switch ($HiveName.ToUpper()) {
        'HKLM' { return [Microsoft.Win32.Registry]::LocalMachine }
        'HKCU' { return [Microsoft.Win32.Registry]::CurrentUser }
        default { throw "未知配置单元: $HiveName" }
    }
}

function Test-Exclude {
    param([string] $Name)
    if (-not $excludeRegex) { return $false }
    return $excludeRegex.IsMatch($Name)
}

function Test-KeywordMatch {
    param([string] $InputString)
    return $kwRegex.IsMatch($InputString)
}

function Process-RegistryKey {
    param(
        [Microsoft.Win32.RegistryKey] $Key,
        [string] $FullPath,
        [bool] $IsPriorityPath = $false
    )

    if ($global:CancelToken) { return }

    $keyName = $FullPath.Split('\')[-1]
    $script:stats.ScannedKeys++

    if (Test-Exclude $keyName) {
        Write-Log 'INFO' "跳过(排除): $FullPath"
        $script:stats.Skipped++
        return
    }

    # 检查键名是否匹配关键词
    $keyMatches = Test-KeywordMatch $keyName

    # 枚举值
    try {
        $valueNames = $Key.GetValueNames()
        foreach ($valueName in $valueNames) {
            if ($global:CancelToken) { return }
            $script:stats.ScannedValues++
            $value = $Key.GetValue($valueName)
            $valueStr = if ($value -is [string]) { $value } else { $value -join ',' }
            $valueMatches = Test-KeywordMatch $valueStr

            if ($keyMatches -or $valueMatches) {
                $action = if ($keyMatches) { "删除键: $FullPath" } else { "清理值: $FullPath\$valueName" }
                Write-Log 'INFO' $action

                if (-not $WhatIf) {
                    try {
                        if ($keyMatches) {
                            # 删除整个子键 - 使用父键删除子键树
                            $parentPath = $FullPath.Substring(0, $FullPath.LastIndexOf('\'))
                            $hiveName = $FullPath.Split('\')[0]
                            $parentHive = Get-RegistryHive $hiveName
                            $subKeyPath = $parentPath.Substring($parentPath.IndexOf('\') + 1)
                            $parentKey = $parentHive.OpenSubKey($subKeyPath, $true)
                            if ($parentKey) {
                                $parentKey.DeleteSubKeyTree($keyName)
                                $parentKey.Close()
                            }
                            $script:stats.DeletedKeys++
                            Write-Log 'OK' "已删除键: $FullPath"
                            return # 键已删除，无需继续处理值
                        } else {
                            # 清空值
                            $Key.SetValue($valueName, '', [Microsoft.Win32.RegistryValueKind]::String)
                            $script:stats.ModifiedValues++
                            Write-Log 'OK' "已清空值: $FullPath\$valueName"
                        }
                    } catch {
                        if ($_.Exception -is [System.Security.SecurityException]) {
                            Write-Log 'WARN' "权限不足，跳过: $FullPath"
                        } else {
                            Write-Log 'ERROR' "操作失败: $FullPath - $_"
                        }
                        $script:stats.Skipped++
                    }
                }
            }
        }
    } catch {
        Write-Log 'WARN' "枚举值失败: $FullPath - $_"
    }

    # 递归处理子键
    try {
        $subKeyNames = $Key.GetSubKeyNames()
        foreach ($subKeyName in $subKeyNames) {
            if ($global:CancelToken) { return }
            $subKey = $Key.OpenSubKey($subKeyName, $true) # 可写
            if ($subKey) {
                Process-RegistryKey $subKey "$FullPath\$subKeyName" $IsPriorityPath
                $subKey.Close()
            }
        }
    } catch {
        Write-Log 'WARN' "枚举子键失败: $FullPath - $_"
    }
}

function Expand-WildcardPath {
    param(
        [Microsoft.Win32.RegistryKey] $Hive,
        [string] $PatternPath
    )
    
    # 如果没有通配符，直接返回
    if ($PatternPath -notmatch '[\*\?]') {
        return @($PatternPath)
    }
    
    $parts = $PatternPath.Split('\\')
    $currentKeys = @(, @{ Key = $Hive; Path = '' })
    $expandedPaths = @()
    
    foreach ($part in $parts) {
        $nextKeys = @()
        $isLast = ($part -eq $parts[-1])
        
        foreach ($item in $currentKeys) {
            $key = $item.Key
            $parentPath = $item.Path
            try {
                $subKeyNames = $key.GetSubKeyNames()
                $matches = @()
                
                if ($part -eq '*') {
                    # 完整通配符：匹配所有子键
                    $matches = $subKeyNames
                } elseif ($part -match '^\{.*\}$') {
                    # GUID 模式：{*}
                    $matches = $subKeyNames | Where-Object { $_ -match '^\{.*\}$' }
                } elseif ($part -match '^\*\..+') {
                    # 后缀通配符：如 *.wps
                    $suffix = $part.Substring(2)
                    $matches = $subKeyNames | Where-Object { $_ -like "*$suffix" }
                } elseif ($part -match '.+\*$') {
                    # 前缀通配符：如 WPS.* 或 .wps*
                    $prefix = $part.Substring(0, $part.Length - 1)
                    $matches = $subKeyNames | Where-Object { $_ -like "$prefix*" }
                } else {
                    # 精确匹配
                    if ($subKeyNames -contains $part) {
                        $matches = @($part)
                    }
                }
                
                foreach ($subName in $matches) {
                    $subKey = $key.OpenSubKey($subName, $true)
                    if ($subKey) {
                        $fullPath = if ($parentPath) { "$parentPath\$subName" } else { $subName }
                        if ($isLast) {
                            $expandedPaths += $fullPath
                        } else {
                            $nextKeys += @{ Key = $subKey; Path = $fullPath }
                        }
                    }
                }
            } catch {
                Write-Log 'WARN' "展开路径失败: $parentPath\$part - $_"
            }
        }
        # 关闭非最后一层的键
        if (-not $isLast) {
            foreach ($item in $currentKeys) { $item.Key.Close() }
        }
        $currentKeys = $nextKeys
    }
    
    # 关闭剩余键
    foreach ($item in $currentKeys) { $item.Key.Close() }
    
    return $expandedPaths
}

function Process-PathList {
    param(
        [hashtable] $Paths,
        [bool] $IsPriority = $false
    )

    foreach ($hiveName in $Paths.Keys) {
        $hive = Get-RegistryHive $hiveName
        $pathList = $Paths[$hiveName]
        
        # 展开通配符路径
        $expandedPaths = @()
        foreach ($patternPath in $pathList) {
            $expanded = Expand-WildcardPath $hive $patternPath
            $expandedPaths += $expanded
        }
        
        $totalPaths = $expandedPaths.Count
        $currentPath = 0

        foreach ($subPath in $expandedPaths) {
            if ($global:CancelToken) { return }
            $currentPath++
            $percent = if ($totalPaths -gt 0) { [int](($currentPath / $totalPaths) * 100) } else { -1 }
            Update-Progress -Status "扫描: $hiveName\$subPath" -PercentComplete $percent

            try {
                $rootKey = $hive.OpenSubKey($subPath, $true)
                if ($rootKey) {
                    Process-RegistryKey $rootKey "$hiveName\$subPath" $IsPriority
                    $rootKey.Close()
                } else {
                    Write-Log 'INFO' "路径不存在: $hiveName\$subPath"
                }
            } catch {
                Write-Log 'WARN' "无法访问: $hiveName\$subPath - $_"
                $script:stats.Errors++
            }
        }
    }
}

# ──────────────────────────────────────────────────────────────
# 图标缓存刷新
# ──────────────────────────────────────────────────────────────
function Invoke-IconCacheRefresh {
    if ($NoRefreshIcons -or -not $settings.autoRefreshIcons) { return }
    if ($settings.confirmRefreshIcons -and -not $Quiet) {
        $confirm = Read-Host "`n是否立即刷新图标缓存 (需重启资源管理器)？[Y/N]"
        if ($confirm -notin @('Y', 'y', '是', 'yes')) {
            Write-Log 'INFO' '用户取消图标缓存刷新'
            Write-Host "已跳过图标缓存刷新。稍后可手动运行: ie4uinit.exe -ClearIconCache" -ForegroundColor Yellow
            return
        }
    }
    Write-Log 'INFO' '刷新图标缓存...'
    try {
        & ie4uinit.exe -ClearIconCache
        Write-Log 'OK' '图标缓存刷新命令已发送'
        Write-Host "图标缓存刷新完成。建议重启资源管理器或重启电脑使更改完全生效。" -ForegroundColor Cyan
    } catch {
        Write-Log 'ERROR' "刷新图标缓存失败: $_"
    }
}

# ──────────────────────────────────────────────────────────────
# 主流程
# ──────────────────────────────────────────────────────────────
try {
    Initialize-Log

    Write-Log 'INFO' '========== Clean-WPS 启动 =========='
    Write-Log 'INFO' "配置文件: $ConfigFile"
    Write-Log 'INFO' "演练模式: $WhatIf"
    Write-Log 'INFO' "自动备份: $((-not $NoBackup) -and $settings.autoBackup)"

    if (-not $Quiet) {
        Write-Host "=== Clean-WPS WPS 注册表清理工具 ===" -ForegroundColor Magenta
        Write-Host "版本: 2.0 (优化版)" -ForegroundColor Gray
        Write-Host "配置: $ConfigFile" -ForegroundColor Gray
        if ($WhatIf) { Write-Host "模式: 演练模式 (-WhatIf)" -ForegroundColor Yellow }
        Write-Host ""
    }

    Export-RegistryBackup

    Update-Progress -Status '开始扫描优先路径...' -PercentComplete 0
    Process-PathList -Paths $priorityPathsTable -IsPriority $true

    Update-Progress -Status '扫描 HKLM 目标路径...' -PercentComplete 30
    if ($registryPaths.HKLM) {
        Process-PathList -Paths @{ 'HKLM' = $registryPaths.HKLM }
    }

    Update-Progress -Status '扫描 HKCU 目标路径...' -PercentComplete 60
    if ($registryPaths.HKCU) {
        Process-PathList -Paths @{ 'HKCU' = $registryPaths.HKCU }
    }

    Update-Progress -Status '清理完成，生成汇总...' -PercentComplete 100

    Show-Summary
    Invoke-IconCacheRefresh

} catch {
    Write-Log 'ERROR' "脚本异常: $_"
    Write-Log 'ERROR' $_.ScriptStackTrace
    $script:stats.Errors++
    if (-not $Quiet) {
        Write-Host "错误: $_" -ForegroundColor Red
    }
} finally {
    if (-not $NoProgress) { Write-Progress -Completed }
    Close-Log
    if ($script:backupPath -and (Test-Path $script:backupPath)) {
        Write-Host "`n注册表备份已保存至: $script:backupPath" -ForegroundColor Green
    }
}