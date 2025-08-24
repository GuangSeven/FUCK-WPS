#Requires -RunAsAdministrator

# 核心配置
$priorityPath = "HKLM:\SOFTWARE\Classes\SystemFileAssociations"
$targetPaths = @(
    "HKCU:\Software",
    "HKLM:\Software",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts"
)
$keywords = @("WPS", "Kingsoft", "KWPS", "WPS Office", "\.wps")
$exclude = @("AMDWPS", "UWPSystem", "amd")  # 排除规则

# 统计变量
$script:deleted = 0
$script:modified = 0
$script:skipped = 0  # 仅统计「WPS相关但被排除」的项


### 判断项是否与WPS相关（核心辅助函数）
function IsWPSRelated {
    param(
        [string]$keyPath,
        [array]$keywords
    )
    $keyName = (Split-Path $keyPath -Leaf)
    
    # 检查项名是否含WPS关键词
    foreach ($kw in $keywords) {
        if ($keyName -match $kw) {
            return $true
        }
    }
    
    # 检查属性值是否含WPS关键词（仅字符串类型）
    try {
        $props = Get-ItemProperty -Path $keyPath -ErrorAction SilentlyContinue
        if ($props) {
            $props | Get-Member -MemberType NoteProperty | ForEach-Object {
                $propValue = $props.$($_.Name)
                if ($propValue -is [string]) {
                    foreach ($kw in $keywords) {
                        if ($propValue -match $kw) {
                            return $true
                        }
                    }
                }
            }
        }
    } catch {}
    
    return $false
}


### 1. 优先清理指定路径（核心路径）
function Remove-WPSPriority {
    if (-not (Test-Path $priorityPath)) {
        Write-Host "Error: Path not found - $priorityPath" -ForegroundColor Red
        return
    }
    Write-Host "=== Processing priority path: $priorityPath ===" -ForegroundColor Cyan

    Get-ChildItem $priorityPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $keyPath = $_.PSPath
        $keyName = $_.Name

        # 第一步：判断是否是WPS相关项（不相关则直接跳过，不统计）
        if (-not (IsWPSRelated $keyPath $keywords)) {
            return
        }

        # 第二步：检查是否匹配排除规则（仅WPS相关项才会进入此判断）
        foreach ($ex in $exclude) {
            if ($keyName -match $ex) {
                Write-Host "Skipped (WPS-related but excluded): $keyPath" -ForegroundColor Cyan
                $script:skipped++
                return
            }
        }

        # ========== 处理TypeOverlay（仅当属性存在时） ==========
        $overlay = Get-ItemProperty -Path $keyPath -Name "TypeOverlay" -ErrorAction SilentlyContinue
        if ($overlay) {
            foreach ($kw in $keywords) {
                if ($overlay.TypeOverlay -match $kw) {
                    try {
                        Remove-ItemProperty -Path $keyPath -Name "TypeOverlay" -Force -ErrorAction Stop
                        Write-Host "Deleted TypeOverlay: $keyPath" -ForegroundColor Green
                        $script:deleted++
                    } catch {
                        Write-Host "Delete TypeOverlay failed: $keyPath - $_" -ForegroundColor Red
                    }
                    break
                }
            }
        }

        # 处理注册表项（删除或修改）
        foreach ($kw in $keywords) {
            if ($keyName -match $kw) {
                try {
                    Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                    Write-Host "Deleted item: $keyPath" -ForegroundColor Green
                    $script:deleted++
                } catch {
                    if ($_.Exception -is [System.Security.SecurityException]) {
                        Write-Host "Permission denied: $keyPath" -ForegroundColor DarkRed
                        $script:skipped++  # 权限不足的WPS项也统计到跳过
                    } else {
                        try {
                            $props = Get-ItemProperty -Path $keyPath -ErrorAction Stop
                            $props | Get-Member -MemberType NoteProperty | ForEach-Object {
                                $propName = $_.Name
                                $propValue = $props.$propName
                                if ($propValue -match $kw) {
                                    Set-ItemProperty -Path $keyPath -Name $propName -Value "" -Force -ErrorAction Stop
                                    Write-Host "Modified value: $keyPath\$propName" -ForegroundColor Yellow
                                    $script:modified++
                                }
                            }
                        } catch {
                            Write-Host "Modify failed: $keyPath - $_" -ForegroundColor Red
                            $script:skipped++  # 修改失败的WPS项也统计到跳过
                        }
                    }
                }
                break
            }
        }
    }
}


### 2. 清理其他相关路径
function Remove-WPSTargets {
    foreach ($path in $targetPaths) {
        if (-not (Test-Path $path)) {
            Write-Host "Error: Path not found - $path" -ForegroundColor Red
            continue
        }
        Write-Host "=== Processing path: $path ===" -ForegroundColor Cyan

        Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $keyPath = $_.PSPath
            $keyName = $_.Name

            # 第一步：判断是否是WPS相关项（不相关则直接跳过，不统计）
            if (-not (IsWPSRelated $keyPath $keywords)) {
                return
            }

            # 第二步：检查是否匹配排除规则（仅WPS相关项才会进入此判断）
            foreach ($ex in $exclude) {
                if ($keyName -match $ex) {
                    Write-Host "Skipped (WPS-related but excluded): $keyPath" -ForegroundColor Cyan
                    $script:skipped++
                    return
                }
            }

            # 处理注册表项（删除或修改）
            foreach ($kw in $keywords) {
                if ($keyName -match $kw) {
                    try {
                        Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                        Write-Host "Deleted item: $keyPath" -ForegroundColor Green
                        $script:deleted++
                    } catch {
                        if ($_.Exception -is [System.Security.SecurityException]) {
                            Write-Host "Permission denied: $keyPath" -ForegroundColor DarkRed
                            $script:skipped++
                        } else {
                            try {
                                $props = Get-ItemProperty -Path $keyPath -ErrorAction Stop
                                $props | Get-Member -MemberType NoteProperty | ForEach-Object {
                                    $propName = $_.Name
                                    $propValue = $props.$propName
                                    if ($propValue -match $kw) {
                                        Set-ItemProperty -Path $keyPath -Name $propName -Value "" -Force -ErrorAction Stop
                                        Write-Host "Modified value: $keyPath\$propName" -ForegroundColor Yellow
                                        $script:modified++
                                    }
                                }
                            } catch {
                                Write-Host "Modify failed: $keyPath - $_" -ForegroundColor Red
                                $script:skipped++
                            }
                        }
                    }
                    break
                }
            }
        }
    }
}


### 执行清理（带开始提示）
Write-Host "=== Starting WPS Registry Cleanup ===" -ForegroundColor Magenta
Remove-WPSPriority
Remove-WPSTargets

# 最终汇总（仅统计WPS相关项的处理结果）
Write-Host "`n=== Cleanup Summary (WPS-related items only) ===" -ForegroundColor Magenta
Write-Host "Deleted: $script:deleted" -ForegroundColor Green
Write-Host "Modified: $script:modified" -ForegroundColor Yellow
Write-Host "Skipped (excluded/denied/failed): $script:skipped" -ForegroundColor Cyan