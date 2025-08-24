#Requires -RunAsAdministrator

# 核心配置
$priorityPath = "HKLM:\SOFTWARE\Classes\SystemFileAssociations"
$targetPaths = @(
    "HKCU:\Software",
    "HKLM:\Software",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts"
)
$keywords = @("WPS", "Kingsoft", "KWPS", "WPS Office", "\.wps")
$exclude = @("AMDWPS", "UWPSystem")

# 统计变量
$script:deleted = 0
$script:modified = 0
$script:skipped = 0


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

        # 跳过排除项
        foreach ($ex in $exclude) {
            if ($keyName -match $ex) {
                Write-Host "Skipped (excluded): $keyPath" -ForegroundColor Cyan
                $script:skipped++
                return
            }
        }

        # ========== 修复核心：先判断TypeOverlay是否存在 ==========
        $overlay = Get-ItemProperty -Path $keyPath -Name "TypeOverlay" -ErrorAction SilentlyContinue
        if ($overlay) {  # 仅当属性存在时才处理
            foreach ($kw in $keywords) {
                if ($overlay.TypeOverlay -match $kw) {
                    try {
                        Remove-ItemProperty -Path $keyPath -Name "TypeOverlay" -Force -ErrorAction Stop
                        Write-Host "Deleted TypeOverlay: $keyPath" -ForegroundColor Green
                        $script:deleted++
                    } catch {
                        Write-Host "Delete TypeOverlay failed: $keyPath - $_" -ForegroundColor Red
                    }
                    break  # 匹配到关键词就停止，避免重复处理
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
                    # 区分权限错误和普通错误
                    if ($_.Exception -is [System.Security.SecurityException]) {
                        Write-Host "Permission denied: $keyPath" -ForegroundColor DarkRed
                        $script:skipped++
                    } else {
                        Write-Host "Delete failed: $keyPath - $_" -ForegroundColor Red
                        # 尝试修改值
                        try {
                            $props = Get-ItemProperty -Path $keyPath -ErrorAction Stop
                            $props | Get-Member -MemberType NoteProperty | ForEach-Object {
                                $propName = $_.Name
                                $propValue = $props.$propName
                                foreach ($kw in $keywords) {
                                    if ($propValue -match $kw) {
                                        Set-ItemProperty -Path $keyPath -Name $propName -Value "" -Force -ErrorAction Stop
                                        Write-Host "Modified value: $keyPath\$propName" -ForegroundColor Yellow
                                        $script:modified++
                                        break
                                    }
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

            # 跳过排除项
            foreach ($ex in $exclude) {
                if ($keyName -match $ex) {
                    Write-Host "Skipped (excluded): $keyPath" -ForegroundColor Cyan
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
                        # 区分权限错误和普通错误
                        if ($_.Exception -is [System.Security.SecurityException]) {
                            Write-Host "Permission denied: $keyPath" -ForegroundColor DarkRed
                            $script:skipped++
                        } else {
                            Write-Host "Delete failed: $keyPath - $_" -ForegroundColor Red
                            # 尝试修改值
                            try {
                                $props = Get-ItemProperty -Path $keyPath -ErrorAction Stop
                                $props | Get-Member -MemberType NoteProperty | ForEach-Object {
                                    $propName = $_.Name
                                    $propValue = $props.$propName
                                    foreach ($kw in $keywords) {
                                        if ($propValue -match $kw) {
                                            Set-ItemProperty -Path $keyPath -Name $propName -Value "" -Force -ErrorAction Stop
                                            Write-Host "Modified value: $keyPath\$propName" -ForegroundColor Yellow
                                            $script:modified++
                                            break
                                        }
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

# 最终汇总
Write-Host "`n=== Cleanup Summary ===" -ForegroundColor Magenta
Write-Host "Deleted: $script:deleted" -ForegroundColor Green
Write-Host "Modified: $script:modified" -ForegroundColor Yellow
Write-Host "Skipped: $script:skipped" -ForegroundColor Cyan