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
        Write-Host "Error: Path not found - $priorityPath"
        return
    }
    Write-Host "Processing priority path: $priorityPath"
    Get-ChildItem $priorityPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $keyPath = $_.PSPath
        Write-Host "Handling: $keyPath"  # 输出当前处理的项
        
        $keyName = $_.Name

        # 跳过排除项
        foreach ($ex in $exclude) {
            if ($keyName -match $ex) {
                $script:skipped++
                return
            }
        }

        # 优先删除 TypeOverlay
        try {
            $overlay = Get-ItemProperty -Path $keyPath -Name "TypeOverlay" -ErrorAction Stop
            foreach ($kw in $keywords) {
                if ($overlay.TypeOverlay -match $kw) {
                    Remove-ItemProperty -Path $keyPath -Name "TypeOverlay" -Force -ErrorAction Stop
                    $script:deleted++
                    break
                }
            }
        } catch {}

        # 处理注册表项
        foreach ($kw in $keywords) {
            if ($keyName -match $kw) {
                try {
                    Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                    $script:deleted++
                } catch {
                    try {
                        $props = Get-ItemProperty -Path $keyPath -ErrorAction Stop
                        $props | Get-Member -MemberType NoteProperty | ForEach-Object {
                            $propName = $_.Name
                            $propValue = $props.$propName
                            foreach ($kw in $keywords) {
                                if ($propValue -match $kw) {
                                    Set-ItemProperty -Path $keyPath -Name $propName -Value "" -Force -ErrorAction Stop
                                    $script:modified++
                                    break
                                }
                            }
                        }
                    } catch { 
                        $script:skipped++ 
                    }
                }
                break
            }
        }
    }
}


### 2. 清理其他路径
function Remove-WPSTargets {
    foreach ($path in $targetPaths) {
        if (-not (Test-Path $path)) {
            Write-Host "Error: Path not found - $path"
            continue
        }
        Write-Host "Processing path: $path"
        Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $keyPath = $_.PSPath
            Write-Host "Handling: $keyPath"  # 输出当前处理的项
            
            $keyName = $_.Name

            # 跳过排除项
            foreach ($ex in $exclude) {
                if ($keyName -match $ex) {
                    $script:skipped++
                    return
                }
            }

            # 处理注册表项
            foreach ($kw in $keywords) {
                if ($keyName -match $kw) {
                    try {
                        Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                        $script:deleted++
                    } catch {
                        try {
                            $props = Get-ItemProperty -Path $keyPath -ErrorAction Stop
                            $props | Get-Member -MemberType NoteProperty | ForEach-Object {
                                $propName = $_.Name
                                $propValue = $props.$propName
                                foreach ($kw in $keywords) {
                                    if ($propValue -match $kw) {
                                        Set-ItemProperty -Path $keyPath -Name $propName -Value "" -Force -ErrorAction Stop
                                        $script:modified++
                                        break
                                    }
                                }
                            }
                        } catch { 
                            $script:skipped++ 
                        }
                    }
                    break
                }
            }
        }
    }
}


### 执行清理
Write-Host "Starting WPS registry cleanup..."  # 开始运行提示
Remove-WPSPriority
Remove-WPSTargets

# 输出结果
Write-Host "`nCleanup complete."
Write-Host "Deleted: $script:deleted"
Write-Host "Modified: $script:modified"
Write-Host "Skipped: $script:skipped"
