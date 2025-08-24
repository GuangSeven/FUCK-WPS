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
$deleted = 0
$modified = 0
$skipped = 0

# 优先清理指定路径
function Remove-WPSPriorityRegistry {
    if (-not (Test-Path $priorityPath)) {
        Write-Host "Path not found: $priorityPath"
        return
    }
    Get-ChildItem $priorityPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $key = $_.PSPath
        $name = $_.Name

        # 检查排除项
        foreach ($ex in $exclude) {
            if ($name -match $ex) {
                $script:skipped++
                return
            }
        }

        # 处理TypeOverlay（优先清理图标相关）
        try {
            $overlay = Get-ItemProperty -Path $key -Name "TypeOverlay" -ErrorAction Stop
            foreach ($kw in $keywords) {
                if ($overlay.TypeOverlay -match $kw) {
                    Remove-ItemProperty -Path $key -Name "TypeOverlay" -Force -ErrorAction Stop
                    $script:deleted++
                    break
                }
            }
        } catch {}

        # 处理注册表项
        foreach ($kw in $keywords) {
            if ($name -match $kw) {
                try {
                    Remove-Item -Path $key -Recurse -Force -ErrorAction Stop
                    $script:deleted++
                } catch {
                    # 处理注册表值（仅字符串类型）
                    try {
                        $props = Get-ItemProperty -Path $key -ErrorAction Stop
                        $props | Get-Member -MemberType NoteProperty | ForEach-Object {
                            $prop = $_.Name
                            $val = $props.$prop
                            foreach ($kw in $keywords) {
                                if ($val -match $kw) {
                                    Set-ItemProperty -Path $key -Name $prop -Value "" -Force -ErrorAction Stop
                                    $script:modified++
                                    break
                                }
                            }
                        }
                    } catch [System.InvalidCastException] { $script:skipped++ }
                      catch [System.Security.SecurityException] { $script:skipped++ }
                }
                break
            }
        }
    }
}

# 清理其他目标路径
function Remove-WPSTargetRegistry {
    foreach ($path in $targetPaths) {
        if (-not (Test-Path $path)) {
            Write-Host "Path not found: $path"
            continue
        }
        Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $key = $_.PSPath
            $name = $_.Name

            # 检查排除项
            foreach ($ex in $exclude) {
                if ($name -match $ex) {
                    $script:skipped++
                    return
                }
            }

            # 处理注册表项
            foreach ($kw in $keywords) {
                if ($name -match $kw) {
                    try {
                        Remove-Item -Path $key -Recurse -Force -ErrorAction Stop
                        $script:deleted++
                    } catch {
                        # 处理注册表值（仅字符串类型）
                        try {
                            $props = Get-ItemProperty -Path $key -ErrorAction Stop
                            $props | Get-Member -MemberType NoteProperty | ForEach-Object {
                                $prop = $_.Name
                                $val = $props.$prop
                                foreach ($kw in $keywords) {
                                    if ($val -match $kw) {
                                        Set-ItemProperty -Path $key -Name $prop -Value "" -Force -ErrorAction Stop
                                        $script:modified++
                                        break
                                    }
                                }
                            }
                        } catch { $script:skipped++ }
                    }
                    break
                }
            }
        }
    }
}

# 执行清理
Clean-Priority
Clean-Targets

# 输出结果
Write-Host "Deleted: $deleted"
Write-Host "Modified: $modified"
Write-Host "Skipped: $skipped"