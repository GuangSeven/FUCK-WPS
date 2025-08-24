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

# 统计变量（script作用域，确保函数内可修改）
$script:deleted = 0
$script:modified = 0
$script:skipped = 0


### 1. 优先清理指定路径（核心路径）
function Remove-WPSPriority {
    if (-not (Test-Path $priorityPath)) {
        Write-Host "Error: Path not found - $priorityPath"
        return
    }
    Get-ChildItem $priorityPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $keyPath = $_.PSPath
        $keyName = $_.Name

        # 跳过排除项
        foreach ($ex in $exclude) {
            if ($keyName -match $ex) {
                $script:skipped++
                return
            }
        }

        # 优先删除 TypeOverlay（图标异常核心项）
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

        # 删除 WPS 相关项，失败则置空值
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


### 2. 清理文章提到的其他路径
function Remove-WPSTargets {
    foreach ($path in $targetPaths) {
        if (-not (Test-Path $path)) {
            Write-Host "Error: Path not found - $path"
            continue
        }
        Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $keyPath = $_.PSPath
            $keyName = $_.Name

            # 跳过排除项
            foreach ($ex in $exclude) {
                if ($keyName -match $ex) {
                    $script:skipped++
                    return
                }
            }

            # 删除 WPS 相关项，失败则置空值
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


### 执行清理 + 输出结果
Remove-WPSPriority  # 调用优先清理函数
Remove-WPSTargets   # 调用其他路径清理函数

Write-Host "Deleted: $script:deleted"
Write-Host "Modified: $script:modified"
Write-Host "Skipped: $script:skipped"