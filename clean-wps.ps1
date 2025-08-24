<#
.SYNOPSIS
精准清理WPS残留注册表（修复类型转换+权限错误，严格对齐文章逻辑）
.DESCRIPTION
1. 仅清理文章提到的关键路径（SystemFileAssociations等）
2. 跳过非字符串类型值，避免InvalidCastException
3. 捕获权限异常，跳过系统保护项
4. 严格排除UWPSystem、AMDWPS等无关项
#>

#Requires -RunAsAdministrator

# 颜色定义
$colorSuccess = "Green"
$colorWarning = "Yellow"
$colorError = "Red"
$colorInfo = "Cyan"

# 核心配置（对齐文章逻辑）
$WPS_KEYWORDS = @("WPS", "Kingsoft", "KWPS", "WPS Office", "\.wps")  # 文章提到的关键词
$EXCLUDE_KEYWORDS = @("UWPSystem", "AMDWPS")  # 文章隐含的排除项
$CRITICAL_PATHS = @(  # 文章重点清理的路径
    "HKCR:\SystemFileAssociations",
    "HKCU:\Software\Kingsoft",
    "HKLM:\Software\Kingsoft",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts"
)

# 计数器
$deleted = 0
$modified = 0
$skipped = 0
$accessDenied = 0

# 初始化提示
Write-Host "`n=== WPS注册表清理工具（对齐文章逻辑） ===" -ForegroundColor $colorInfo
Write-Host "重点清理：SystemFileAssociations等路径 | 排除：$($EXCLUDE_KEYWORDS -join ', ')`n" -ForegroundColor $colorInfo

# 函数：检查是否排除（含文章要求的无关项）
function ShouldExclude($name) {
    foreach ($ex in $EXCLUDE_KEYWORDS) {
        if ($name -match $ex) { return $true }
    }
    return $false
}

# 主清理逻辑（聚焦文章关键路径）
foreach ($path in $CRITICAL_PATHS) {
    if (-not (Test-Path $path)) {
        Write-Host "路径不存在（跳过）: $path" -ForegroundColor $colorWarning
        $skipped++
        continue
    }

    Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $keyPath = $_.PSPath
        $keyName = $_.Name

        # 第一步：排除无关项（必做）
        if (ShouldExclude $keyName) {
            Write-Host "跳过无关项: $keyPath" -ForegroundColor $colorWarning
            $skipped++
            return
        }

        # 第二步：检查是否含WPS关键词（文章核心逻辑）
        $isWPS = $false
        foreach ($kw in $WPS_KEYWORDS) {
            if ($keyName -match $kw -or $_.GetValueNames() -contains $kw) {
                $isWPS = $true
                break
            }
        }
        if (-not $isWPS) { return }  # 非WPS项，直接跳过

        # 第三步：处理注册表项（优先删除，失败则修改）
        try {
            # 先尝试删除整个项（文章建议彻底清理）
            Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
            Write-Host "已删除项: $keyPath" -ForegroundColor $colorSuccess
            $deleted++
        }
        catch {
            # 情况1：无法删除，尝试清理值（仅处理字符串类型，避免类型错误）
            $props = Get-ItemProperty -Path $keyPath -ErrorAction SilentlyContinue
            if ($props) {
                $props | Get-Member -MemberType NoteProperty | ForEach-Object {
                    $propName = $_.Name
                    $propValue = $props.$propName

                    # 仅处理字符串类型（修复InvalidCastException）
                    if ($propValue -isnot [string]) {
                        Write-Host "跳过非字符串值: $keyPath\$propName" -ForegroundColor $colorWarning
                        $skipped++
                        return
                    }

                    # 检查值是否含WPS关键词
                    foreach ($kw in $WPS_KEYWORDS) {
                        if ($propValue -match $kw) {
                            try {
                                Remove-ItemProperty -Path $keyPath -Name $propName -Force -ErrorAction Stop
                                Write-Host "已删除值: $keyPath\$propName" -ForegroundColor $colorSuccess
                                $deleted++
                            }
                            catch [System.Security.SecurityException] {
                                # 情况2：权限不足（修复SecurityException）
                                Write-Host "权限不足（跳过）: $keyPath\$propName" -ForegroundColor $colorError
                                $accessDenied++
                            }
                            catch {
                                # 情况3：其他错误，尝试置空
                                Set-ItemProperty -Path $keyPath -Name $propName -Value "" -ErrorAction SilentlyContinue
                                Write-Host "已置空值: $keyPath\$propName" -ForegroundColor $colorWarning
                                $modified++
                            }
                            break
                        }
                    }
                }
            }
        }
    }
}

# 最终报告（对齐文章清理结果）
Write-Host "`n=== 清理完成（对齐文章逻辑） ===" -ForegroundColor $colorInfo
Write-Host "已删除项/值: $deleted" -ForegroundColor $colorSuccess
Write-Host "已置空值: $modified" -ForegroundColor $colorWarning
Write-Host "跳过无关/非字符串项: $skipped" -ForegroundColor $colorInfo
Write-Host "系统保护项（权限不足）: $accessDenied" -ForegroundColor $colorError
Write-Host "`n建议重启电脑，若图标仍异常，参考文章手动检查TypeOverlay值`n" -ForegroundColor $colorInfo