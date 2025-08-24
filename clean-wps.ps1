<#
.SYNOPSIS
清除WPS卸载残留注册表项，修复Office图标异常问题
.DESCRIPTION
针对路径HKLM:\SOFTWARE\Classes\SystemFileAssociations优化，解决路径找不到和类型转换错误
#>

# 检查管理员权限
#Requires -RunAsAdministrator

# 颜色定义
$colorSuccess = "Green"
$colorWarning = "Yellow"
$colorError = "Red"
$colorInfo = "Cyan"

# 初始化计数器
$deletedCount = 0
$modifiedCount = 0
$skippedCount = 0
$errorCount = 0

# 核心修复：更新目标路径，优先包含你提供的正确路径
$targetPaths = @(
    "HKLM:\SOFTWARE\Classes\SystemFileAssociations",  # 你指定的核心路径
    "HKCU:\Software",
    "HKLM:\Software",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts",
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\AppVMachineRegistry\Store\Integration\Backup\Software\Classes"
)

# 搜索关键词（保持不变）
$searchKeywords = @("WPS", "\.wps", "Kingsoft", "KWPS")
$excludeKeywords = @("AMDWPS", "UWPSystem")

# 警告信息
Write-Host "`n=== WPS注册表残留清理工具 ===" -ForegroundColor $colorInfo
Write-Host "注意：此操作将修改注册表，请先备份重要数据！" -ForegroundColor $colorWarning
Write-Host "正在准备清理，请稍候...`n" -ForegroundColor $colorInfo

# 函数：检查是否需要排除
function ShouldExclude($name) {
    foreach ($ex in $excludeKeywords) {
        if ($name -match $ex) {
            return $true
        }
    }
    return $false
}

# 1. 优先清理指定路径下的TypeOverlay项（核心修复）
$targetTypeOverlayPath = "HKLM:\SOFTWARE\Classes\SystemFileAssociations"
Write-Host "`n[1/2] 开始清理指定路径下的TypeOverlay项：$targetTypeOverlayPath" -ForegroundColor $colorInfo

if (Test-Path $targetTypeOverlayPath) {
    Get-ChildItem $targetTypeOverlayPath -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $keyPath = $_.PSPath
        $keyName = $_.Name
        
        if (ShouldExclude $keyName) {
            Write-Host "跳过排除项: $keyName" -ForegroundColor $colorWarning
            $script:skippedCount++
            return
        }

        # 读取TypeOverlay时避免类型转换错误
        try {
            $prop = Get-ItemProperty -Path $keyPath -Name "TypeOverlay" -ErrorAction Stop
            if ($prop.TypeOverlay -match "WPS Office") {
                Remove-ItemProperty -Path $keyPath -Name "TypeOverlay" -Force -ErrorAction Stop
                Write-Host "已删除TypeOverlay: $keyPath" -ForegroundColor $colorSuccess
                $script:deletedCount++
            }
        }
        catch [System.InvalidCastException] {
            Write-Host "跳过类型异常项（TypeOverlay）: $keyPath" -ForegroundColor $colorWarning
            $script:skippedCount++
        }
        catch {
            Write-Host "删除TypeOverlay失败: $keyPath - $_" -ForegroundColor $colorError
            $script:errorCount++
        }
    }
}
else {
    Write-Host "指定路径不存在: $targetTypeOverlayPath，请手动检查注册表" -ForegroundColor $colorError
    $script:errorCount++
}

# 2. 清理所有目标路径下的其他WPS相关项
Write-Host "`n[2/2] 开始清理其他WPS相关项..." -ForegroundColor $colorInfo
foreach ($keyword in $searchKeywords) {
    Write-Host "`n搜索关键词: $keyword" -ForegroundColor $colorInfo
    foreach ($path in $targetPaths) {
        if (-not (Test-Path $path)) {
            Write-Host "路径不存在: $path（已跳过）" -ForegroundColor $colorWarning
            $script:skippedCount++
            continue
        }

        Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $keyPath = $_.PSPath
            $keyName = $_.Name

            if (ShouldExclude $keyName) { return }

            # 处理注册表项
            try {
                if ($keyName -match $keyword) {
                    Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                    Write-Host "已删除项: $keyPath" -ForegroundColor $colorSuccess
                    $script:deletedCount++
                }
            }
            catch {
                try {
                    $defaultProp = Get-ItemProperty -Path $keyPath -Name "(Default)" -ErrorAction Stop
                    if ($defaultProp."(Default)" -match $keyword) {
                        Set-ItemProperty -Path $keyPath -Name "(Default)" -Value "" -Force -ErrorAction Stop
                        Write-Host "已修改默认值: $keyPath" -ForegroundColor $colorWarning
                        $script:modifiedCount++
                    }
                }
                catch {
                    Write-Host "处理项失败: $keyPath - $_" -ForegroundColor $colorError
                    $script:errorCount++
                }
            }

            # 处理注册表值（彻底解决类型转换问题）
            try {
                $props = Get-ItemProperty -Path $keyPath -ErrorAction Stop
                if ($props) {
                    # 只处理字符串类型的属性（避免二进制等特殊类型）
                    $stringProps = $props | Get-Member -MemberType NoteProperty | Where-Object {
                        $propValue = $props.$($_.Name)
                        $propValue -is [string]  # 仅保留字符串类型
                    }

                    $stringProps | ForEach-Object {
                        $propName = $_.Name
                        if ($propName -eq "(Default)") { return }
                        
                        $propValue = $props.$propName
                        if ($propValue -match $keyword -and -not (ShouldExclude $propValue)) {
                            Remove-ItemProperty -Path $keyPath -Name $propName -Force -ErrorAction Stop
                            Write-Host "已删除值: $keyPath\$propName" -ForegroundColor $colorSuccess
                            $script:deletedCount++
                        }
                    }
                }
            }
            catch [System.InvalidCastException] {
                Write-Host "跳过类型异常项（属性）: $keyPath" -ForegroundColor $colorWarning
                $script:skippedCount++
            }
            catch {
                Write-Host "处理属性失败: $keyPath - $_" -ForegroundColor $colorError
                $script:errorCount++
            }
        }
    }
}

# 清理完成报告
Write-Host "`n`n=== 清理完成 ===" -ForegroundColor $colorInfo
Write-Host "已删除项数量: $deletedCount" -ForegroundColor $colorSuccess
Write-Host "已修改值数量: $modifiedCount" -ForegroundColor $colorWarning
Write-Host "跳过项数量: $skippedCount" -ForegroundColor $colorInfo
Write-Host "错误项数量: $errorCount" -ForegroundColor $colorError
Write-Host "`n建议重启电脑使更改生效！`n" -ForegroundColor $colorInfo
