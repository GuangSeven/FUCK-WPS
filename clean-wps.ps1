<#
.SYNOPSIS
清除WPS卸载残留注册表项，修复Office图标异常问题
.DESCRIPTION
根据 https://www.nousbuild.org/codeu/fix-office-icon-due-to-wps/ 的步骤实现，排除AMDWPS等无关项
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
$errorCount = 0  # 新增：记录错误项数量

# 配置参数
$searchKeywords = @("WPS", "\.wps", "Kingsoft", "KWPS")  # 搜索关键词(正则表达式)
$excludeKeywords = @("AMDWPS", "UWPSystem")             # 排除关键词
$targetPaths = @(
    "HKCR:\SystemFileAssociations",
    "HKCU:\Software",
    "HKLM:\Software",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts",
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\AppVMachineRegistry\Store\Integration\Backup\Software\Classes"
)

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

# 1. 清理TypeOverlay项（优先处理）
Write-Host "`n[1/2] 开始清理TypeOverlay项..." -ForegroundColor $colorInfo
# 先检查路径是否存在，避免无效操作
if (Test-Path "HKCR:\SystemFileAssociations\") {
    Get-ChildItem "HKCR:\SystemFileAssociations\" -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $keyPath = $_.PSPath
        $keyName = $_.Name
        
        if (ShouldExclude $keyName) {
            Write-Host "跳过排除项: $keyName" -ForegroundColor $colorWarning
            $script:skippedCount++
            return
        }

        # 检查并删除TypeOverlay值（添加异常捕获）
        try {
            if (Get-ItemProperty -Path $keyPath -Name "TypeOverlay" -ErrorAction Stop) {
                $overlayValue = (Get-ItemProperty -Path $keyPath).TypeOverlay
                if ($overlayValue -match "WPS Office") {
                    Remove-ItemProperty -Path $keyPath -Name "TypeOverlay" -Force -ErrorAction Stop
                    Write-Host "已删除TypeOverlay: $keyPath" -ForegroundColor $colorSuccess
                    $script:deletedCount++
                }
            }
        }
        catch {
            Write-Host "处理TypeOverlay失败: $keyPath - $_" -ForegroundColor $colorError
            $script:errorCount++
        }
    }
}
else {
    Write-Host "路径不存在: HKCR:\SystemFileAssociations\" -ForegroundColor $colorWarning
}

# 2. 清理其他相关注册表项
Write-Host "`n[2/2] 开始清理其他WPS相关项..." -ForegroundColor $colorInfo
foreach ($keyword in $searchKeywords) {
    Write-Host "`n搜索关键词: $keyword" -ForegroundColor $colorInfo
    foreach ($path in $targetPaths) {
        if (-not (Test-Path $path)) {
            Write-Host "路径不存在: $path" -ForegroundColor $colorWarning
            continue
        }

        Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $keyPath = $_.PSPath
            $keyName = $_.Name

            # 检查排除项
            if (ShouldExclude $keyName) {
                return
            }

            # 检查项名是否匹配关键词（添加异常捕获）
            try {
                if ($keyName -match $keyword) {
                    Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                    Write-Host "已删除项: $keyPath" -ForegroundColor $colorSuccess
                    $script:deletedCount++
                }
            }
            catch {
                # 尝试修改默认值为空白
                try {
                    if (Get-ItemProperty -Path $keyPath -Name "(Default)" -ErrorAction Stop) {
                        $defaultValue = (Get-ItemProperty -Path $keyPath)."(Default)"
                        if ($defaultValue -match $keyword) {
                            Set-ItemProperty -Path $keyPath -Name "(Default)" -Value "" -Force -ErrorAction Stop
                            Write-Host "已修改默认值: $keyPath" -ForegroundColor $colorWarning
                            $script:modifiedCount++
                        }
                    }
                }
                catch {
                    Write-Host "处理项失败: $keyPath - $_" -ForegroundColor $colorError
                    $script:errorCount++
                }
            }

            # 检查值是否匹配关键词（核心修复：添加try/catch捕获类型转换错误）
            try {
                # 尝试读取属性，若失败则直接跳过
                $props = Get-ItemProperty -Path $keyPath -ErrorAction Stop
                if ($props) {
                    $props | Get-Member -MemberType NoteProperty -ErrorAction SilentlyContinue | ForEach-Object {
                        $propName = $_.Name
                        if ($propName -eq "(Default)") { return }
                        
                        $propValue = $props.$propName
                        if ($propValue -match $keyword -and -not (ShouldExclude $propValue)) {
                            try {
                                Remove-ItemProperty -Path $keyPath -Name $propName -Force -ErrorAction Stop
                                Write-Host "已删除值: $keyPath\$propName" -ForegroundColor $colorSuccess
                                $script:deletedCount++
                            }
                            catch {
                                Set-ItemProperty -Path $keyPath -Name $propName -Value "" -Force -ErrorAction Stop
                                Write-Host "已修改值: $keyPath\$propName" -ForegroundColor $colorWarning
                                $script:modifiedCount++
                            }
                        }
                    }
                }
            }
            catch [System.InvalidCastException] {
                # 专门捕获类型转换错误，仅记录不中断
                Write-Host "跳过无法解析的项（类型转换错误）: $keyPath" -ForegroundColor $colorWarning
                $script:skippedCount++
            }
            catch {
                # 其他错误
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
Write-Host "跳过排除项数量: $skippedCount" -ForegroundColor $colorInfo
Write-Host "错误项数量: $errorCount" -ForegroundColor $colorError
Write-Host "`n建议重启电脑使更改生效！`n" -ForegroundColor $colorInfo
