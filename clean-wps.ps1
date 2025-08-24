<#
.SYNOPSIS
按关键词精准清理WPS相关注册表项（排除AMDWPS等无关项）
.DESCRIPTION
- 搜索"WPS"：删除所有相关项（排除UWPSystem、AMDWPS等）
- 搜索".wps"：删除所有相关注册表文件夹
- 搜索"Kingsoft"：删除所有相关注册表文件夹
- 搜索"WPS Office"：删除相关项，无法删除则置空值
- 搜索"KWPS"：删除相关项，无法删除则置空值
#>

# 检查管理员权限
#Requires -RunAsAdministrator

# 颜色定义
$colorSuccess = "Green"
$colorWarning = "Yellow"
$colorError = "Red"
$colorInfo = "Cyan"

# 计数器
$deletedCount = 0
$modifiedCount = 0
$skippedCount = 0  # 记录被排除的无关项数量

# 核心配置：关键词及对应操作
$keywordActions = @(
    @{ Keyword = "WPS"; Action = "DeleteAll"; Description = "删除所有含WPS的项（排除无关项）" },
    @{ Keyword = "\.wps"; Action = "DeleteFolder"; Description = "删除.wps相关注册表文件夹" },
    @{ Keyword = "Kingsoft"; Action = "DeleteFolder"; Description = "删除Kingsoft相关注册表文件夹" },
    @{ Keyword = "WPS Office"; Action = "DeleteOrEmpty"; Description = "删除WPS Office相关项，失败则置空" },
    @{ Keyword = "KWPS"; Action = "DeleteOrEmpty"; Description = "删除KWPS相关项，失败则置空" }
)

# 排除列表（含WPS但不相关的项，已添加AMDWPS）
$excludeKeywords = @("UWPSystem", "AMDWPS")  # 关键修改：新增AMDWPS排除规则

# 目标注册表路径
$targetPaths = @(
    "HKLM:\SOFTWARE\Classes\SystemFileAssociations",
    "HKCU:\Software",
    "HKLM:\Software",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FileExts",
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\AppVMachineRegistry\Store\Integration\Backup\Software\Classes"
)

# 初始化提示
Write-Host "`n=== WPS注册表清理工具 ===" -ForegroundColor $colorInfo
Write-Host "已排除无关项：UWPSystem、AMDWPS`n" -ForegroundColor $colorInfo

# 函数：检查是否需要排除（含WPS但无关的项）
function ShouldExclude($name) {
    foreach ($ex in $excludeKeywords) {
        if ($name -match $ex) {
            return $true  # 匹配到排除关键词则返回需要排除
        }
    }
    return $false
}

# 执行清理逻辑（完整处理流程）
foreach ($action in $keywordActions) {
    $keyword = $action.Keyword
    $actionType = $action.Action
    $description = $action.Description

    Write-Host "`n=== 处理关键词：$keyword（$description） ===" -ForegroundColor $colorInfo

    foreach ($path in $targetPaths) {
        if (-not (Test-Path $path)) {
            Write-Host "路径不存在（已跳过）：$path" -ForegroundColor $colorWarning
            continue
        }

        Get-ChildItem $path -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $keyPath = $_.PSPath
            $keyName = $_.Name

            # 先检查是否属于排除项（如AMDWPS）
            if (ShouldExclude $keyName) {
                Write-Host "已跳过无关项：$keyPath" -ForegroundColor $colorWarning
                $script:skippedCount++
                return
            }

            # 检查项名是否匹配当前关键词
            if ($keyName -match $keyword) {
                switch ($actionType) {
                    "DeleteAll" {
                        # 删除所有相关项
                        try {
                            Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                            Write-Host "已删除项：$keyPath" -ForegroundColor $colorSuccess
                            $script:deletedCount++
                        }
                        catch {
                            Write-Host "删除项失败：$keyPath - $_" -ForegroundColor $colorError
                        }
                    }
                    "DeleteFolder" {
                        # 删除相关文件夹（注册表项）
                        try {
                            Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                            Write-Host "已删除文件夹：$keyPath" -ForegroundColor $colorSuccess
                            $script:deletedCount++
                        }
                        catch {
                            Write-Host "删除文件夹失败：$keyPath - $_" -ForegroundColor $colorError
                        }
                    }
                    "DeleteOrEmpty" {
                        # 删除项，失败则置空值
                        try {
                            Remove-Item -Path $keyPath -Recurse -Force -ErrorAction Stop
                            Write-Host "已删除项：$keyPath" -ForegroundColor $colorSuccess
                            $script:deletedCount++
                        }
                        catch {
                            # 尝试置空默认值
                            try {
                                if (Get-ItemProperty -Path $keyPath -Name "(Default)" -ErrorAction Stop) {
                                    Set-ItemProperty -Path $keyPath -Name "(Default)" -Value "" -Force -ErrorAction Stop
                                    Write-Host "已置空值：$keyPath" -ForegroundColor $colorWarning
                                    $script:modifiedCount++
                                }
                            }
                            catch {
                                Write-Host "处理项失败：$keyPath - $_" -ForegroundColor $colorError
                            }
                        }
                    }
                }
                return  # 已处理整个项，无需检查内部值
            }

            # 检查项内属性值是否匹配关键词
            $props = Get-ItemProperty -Path $keyPath -ErrorAction SilentlyContinue
            if ($props) {
                $props | Get-Member -MemberType NoteProperty -ErrorAction SilentlyContinue | ForEach-Object {
                    $propName = $_.Name
                    if ($propName -eq "(Default)") { return }

                    $propValue = $props.$propName
                    if ($propValue -match $keyword) {
                        # 处理属性值匹配的情况
                        try {
                            Remove-ItemProperty -Path $keyPath -Name $propName -Force -ErrorAction Stop
                            Write-Host "已删除值：$keyPath\$propName" -ForegroundColor $colorSuccess
                            $script:deletedCount++
                        }
                        catch {
                            Set-ItemProperty -Path $keyPath -Name $propName -Value "" -Force -ErrorAction Stop
                            Write-Host "已置空值：$keyPath\$propName" -ForegroundColor $colorWarning
                            $script:modifiedCount++
                        }
                    }
                }
            }
        }
    }
}

# 清理完成报告
Write-Host "`n`n=== 清理完成 ===" -ForegroundColor $colorInfo
Write-Host "已删除项/值数量：$deletedCount" -ForegroundColor $colorSuccess
Write-Host "已置空值数量：$modifiedCount" -ForegroundColor $colorWarning
Write-Host "跳过的无关项数量（含AMDWPS等）：$skippedCount" -ForegroundColor $colorInfo
Write-Host "`n建议重启电脑使更改生效！`n" -ForegroundColor $colorInfo
