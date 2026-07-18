# FUCK-WPS - WPS 残留注册表清理工具

基于 PowerShell 7+ / .NET RegistryKey API 的高性能注册表清理工具，通过 `config.json` 配置关键词、排除模式、扫描路径（支持通配符）。

## 功能特性

- **高性能**：.NET API 直接操作注册表，比原生命令快数倍
- **配置驱动**：`config.json` 定义关键词、排除、优先/常规扫描路径、通配符
- **通配符**：`*`(任意)、`{*}`(GUID)、`Prefix*`、`*.suffix`
- **优先路径**：文件关联、CLSID 等高频残留优先扫描
- **自动备份**：运行前导出 `.reg` 备份（`-NoBackup` 禁用）
- **演练模式**：`-WhatIf` 预览操作不修改
- **进度条/静默/详细日志**：`-Quiet` / `-Verbose` / `-NoProgress`
- **图标缓存刷新**：可选自动 `ie4uinit.exe -ClearIconCache`
- **统计报告**：扫描/删除/修改/跳过/错误计数
- **安全中断**：`Ctrl+C` 汇总后退出

## 要求

- Windows 10/11 (64位)
- **PowerShell 7+**（`winget install Microsoft.PowerShell`）
- **管理员权限**

## 使用

```powershell
# 下载 Clean-WPS.ps1 和 config.json 到同一目录
cd D:\Tools

# 默认：静默 + 自动备份 + 自动刷新图标
pwsh -File .\Clean-WPS.ps1

# 详细日志
pwsh -File .\Clean-WPS.ps1 -Verbose

# 演练模式（不修改注册表）
pwsh -File .\Clean-WPS.ps1 -WhatIf -Verbose

# 无人值守
pwsh -File .\Clean-WPS.ps1 -Quiet

# 自定义配置/日志，禁用备份/图标刷新
pwsh -File .\Clean-WPS.ps1 -ConfigFile "C:\cfg.json" -LogFile "C:\log.txt" -NoBackup -NoRefreshIcons
```

## 常用参数

| 参数 | 说明 |
|------|------|
| `-ConfigFile` | 配置文件路径（默认脚本目录 `config.json`） |
| `-WhatIf` | 演练模式，只显示将执行的操作 |
| `-NoBackup` | 禁用自动 `.reg` 备份 |
| `-LogFile` | 指定日志文件路径 |
| `-Verbose` | 显示详细操作日志 |
| `-Quiet` | 静默模式，仅进度条+最终汇总 |
| `-NoRefreshIcons` | 跳过图标缓存刷新 |
| `-NoProgress` | 禁用进度条 |

## 配置文件

`config.json` 主要字段：
- `keywords` - 匹配键名/值的正则关键词（不区分大小写）
- `excludePatterns` - 排除模式，匹配则跳过
- `priorityPaths` - 优先扫描路径（支持通配符）
- `registryPaths.HKLM/HKCU` - 常规扫描路径（支持通配符）
- `settings.autoBackup` - 自动导出 `.reg` 备份
- `settings.autoRefreshIcons` - 完成后自动刷新图标缓存
- `settings.confirmRefreshIcons` - 刷新前确认（静默模式下无效）

## 注意事项

1. **运行前请手动备份注册表**（`regedit → 文件 → 导出`），脚本自动备份仅作补充
2. 关闭所有 Office/WPS/资源管理器预览窗口
3. **必须用 PowerShell 7+**，5.1 不支持 `.NET RegistryKey` API
4. 完成后检查日志（`Clean-WPS-*.log`），如有误删立即导入备份 `.reg`
5. 建议重启或重启 `explorer.exe` 使图标生效

## 常见问题

| 问题 | 解决 |
|------|------|
| "未对脚本进行数字签名" | `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` |
| 提示需管理员 | 右键终端 → "以管理员身份运行" |
| 图标仍异常 | `ie4uinit.exe -ClearIconCache` 或重建 `IconCache.db` |
| 误删重要项 | 双击脚本目录下 `WPS-RegBackup-*.reg` 恢复 |

## 贡献指南

欢迎提交 Issue 和 PR：
- **Bug 报告**：包含系统版本、PowerShell 版本 (`$PSVersionTable`)、错误日志
- **功能建议**：在 Issue 中描述需求
- **代码贡献**：Fork 后提交 PR，保持兼容 PowerShell 7+ 和 Windows 10/11

详见 [CONTRIBUTING.md](CONTRIBUTING.md)

## 许可证

[MIT License](LICENSE) - 允许自由使用、修改、分发，需保留版权声明。

## 免责声明

本工具仅用于清理 WPS 残留注册表项。使用前请务必备份注册表。因不当使用、配置错误或系统差异导致的任何问题，开发者不承担责任。

## 致谢

- 参考 [NousBuild 技术文章](https://www.nousbuild.org/codeu/fix-office-icon-due-to-wps/) 的清理思路
- 感谢所有提交 Issue 和 PR 的贡献者