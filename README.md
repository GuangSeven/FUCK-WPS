# FUCK-WPS - 清理 WPS 残留注册表项工具

一款用于彻底清理 WPS Office 卸载后残留注册表项的 PowerShell 脚本，可解决因 WPS 残留导致的 Office 图标异常、文件关联错乱等问题。

## 项目简介

因为主包电脑之前预装了 Office 2019 和 WPS 卸载 WPS 后发现 Office 图标出现很多异常所以决定写一个脚本来修复这些问题。

当卸载 WPS Office 后，系统中常残留大量注册表项，可能导致 Microsoft Office 图标显示异常（如空白图标、图标错乱）或文件关联冲突。本工具基于 [NousBuild 技术文章](https://www.nousbuild.org/codeu/fix-office-icon-due-to-wps/) 的清理步骤开发，通过精准扫描并删除 WPS 相关残留注册表项，帮助恢复系统正常的文件关联和图标显示。

## 功能特点

* **精准清理**：针对性扫描并删除 WPS（Kingsoft、KWPS）相关注册表项及值

* **彩色反馈**：实时显示操作状态（绿色：成功删除 / 黄色：修改项 / 红色：错误 / 青色：提示）

* **安全处理**：无法删除的项会尝试修改为空白值，避免强制删除导致的系统风险

* **统计报告**：清理完成后生成详细统计（删除 / 修改 / 跳过项数量）

## 使用步骤

### 前置条件

* 操作系统：Windows 10 / 11（64 位）

* 运行环境：PowerShell 5.1+（系统默认预装，无需额外安装）

* 权限要求：必须以**管理员身份**运行（注册表操作需要管理员权限）

### 操作流程

1. **下载脚本**

    从仓库 Releases 页面下载最新版本的 `Clean-WPS.ps1` 脚本，保存到本地（如 `D:\Tools\` 目录）

2. **启动 PowerShell**

    * 按下 `Win + X` 组合键，选择「Windows PowerShell (管理员)」

    * 若提示 “用户账户控制”，点击「是」授予管理员权限

3. **执行脚本**

    在 PowerShell 中输入以下命令并回车（替换实际脚本路径）：

    ```PowerShell
    cd D:\Tools\  # 切换到脚本所在目录

    .\Clean-WPS.ps1  # 运行脚本
    ```

4. **查看结果**

    脚本运行完成后会显示清理统计，建议重启电脑使更改生效。

## 兼容性说明

| 系统版本            | 支持情况    | 备注                            |
| --------------- | ------- | ----------------------------- |
| Windows 11      | ✅ 完全支持  | 默认预装 PowerShell 5.1，无需额外配置    |
| Windows 10      | ✅ 完全支持  | 需确保已升级到 PowerShell 5.1        |
| Windows 7       | ⚠️ 有限支持 | 需手动升级 PowerShell 5.1，可能存在路径差异 |
| PowerShell <5.1 | ❌ 不支持   | 低版本存在功能缺陷，可能导致运行失败            |

## 注意事项

1. **注册表备份**：运行前建议通过「注册表编辑器 → 文件 → 导出」备份关键注册表分支（如 `HKCR`、`HKLM\SOFTWARE`）(一定一定要备份！！！)

2. **程序关闭**：清理前请关闭所有 Office 和 WPS 相关程序，避免文件锁定导致清理不完整

3. **风险提示**：注册表修改可能影响系统稳定性，非专业用户建议在技术人员指导下操作

4. **结果验证**：重启后若 Office 图标仍异常，可尝试重建图标缓存（`ie4uinit.exe -ClearIconCache`）

5. **看日志**：翻一遍日志或者直接用 Ctrl + F 搜索一下有没有包含和WPS不相关的注册表文件，如果有的话马上导入你之前备份的注册表文件，接着回到项目页反馈问题

## 常见问题

### Q：脚本运行时报 “权限不足” 错误？

A：请确保已以管理员身份启动 PowerShell（右键 PowerShell 图标选择「以管理员身份运行」）

### Q：清理后 Office 图标仍异常怎么办？

A：尝试手动重建图标缓存：

```PowerShell
\# 在管理员PowerShell中执行

taskkill /f /im explorer.exe

del /f /q %userprofile%\AppData\Local\IconCache.db

start explorer.exe
```

### Q：脚本会删除我的 Office 注册表项吗？

A：不会。脚本仅针对含 "WPS"、"Kingsoft"、"KWPS" 关键词的项，且会自动排除 Office 原生路径

### Q：Win7 系统如何使用？

A：需先手动升级 PowerShell 5.1（[下载微软官方升级包](https://www.microsoft.com/en-us/download/details.aspx?id=54616)），但可能存在兼容性问题

## 贡献指南

欢迎通过以下方式参与项目改进：

* 报告 Bug：通过 Issues 提交，需包含系统版本、PowerShell 版本和错误截图

* 功能建议：在 Issues 中提出新功能需求或改进建议

* 代码贡献：Fork 仓库后提交 PR，修改需兼容 Windows 10/11 和 PowerShell 5.1+

详细贡献规则请参考 [CONTRIBUTING.md](CONTRIBUTING.md)

## 许可证

本项目采用 [MIT License](LICENSE) 开源协议，允许自由使用、修改和分发，需保留原版权声明。

## 免责声明

本工具仅用于清理 WPS 残留注册表项，使用前请务必备份数据。因不当使用或系统差异导致的问题，开发者不承担责任。
