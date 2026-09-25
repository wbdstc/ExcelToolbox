# Installer 目录

本目录包含 ExcelToolbox 的部署与安装脚本。

## 脚本列表
- **`install.ps1`**: 将编译好的 `build\ExcelToolbox.xlam` 自动复制到系统的 Excel 加载项目录（`%APPDATA%\Microsoft\AddIns`）。
- **`uninstall.ps1`**: 从系统加载项目录中移除 `ExcelToolbox.xlam`。

## 安装步骤
1. 打开 PowerShell
2. 执行安装：
   ```powershell
   .\installer\install.ps1
   ```
3. 打开 Excel，在【开发工具】->【Excel 加载项】中勾选启用。
