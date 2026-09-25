# Build 目录

本目录用于存放构建脚本及编译生成的最终加载项文件（`.xlam` / `.zip`）。

## 构建产物说明
- **`ExcelToolbox.xlam`**: 编译打包后的 Excel 宏加载项文件。
- **`build.ps1`**: 自动化构建脚本，负责将 `src/` 源码模块与 `src/Ribbon/customUI14.xml` 注入并打包为标准 `.xlam` 加载项。

## 构建方式
在 PowerShell 中运行：
```powershell
.\build\build.ps1
```
