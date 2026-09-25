# ExcelToolbox (Excel 效率工具箱)

> 作者: **王不动**

基于 VBA 与 Office Fluent Ribbon XML 构建的高效 Excel 宏加载项工具箱。

## 目录结构

```text
ExcelToolbox/
├── src/
│   ├── Ribbon/                  # 存放功能区 XML 布局文件
│   │   └── customUI14.xml       # Ribbon 界面定义
│   ├── Core/                    # 核心控制与事件调度模块
│   │   ├── clsAppEvents.cls     # 全局 Application 级别事件捕获（监听双击、选择变动）
│   │   └── modRibbonCallbacks.bas # 功能区按钮点击回调调度器
│   └── Features/                # 业务功能模块（每个小功能一个独立文件）
│       ├── modExportPdf.bas     # 功能1：多 Sheet 批量导出 PDF
│       ├── modCrosshair.bas     # 功能2：十字聚光灯行列高亮
│       └── modDoubleClick.bas   # 功能3：双击填充“√”
├── build/                       # 打包生成物与构建脚本目录
│   ├── build.ps1                # 构建脚本
│   └── README.md
├── installer/                   # 安装包制作与部署脚本
│   ├── install.ps1              # 一键安装脚本
│   ├── uninstall.ps1            # 卸载脚本
│   └── README.md
└── README.md
```

## 功能特性

1. **批量导出 PDF (`modExportPdf.bas`)**:
   - 一键将当前工作簿中的所有可见工作表分别导出为独立的 PDF 文件。
   - 自动过滤文件名中的非法字符，支持自定义保存目录。

2. **十字聚光灯 (`modCrosshair.bas`)**:
   - 动态高亮显示当前选中单元格所在的行和列，便于在大表格中快速定位与核对数据。

3. **双击快速打勾 (`modDoubleClick.bas`)**:
   - 双击单元格快速切换/填入“√”符号，并自动拦截默认编辑状态，极大提升数据复核效率。

## 开发与部署流程

1. **修改与添加功能**：
   - 在 `src/Features/` 中新增或编辑业务模块。
   - 在 `src/Ribbon/customUI14.xml` 中添加按钮，并在 `src/Core/modRibbonCallbacks.bas` 中补充回调处理。
2. **构建与打包**：
   - 运行 `.\build\build.ps1` 校验与打包生成 `ExcelToolbox.xlam`。
3. **安装加载项**：
   - 运行 `.\installer\install.ps1` 一键部署至系统 Excel 加载项目录。
