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
│       ├── frmExportPdf.frm     # 功能1交互窗体：工作表选择与导出模式面板
│       ├── frmExportPdf.frx     # 窗体二进制资源文件
│       ├── modExportPdf.bas     # 功能1：PDF 导出逻辑（单表/合并/独立导出）
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

1. **导出为 PDF (`modExportPdf.bas` + `frmExportPdf.frm`)**:
   - **选择导出指定工作表**：提供交互界面下拉选择任意指定工作表导出为独立 PDF，默认预选当前活动工作表。
   - **一键导出当前活动表**：支持直接将正在编辑/查看的工作表快速导出为 PDF。
   - **全表合并导出**：将所有可见工作表按顺序合并导出为一个完整的 PDF 文件。
   - **全表分别导出**：将所有可见工作表分别导出为独立的 PDF 文件，支持自定义保存目录。
   - **贴心细节体验**：支持导出后自动打开 PDF、自动过滤文件名非法字符、隐藏表防误触提示与状态还原。

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
