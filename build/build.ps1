<#
.SYNOPSIS
    ExcelToolbox 自动化构建与打包脚本
.DESCRIPTION
    1. 关闭后台运行的 Excel 进程并配置安全权限；
    2. 通过 Excel COM 自动化创建 AddIn 工作簿，导入 src/ 下所有 VBA 模块（.cls / .bas）；
    3. 将工作簿保存为标准 .xlam 宏加载项；
    4. 对生成的 .xlam（OpenXML 压缩包）注入 customUI/customUI14.xml 并在 _rels/.rels 中配置关系节点；
    5. 执行自动化验证并输出构建报告。
#>

[CmdletBinding()]
param(
    [string]$OutputDir,
    [string]$OutputFileName = "ExcelToolbox.xlam"
)

$ErrorActionPreference = "Stop"

# 路径计算
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $OutputDir) {
    $OutputDir = $ScriptDir
}

$ProjectRoot = (Get-Item $ScriptDir).Parent.FullName
$SrcDir = Join-Path $ProjectRoot "src"
$RibbonXml = Join-Path $SrcDir "Ribbon\customUI14.xml"
$TargetPath = Join-Path $OutputDir $OutputFileName

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "            ExcelToolbox 宏加载项自动化构建               " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " 项目根目录: $ProjectRoot"
Write-Host " 源码目录:   $SrcDir"
Write-Host " 输出文件:   $TargetPath`n"

# 1. 检查基础环境与源码文件
if (-not (Test-Path $RibbonXml)) {
    Write-Error "未找到 Ribbon XML 配置文件: $RibbonXml"
    exit 1
}

$vbaFiles = Get-ChildItem -Path $SrcDir -Recurse -Include *.bas, *.cls, *.frm
if ($vbaFiles.Count -eq 0) {
    Write-Error "源码目录中未找到任何 VBA 模块文件 (.bas / .cls / .frm)！"
    exit 1
}

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# 2. 关闭后台 Excel 进程以避免文件占用
Write-Host "[1/5] 关闭后台 Excel 进程并准备环境..." -ForegroundColor Green
Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

# 配置注册表：信任对 VBA 工程对象模型的访问 (AccessVBOM)
$officeVersions = @("16.0", "15.0", "14.0")
foreach ($ver in $officeVersions) {
    $secPath = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (-not (Test-Path $secPath)) {
        New-Item -Path $secPath -Force -ErrorAction SilentlyContinue | Out-Null
    }
    Set-ItemProperty -Path $secPath -Name "AccessVBOM" -Value 1 -Type DWord -ErrorAction SilentlyContinue | Out-Null
}

# 3. 规范化 VBA 源文件行尾为 CRLF 与 GBK 编码（确保 Excel VBIDE 正确解析 Class 属性）
Write-Host "[2/5] 检查并规范化 VBA 源码格式 (CRLF / GBK)..." -ForegroundColor Green
$gbkEnc = [System.Text.Encoding]::GetEncoding("GBK")
foreach ($file in $vbaFiles) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $text = $null
    try {
        $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
        $text = $utf8Strict.GetString($bytes)
    } catch {
        $text = $gbkEnc.GetString($bytes)
    }
    
    # 统一规范为 Windows CRLF
    $crlfText = $text.Replace("`r`n", "`n").Replace("`n", "`r`n")
    [System.IO.File]::WriteAllText($file.FullName, $crlfText, $gbkEnc)
    Write-Host "  - 模块格式已就绪: $($file.Name)"
}

# 4. 调用 Excel COM 接口编译生成基础 .xlam
Write-Host "[3/5] 启动 Excel COM 自动化并导入 VBA 模块..." -ForegroundColor Green
$excelType = [Type]::GetTypeFromProgID("Excel.Application")
if ($null -eq $excelType) {
    Write-Error "系统未检测到 Excel.Application COM 组件，请确认已安装 Microsoft Excel！"
    exit 1
}

$excel = [Activator]::CreateInstance($excelType)
try {
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $wb = $excel.Workbooks.Add()
    $vbProj = $wb.VBProject

    # 导入所有 VBA 模块
    foreach ($file in $vbaFiles) {
        Write-Host "  -> 导入模块: $($file.Name)"
        $importedComp = $vbProj.VBComponents.Import($file.FullName)
        if ($null -eq $importedComp) {
            Write-Warning "导入模块 $($file.Name) 返回空引用，请检查语法。"
        }
    }

    # 删除已有输出文件
    if (Test-Path $TargetPath) {
        Remove-Item $TargetPath -Force
    }

    # 保存为 55 = xlOpenXMLAddIn (.xlam)
    Write-Host "  -> 保存为 Excel 加载项: $TargetPath"
    $wb.SaveAs($TargetPath, 55)
    $wb.Close($false)
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    # 再次清理可能残余的进程
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

# 5. 向 OpenXML 压缩包中注入 customUI14.xml 与配置 _rels/.rels 关系
Write-Host "[4/5] 注入 Ribbon 界面 (customUI14.xml) 到 OpenXML 包..." -ForegroundColor Green
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$zip = [System.IO.Compression.ZipFile]::Open($TargetPath, [System.IO.Compression.ZipArchiveMode]::Update)
try {
    # 5.1 注入 customUI/customUI14.xml
    $customUiEntry = $zip.GetEntry("customUI/customUI14.xml")
    if ($null -ne $customUiEntry) {
        $customUiEntry.Delete()
    }
    
    $newCustomUiEntry = $zip.CreateEntry("customUI/customUI14.xml", [System.IO.Compression.CompressionLevel]::Optimal)
    $entryStream = $newCustomUiEntry.Open()
    $writer = New-Object System.IO.StreamWriter($entryStream, [System.Text.Encoding]::UTF8)
    $xmlContent = [System.IO.File]::ReadAllText($RibbonXml, [System.Text.Encoding]::UTF8)
    $writer.Write($xmlContent)
    $writer.Flush()
    $writer.Close()
    $entryStream.Close()
    Write-Host "  -> 成功写入: customUI/customUI14.xml"

    # 5.2 更新 _rels/.rels
    $relsEntry = $zip.GetEntry("_rels/.rels")
    if ($null -eq $relsEntry) {
        throw "未能在 .xlam 包中找到 _rels/.rels 关系配置文件！"
    }

    $rStream = $relsEntry.Open()
    $rReader = New-Object System.IO.StreamReader($rStream)
    $relsContent = $rReader.ReadToEnd()
    $rReader.Close()
    $rStream.Close()

    $relType = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
    $relTarget = "customUI/customUI14.xml"

    [xml]$relsXml = $relsContent
    $ns = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
    $ns.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")

    $existingRel = $relsXml.SelectSingleNode("//r:Relationship[@Target='$relTarget']", $ns)
    if ($null -eq $existingRel) {
        $existingRel = $relsXml.SelectSingleNode("//r:Relationship[@Type='$relType']", $ns)
    }

    if ($null -eq $existingRel) {
        # 计算下一个可用的 rId 序号
        $maxId = 1
        foreach ($node in $relsXml.SelectNodes("//r:Relationship", $ns)) {
            $idStr = $node.GetAttribute("Id")
            if ($idStr -match "^rId(\d+)$") {
                $num = [int]$matches[1]
                if ($num -ge $maxId) { $maxId = $num + 1 }
            }
        }
        $newId = "rId$maxId"

        $newRel = $relsXml.CreateElement("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships")
        $newRel.SetAttribute("Id", $newId)
        $newRel.SetAttribute("Type", $relType)
        $newRel.SetAttribute("Target", $relTarget)
        $relsXml.DocumentElement.AppendChild($newRel) | Out-Null
        Write-Host "  -> 成功在 _rels/.rels 中关联 Ribbon: Id=$newId, Target=$relTarget"
    } else {
        Write-Host "  -> Ribbon 关系已存在: Target=$($existingRel.GetAttribute('Target'))"
    }

    # 重新写回 _rels/.rels
    $relsEntry.Delete()
    $newRelsEntry = $zip.CreateEntry("_rels/.rels", [System.IO.Compression.CompressionLevel]::Optimal)
    $relsStream = $newRelsEntry.Open()
    $relsXml.Save($relsStream)
    $relsStream.Close()
}
finally {
    $zip.Dispose()
}

# 6. 验证最终生成的 .xlam 文件
Write-Host "[5/5] 验证生成物与模块完整性..." -ForegroundColor Green
$verifyExcel = [Activator]::CreateInstance($excelType)
try {
    $verifyExcel.Visible = $false
    $verifyExcel.DisplayAlerts = $false
    $verifyWb = $verifyExcel.Workbooks.Open($TargetPath)
    
    $compNames = @()
    foreach ($comp in $verifyWb.VBProject.VBComponents) {
        $typeName = switch ($comp.Type) {
            1 { "标准模块 (StdModule)" }
            2 { "类模块 (ClassModule)" }
            3 { "窗体 (UserForm)" }
            100 { "文档对象 (Document)" }
            default { "未知类型" }
        }
        $compNames += "      * $($comp.Name) -> $typeName"
    }
    $verifyWb.Close($false)
}
finally {
    $verifyExcel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($verifyExcel) | Out-Null
}

$fileInfo = Get-Item $TargetPath
Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "                   构建成功完成！                         " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " 生成文件: $($fileInfo.FullName)"
Write-Host " 文件大小: $($fileInfo.Length) 字节"
Write-Host " 修改时间: $($fileInfo.LastWriteTime)"
Write-Host " 包含组件:"
foreach ($line in $compNames) {
    Write-Host $line
}
Write-Host "==========================================================" -ForegroundColor Cyan
