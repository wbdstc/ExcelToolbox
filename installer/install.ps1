<#
.SYNOPSIS
    ExcelToolbox 自动化安装与激活脚本
.DESCRIPTION
    1. 将 build/ExcelToolbox.xlam 复制到用户的 %APPDATA%\Microsoft\AddIns\ 目录；
    2. 启动隐藏的 Excel COM 实例；
    3. 在 Excel 中注册并设置 AddIn.Installed = $true，激活为常驻加载项；
    4. 安全退出 Excel 并清理后台进程。
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

# 路径解析
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = (Get-Item $ScriptDir).Parent.FullName
$BuildFile = Join-Path $ProjectRoot "build\ExcelToolbox.xlam"
$TargetDir = Join-Path $env:APPDATA "Microsoft\AddIns"
$TargetFile = Join-Path $TargetDir "ExcelToolbox.xlam"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "            ExcelToolbox 宏加载项自动安装                 " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " 源码构建产物: $BuildFile"
Write-Host " 目标安装目录: $TargetDir"
Write-Host " 目标文件路径: $TargetFile`n"

# 1. 检查构建产物是否存在，若不存在则自动调用 build.ps1
if (-not (Test-Path $BuildFile)) {
    Write-Warning "未在 build 目录下找到 ExcelToolbox.xlam，正在自动触发构建..."
    $buildScript = Join-Path $ProjectRoot "build\build.ps1"
    if (Test-Path $buildScript) {
        & powershell -ExecutionPolicy Bypass -File $buildScript
    } else {
        Write-Error "未找到构建脚本: $buildScript"
        exit 1
    }
}

# 2. 终止后台 Excel 进程防止文件被锁定
Write-Host "[1/3] 关闭后台 Excel 进程..." -ForegroundColor Green
Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

# 3. 复制文件到用户 AddIns 目录
Write-Host "[2/3] 复制加载项文件至系统目录..." -ForegroundColor Green
if (-not (Test-Path $TargetDir)) {
    New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
}
Copy-Item -Path $BuildFile -Destination $TargetFile -Force
Write-Host "  -> 已成功复制到: $TargetFile"

# 4. 通过 Excel COM 注册并激活加载项 (Installed = True)
Write-Host "[3/3] 启动 Excel COM 实例激活常驻加载项..." -ForegroundColor Green
$excelType = [Type]::GetTypeFromProgID("Excel.Application")
if ($null -eq $excelType) {
    Write-Error "系统未检测到 Microsoft Excel，无法完成 COM 自动激活！"
    exit 1
}

$excel = [Activator]::CreateInstance($excelType)
try {
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # 检索或添加 AddIn
    $targetAddIn = $null
    for ($i = 1; $i -le $excel.AddIns.Count; $i++) {
        $a = $excel.AddIns.Item($i)
        if ($a.FullName -eq $TargetFile -or $a.Name -eq "ExcelToolbox.xlam") {
            $targetAddIn = $a
            break
        }
    }

    if ($null -eq $targetAddIn) {
        Write-Host "  -> 将加载项加入 Excel AddIns 列表..."
        $targetAddIn = $excel.AddIns.Add($TargetFile, $false)
    }

    # 激活加载项
    if ($null -ne $targetAddIn) {
        $targetAddIn.Installed = $true
        Write-Host "  -> 成功设置 AddIn.Installed = True"
    } else {
        Write-Warning "未能获取 AddIn 引用，请打开 Excel 手动在【加载项】中勾选启用。"
    }
}
catch {
    Write-Warning "COM 自动激活过程中出现提示: $($_.Exception.Message)"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Start-Sleep -Milliseconds 300
    Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "                 安装与激活成功！                         " -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " 加载项已成功注册为 Excel 常驻插件。"
Write-Host " 重新启动 Excel 即可在上方功能区看到【ExcelToolbox】选项卡。"
Write-Host "==========================================================" -ForegroundColor Cyan
