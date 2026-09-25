<#
.SYNOPSIS
    ExcelToolbox 自动化卸载与注销脚本
.DESCRIPTION
    1. 启动 Excel COM 实例取消激活加载项 (Installed = $false)；
    2. 从系统加载项目录移除 ExcelToolbox.xlam。
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

$TargetFile = Join-Path $env:APPDATA "Microsoft\AddIns\ExcelToolbox.xlam"

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "            ExcelToolbox 宏加载项卸载向导                 " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. 关闭后台 Excel
Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

# 2. 通过 Excel COM 取消激活
$excelType = [Type]::GetTypeFromProgID("Excel.Application")
if ($null -ne $excelType) {
    $excel = [Activator]::CreateInstance($excelType)
    try {
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        for ($i = 1; $i -le $excel.AddIns.Count; $i++) {
            $a = $excel.AddIns.Item($i)
            if ($a.FullName -eq $TargetFile -or $a.Name -eq "ExcelToolbox.xlam") {
                $a.Installed = $false
                Write-Host "  -> 已取消 Excel 注册状态 (Installed = False)"
                break
            }
        }
    }
    catch {
        Write-Warning "注销时捕获提示: $($_.Exception.Message)"
    }
    finally {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Start-Sleep -Milliseconds 300
        Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

# 3. 删除文件
if (Test-Path $TargetFile) {
    Remove-Item -Path $TargetFile -Force
    Write-Host "  -> 已成功删除系统加载项文件: $TargetFile"
} else {
    Write-Host "  -> 系统加载项目录中未检测到残留文件。"
}

Write-Host "`n卸载完成！" -ForegroundColor Green
