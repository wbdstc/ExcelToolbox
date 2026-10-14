Attribute VB_Name = "modRibbonCallbacks"
Option Explicit

' ==============================================================================
' 模块: modRibbonCallbacks
' 功能: 功能区按钮回调与全局生命周期管理
' 作者: 开发者
' ==============================================================================

' 全局变量
Public g_RibbonUI As IRibbonUI
Public g_AppEvents As clsAppEvents

' 功能区加载完成回调
Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    Set g_RibbonUI = ribbon
    
    ' 初始化全局 Application 事件监听
    If g_AppEvents Is Nothing Then
        Set g_AppEvents = New clsAppEvents
    End If
End Sub

' 普通按钮/菜单项回调 (onAction)
Public Sub OnRibbonAction(ByVal control As IRibbonControl)
    Select Case control.Tag
        Case "ExportPdf"
            modExportPdf.ExportMultiSheetsToPdf
            
        Case "ExportPdfSingle"
            modExportPdf.ShowExportSingleSheetDialog
            
        Case "ExportPdfActive"
            modExportPdf.ExportActiveSheetToPdf
            
        Case "ExportPdfMerged"
            modExportPdf.ExportSheetsToMergedPdf
            
        Case "ExportPdfSeparate"
            modExportPdf.ExportSheetsToSeparatePdfs
            
        Case "About"
            MsgBox "【ExcelToolbox】 Excel 效率工具箱 v1.1.0" & vbCrLf & vbCrLf & _
                   "主要功能：" & vbCrLf & _
                   "  1. 导出 PDF（支持指定 Sheet 导出、当前表导出、全表合并或独立导出）" & vbCrLf & _
                   "  2. 十字聚光灯行列高亮" & vbCrLf & _
                   "  3. 双击单元格快速/切换打勾" & vbCrLf & vbCrLf & _
                   "作者: 开发者", _
                   vbInformation, "关于 ExcelToolbox"
                   
        Case Else
            MsgBox "未识别的操作标识: " & control.Id, vbExclamation, "ExcelToolbox"
    End Select
End Sub

' 状态切换按钮(ToggleButton)回调 (onAction)
Public Sub OnRibbonToggleAction(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    Select Case control.Tag
        Case "ToggleCrosshair"
            modCrosshair.SetCrosshairEnabled pressed
            
        Case "ToggleDoubleClick"
            modDoubleClick.SetDoubleClickEnabled pressed
            
        Case Else
            ' 默认忽略
    End Select
End Sub

' 获取切换按钮初始状态 (getPressed)
Public Sub GetRibbonPressed(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
    Select Case control.Tag
        Case "ToggleCrosshair"
            returnedVal = modCrosshair.IsCrosshairEnabled()
            
        Case "ToggleDoubleClick"
            returnedVal = modDoubleClick.IsDoubleClickEnabled()
            
        Case Else
            returnedVal = False
    End Select
End Sub

' 刷新指定控件 UI 状态
Public Sub InvalidateControl(ByVal controlId As String)
    If Not g_RibbonUI Is Nothing Then
        On Error Resume Next
        g_RibbonUI.InvalidateControl controlId
        On Error GoTo 0
    End If
End Sub

' 刷新所有控件 UI 状态
Public Sub InvalidateRibbon()
    If Not g_RibbonUI Is Nothing Then
        On Error Resume Next
        g_RibbonUI.Invalidate
        On Error GoTo 0
    End If
End Sub