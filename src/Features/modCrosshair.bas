Attribute VB_Name = "modCrosshair"
Option Explicit

' ==============================================================================
' 模块: modCrosshair
' 功能: 十字聚光灯行列高亮（基于 FormatConditions 条件格式实现）
' 作者: 开发者
' 特点:
'   - 不修改单元格原始 Interior.Color，完全无损，保留用户底色
'   - 动态注入/移除专用条件格式规则，不破坏已有格式
'   - 通过 CELL("row") 和 CELL("col") 动态实现极致十字跟随
' ==============================================================================

Private p_CrosshairEnabled As Boolean

' 十字高亮核心公式（VBA 内部使用标准英文逗号函数）
Private Const CF_FORMULA As String = "=OR(ROW()=CELL(""row""),COLUMN()=CELL(""col""))"
' 柔和的高亮颜色 (RGB: 225, 238, 255 - 浅天蓝)
Private Const HIGHLIGHT_RGB As Long = 16772833

' 查询聚光灯当前是否开启
Public Function IsCrosshairEnabled() As Boolean
    IsCrosshairEnabled = p_CrosshairEnabled
End Function

' 设置聚光灯开启/关闭
Public Sub SetCrosshairEnabled(ByVal enableState As Boolean)
    p_CrosshairEnabled = enableState
    
    If Not p_CrosshairEnabled Then
        ' 关闭时清除当前所有工作表中的聚光灯条件格式
        RemoveCrosshairFromAllSheets
    Else
        ' 开启时立即为当前活动工作表应用聚光灯规则并刷新
        If Not ActiveSheet Is Nothing Then
            If TypeName(ActiveSheet) = "Worksheet" Then
                ApplyCrosshairCondition ActiveSheet
                ActiveSheet.Calculate
            End If
        End If
    End If
    
    ' 刷新工具栏 ToggleButton 状态
    modRibbonCallbacks.InvalidateControl "btnToggleCrosshair"
End Sub

' 切换聚光灯
Public Sub ToggleCrosshair()
    SetCrosshairEnabled Not p_CrosshairEnabled
End Sub

' 选区变动调度接口（由 clsAppEvents 调用）
Public Sub HandleSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Not p_CrosshairEnabled Then Exit Sub
    If Sh Is Nothing Or Target Is Nothing Then Exit Sub
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    
    On Error Resume Next
    ' 确保当前表具备聚光灯条件格式
    ApplyCrosshairCondition Sh
    
    ' 强制重算以刷新 CELL("row") 和 CELL("col")
    Application.ScreenUpdating = False
    Sh.Calculate
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' 为指定工作表添加或确认聚光灯条件格式
Private Sub ApplyCrosshairCondition(ByVal ws As Worksheet)
    On Error Resume Next
    
    ' 保护表跳过
    If ws.ProtectContents Then Exit Sub
    
    ' 检查是否已经存在聚光灯条件格式
    Dim fc As FormatCondition
    Dim i As Long
    Dim exists As Boolean
    exists = False
    
    For i = 1 To ws.Cells.FormatConditions.Count
        Set fc = ws.Cells.FormatConditions(i)
        If fc.Type = 2 Then ' 2 = xlExpression
            If InStr(1, fc.Formula1, "CELL(""row"")", vbTextCompare) > 0 Or _
               InStr(1, fc.Formula1, "CELL(""col"")", vbTextCompare) > 0 Then
                exists = True
                Exit For
            End If
        End If
    Next i
    
    ' 不存在时动态添加一条条件格式
    If Not exists Then
        Set fc = ws.Cells.FormatConditions.Add(Type:=2, Formula1:=CF_FORMULA)
        With fc
            .Interior.Pattern = 1 ' xlSolid
            .Interior.PatternColorIndex = -4105 ' xlAutomatic
            .Interior.Color = HIGHLIGHT_RGB
            .StopIfTrue = False ' 不中断后续格式
        End With
    End If
    On Error GoTo 0
End Sub

' 移除指定工作表中的聚光灯条件格式（精确匹配，不误删用户自定义格式）
Private Sub RemoveCrosshairCondition(ByVal ws As Worksheet)
    On Error Resume Next
    If ws.ProtectContents Then Exit Sub
    
    Dim i As Long
    Dim fc As FormatCondition
    
    ' 倒序遍历删除
    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        Set fc = ws.Cells.FormatConditions(i)
        If fc.Type = 2 Then ' 2 = xlExpression
            If InStr(1, fc.Formula1, "CELL(""row"")", vbTextCompare) > 0 Or _
               InStr(1, fc.Formula1, "CELL(""col"")", vbTextCompare) > 0 Then
                fc.Delete
            End If
        End If
    Next i
    
    ' 刷新视图
    ws.Calculate
    On Error GoTo 0
End Sub

' 移除当前打开的所有工作簿所有工作表中的聚光灯条件格式
Private Sub RemoveCrosshairFromAllSheets()
    On Error Resume Next
    Dim wb As Workbook
    Dim ws As Worksheet
    
    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            RemoveCrosshairCondition ws
        Next ws
    Next wb
    On Error GoTo 0
End Sub
