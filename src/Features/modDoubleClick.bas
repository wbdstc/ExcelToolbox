Attribute VB_Name = "modDoubleClick"
Option Explicit

' ==============================================================================
' 模块: modDoubleClick
' 功能: 双击单元格快速打勾/切换打勾
' 作者: 开发者
' ==============================================================================

Private p_DoubleClickEnabled As Boolean
Private Const CHECK_SYMBOL As String = "√"

' 查询双击功能是否开启
Public Function IsDoubleClickEnabled() As Boolean
    IsDoubleClickEnabled = p_DoubleClickEnabled
End Function

' 设置双击功能开启/关闭
Public Sub SetDoubleClickEnabled(ByVal enableState As Boolean)
    p_DoubleClickEnabled = enableState
    ' 同步刷新 Ribbon 按钮状态
    modRibbonCallbacks.InvalidateControl "btnToggleDoubleClick"
End Sub

' 切换功能状态
Public Sub ToggleDoubleClick()
    SetDoubleClickEnabled Not p_DoubleClickEnabled
End Sub

' 双击事件核心处理
Public Sub HandleDoubleClick(ByVal Sh As Object, ByVal Target As Range, ByRef Cancel As Boolean)
    If Not p_DoubleClickEnabled Then Exit Sub
    If Target Is Nothing Then Exit Sub
    
    ' 仅对单单元格有效，防止多选区域误操作
    If Target.Cells.CountLarge > 1 Then Exit Sub
    
    Dim cellText As String
    cellText = Trim(CStr(Target.Value2))
    
    ' 切换逻辑:
    ' 1. 当前为“√”时，清空
    ' 2. 当前为空或其他内容时，置为“√”
    If cellText = CHECK_SYMBOL Then
        Target.Value2 = vbNullString
    Else
        Target.Value2 = CHECK_SYMBOL
    End If
    
    ' 拦截 Excel 默认的双击进入单元格编辑状态行为
    Cancel = True
End Sub
