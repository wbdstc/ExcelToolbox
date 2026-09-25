VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' 类: clsAppEvents
' 功能: 全局 Application 事件监听类（处理双击、选区变动等）
' 作者: 开发者
' ==============================================================================

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

' 初始化
Private Sub Class_Initialize()
    Set App = Application
End Sub

' 释放
Private Sub Class_Terminate()
    Set App = Nothing
End Sub

' 双击事件 -> 转发至 modDoubleClick
Private Sub App_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    On Error Resume Next
    modDoubleClick.HandleDoubleClick Sh, Target, Cancel
    On Error GoTo 0
End Sub

' 选区变动事件 -> 转发至 modCrosshair
Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    modCrosshair.HandleSelectionChange Sh, Target
    On Error GoTo 0
End Sub
