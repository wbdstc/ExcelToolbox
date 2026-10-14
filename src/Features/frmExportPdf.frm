VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExportPdf 
   Caption         =   "导出 PDF - ExcelToolbox"
   ClientHeight    =   4332
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6180
   OleObjectBlob   =   "frmExportPdf.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmExportPdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim activeWsName As String
    Dim visibleCount As Long
    
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    
    If Not wb.ActiveSheet Is Nothing Then
        activeWsName = wb.ActiveSheet.Name
    End If
    
    visibleCount = 0
    cboSheets.Clear
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            cboSheets.AddItem ws.Name
            visibleCount = visibleCount + 1
        End If
    Next ws
    
    If Len(activeWsName) > 0 Then
        On Error Resume Next
        cboSheets.Value = activeWsName
        On Error GoTo 0
    End If
    If cboSheets.ListIndex = -1 And cboSheets.ListCount > 0 Then
        cboSheets.ListIndex = 0
    End If
    
    optMerged.Caption = "合并所有可见工作表为单个 PDF (" & visibleCount & " 个)"
    optSeparate.Caption = "各可见工作表分别导出为独立 PDF (" & visibleCount & " 个)"
    
    optSingle.Value = True
    cboSheets.Enabled = True
End Sub

Private Sub optSingle_Click()
    cboSheets.Enabled = True
    chkOpenAfter.Enabled = True
End Sub

Private Sub optMerged_Click()
    cboSheets.Enabled = False
    chkOpenAfter.Enabled = True
End Sub

Private Sub optSeparate_Click()
    cboSheets.Enabled = False
    chkOpenAfter.Value = False
    chkOpenAfter.Enabled = False
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim sheetName As String
    Dim openPdf As Boolean
    
    openPdf = chkOpenAfter.Value
    
    If optSingle.Value Then
        sheetName = Trim(cboSheets.Value)
        If Len(sheetName) = 0 Then
            MsgBox "请选择要导出的工作表！", vbExclamation, "提示"
            Exit Sub
        End If
        Unload Me
        modExportPdf.ExportSingleSheetToPdf sheetName, openPdf
    ElseIf optMerged.Value Then
        Unload Me
        modExportPdf.ExportSheetsToMergedPdf openPdf
    ElseIf optSeparate.Value Then
        Unload Me
        modExportPdf.ExportSheetsToSeparatePdfs
    End If
End Sub
