Attribute VB_Name = "modExportPdf"
Option Explicit

' ==============================================================================
' 模块: modExportPdf
' 功能: 批量导出工作表为 PDF（支持合并导出与各 Sheet 独立导出）
' 作者: 开发者
' ==============================================================================

' 主入口函数：弹出交互选择或默认分流
Public Sub ExportMultiSheetsToPdf()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "当前没有打开的活动工作簿！", vbExclamation, "提示"
        Exit Sub
    End If
    
    Dim choice As VbMsgBoxResult
    choice = MsgBox("请选择 PDF 导出模式：" & vbCrLf & vbCrLf & _
                    "【是 (Yes)】  ：合并所有可见 Sheet 导出为一个 PDF 文件" & vbCrLf & _
                    "【否 (No)】   ：每个可见 Sheet 分别导出为独立 PDF 文件" & vbCrLf & _
                    "【取消 (Cancel)】：取消并退出", _
                    vbYesNoCancel + vbQuestion + vbDefaultButton1, "批量导出 PDF - 选择导出模式")
                    
    Select Case choice
        Case vbYes
            ExportSheetsToMergedPdf
        Case vbNo
            ExportSheetsToSeparatePdfs
        Case vbCancel
            Exit Sub
    End Select
End Sub

' 模式 1：合并所有可见工作表导出为一个完整的 PDF 文件
Public Sub ExportSheetsToMergedPdf()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "当前没有打开的活动工作簿！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 获取所有可见工作表名称
    Dim visibleNames() As String
    Dim visibleCount As Long
    visibleCount = GetVisibleWorksheetNames(wb, visibleNames)
    
    If visibleCount = 0 Then
        MsgBox "当前工作簿中没有可见的工作表！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 默认保存文件名和路径
    Dim defaultFileName As String
    Dim defaultFolder As String
    Dim defaultFullPath As String
    Dim varFilePath As Variant
    Dim filePath As String
    
    defaultFileName = GetBaseName(wb.Name) & ".pdf"
    If Len(wb.Path) > 0 Then
        defaultFolder = wb.Path & Application.PathSeparator
    Else
        defaultFolder = Application.DefaultFilePath & Application.PathSeparator
    End If
    defaultFullPath = defaultFolder & defaultFileName
    
    ' 使用 GetSaveAsFilename 锁定保存类型为 PDF (*.pdf)，避免出现 Excel 工作簿类型
    varFilePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultFullPath, _
        FileFilter:="PDF 文件 (*.pdf), *.pdf", _
        FilterIndex:=1, _
        Title:="请选择合并后 PDF 文件的保存路径与名称")
        
    ' 用户取消操作
    If VarType(varFilePath) = vbBoolean Then
        If varFilePath = False Then Exit Sub
    End If
    
    filePath = CStr(varFilePath)
    If Len(Trim(filePath)) = 0 Or filePath = "False" Then
        Exit Sub
    End If
    
    ' 确保后缀为 .pdf
    If LCase(Right(filePath, 4)) <> ".pdf" Then
        filePath = filePath & ".pdf"
    End If
    
    ' 记住当前选中的 Sheet 以便恢复
    Dim origActiveSheet As Object
    Set origActiveSheet = wb.ActiveSheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error GoTo MergedError
    ' 选中所有可见工作表（多选组合）
    wb.Worksheets(visibleNames).Select
    
    ' 导出多选工作表为一个合并 PDF
    wb.ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
    ' 恢复原活动工作表
    If Not origActiveSheet Is Nothing Then
        origActiveSheet.Select
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "合并导出成功！" & vbCrLf & _
           "已将 " & visibleCount & " 个可见工作表合并为一个 PDF 文件。" & vbCrLf & _
           "文件路径: " & filePath, vbInformation, "导出成功"
    Exit Sub
    
MergedError:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not origActiveSheet Is Nothing Then
        On Error Resume Next
        origActiveSheet.Select
        On Error GoTo 0
    End If
    MsgBox "合并导出 PDF 失败: " & Err.Description, vbCritical, "导出错误"
End Sub

' 模式 2：各个可见工作表分别导出为独立的 PDF 文件
Public Sub ExportSheetsToSeparatePdfs()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim exportFolder As String
    Dim exportCount As Long
    Dim filePath As String
    Dim safeSheetName As String
    Dim fDialog As FileDialog
    
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "当前没有打开的活动工作簿！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 文件夹选择对话框
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "请选择单独导出 PDF 的保存文件夹"
        If Len(wb.Path) > 0 Then
            .InitialFileName = wb.Path & Application.PathSeparator
        End If
        If .Show = -1 Then
            exportFolder = .SelectedItems(1)
        Else
            Exit Sub ' 用户取消选择
        End If
    End With
    
    If Right(exportFolder, 1) <> Application.PathSeparator Then
        exportFolder = exportFolder & Application.PathSeparator
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    exportCount = 0
    Dim baseName As String
    baseName = GetBaseName(wb.Name)
    
    For Each ws In wb.Worksheets
        ' 仅导出可见工作表
        If ws.Visible = xlSheetVisible Then
            safeSheetName = SanitizeFileName(ws.Name)
            filePath = exportFolder & baseName & "_" & safeSheetName & ".pdf"
            
            On Error Resume Next
            ws.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=filePath, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
                
            If Err.Number = 0 Then
                exportCount = exportCount + 1
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next ws
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "独立导出完成！" & vbCrLf & _
           "成功将 " & exportCount & " 个工作表分别导出为独立 PDF。" & vbCrLf & _
           "保存目录: " & exportFolder, vbInformation, "导出成功"
End Sub

' 获取工作簿中所有可见工作表的名称数组
Private Function GetVisibleWorksheetNames(ByVal wb As Workbook, ByRef outNames() As String) As Long
    Dim ws As Worksheet
    Dim cnt As Long
    cnt = 0
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            ReDim Preserve outNames(0 To cnt)
            outNames(cnt) = ws.Name
            cnt = cnt + 1
        End If
    Next ws
    GetVisibleWorksheetNames = cnt
End Function

' 过滤文件名中的非法字符
Private Function SanitizeFileName(ByVal rawName As String) As String
    Dim invalidChars As Variant
    Dim i As Long
    invalidChars = Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
    For i = LBound(invalidChars) To UBound(invalidChars)
        rawName = Replace(rawName, invalidChars(i), "_")
    Next i
    SanitizeFileName = Trim(rawName)
End Function

' 获取去掉扩展名的纯文件名
Private Function GetBaseName(ByVal fileName As String) As String
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 1 Then
        GetBaseName = Left(fileName, pos - 1)
    Else
        GetBaseName = fileName
    End If
End Function
