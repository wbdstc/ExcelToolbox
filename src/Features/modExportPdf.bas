Attribute VB_Name = "modExportPdf"
Option Explicit

' ==============================================================================
' 模块: modExportPdf
' 功能: 导出工作表为 PDF（支持选择导出指定 Sheet、当前活动 Sheet、合并可见 Sheet 或分别导出）
' 作者: 开发者
' ==============================================================================

' 主入口函数：弹出交互面板供用户选择导出模式及具体工作表
Public Sub ExportMultiSheetsToPdf()
    ShowExportDialog
End Sub

' 弹出导出主对话框
Public Sub ShowExportDialog()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "当前没有打开的活动工作簿！", vbExclamation, "提示"
        Exit Sub
    End If
    
    On Error Resume Next
    frmExportPdf.Show
    If Err.Number <> 0 Then
        ' 降级容错方案：若窗体因环境原因无法显示，转入交互式选择
        Dim errDesc As String
        errDesc = Err.Description
        Err.Clear
        On Error GoTo 0
        FallbackExportDialog wb
    End If
    On Error GoTo 0
End Sub

' 快捷打开指定工作表导出对话框
Public Sub ShowExportSingleSheetDialog()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "当前没有打开的活动工作簿！", vbExclamation, "提示"
        Exit Sub
    End If
    
    On Error Resume Next
    Load frmExportPdf
    frmExportPdf.optSingle.Value = True
    frmExportPdf.cboSheets.Enabled = True
    frmExportPdf.Show
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        FallbackExportDialog wb
    End If
    On Error GoTo 0
End Sub

' 导出当前活动工作表为独立 PDF
Public Sub ExportActiveSheetToPdf()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "当前没有打开的活动工作簿！", vbExclamation, "提示"
        Exit Sub
    End If
    
    If wb.ActiveSheet Is Nothing Then
        MsgBox "当前没有活动工作表！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ExportSingleSheetToPdf wb.ActiveSheet.Name, False
End Sub

' 导出指定名称的工作表为独立 PDF
Public Sub ExportSingleSheetToPdf(Optional ByVal targetSheetName As String = "", Optional ByVal openAfter As Boolean = False)
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "当前没有打开的活动工作簿！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 若未传入目标工作表名称，则让用户选择
    If Len(Trim(targetSheetName)) = 0 Then
        ShowExportSingleSheetDialog
        Exit Sub
    End If
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(targetSheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "未找到名为【" & targetSheetName & "】的工作表！", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 检查工作表是否隐藏
    Dim origVisible As Long
    Dim wasHidden As Boolean
    wasHidden = False
    If ws.Visible <> xlSheetVisible Then
        Dim unhideChoice As VbMsgBoxResult
        unhideChoice = MsgBox("工作表【" & ws.Name & "】当前处于隐藏状态，是否临时取消隐藏并导出？", _
                              vbYesNo + vbQuestion, "工作表已隐藏")
        If unhideChoice = vbNo Then Exit Sub
        origVisible = ws.Visible
        wasHidden = True
        ws.Visible = xlSheetVisible
    End If
    
    ' 构造默认保存文件名和路径
    Dim safeSheetName As String
    Dim baseName As String
    Dim defaultFileName As String
    Dim defaultFolder As String
    Dim defaultFullPath As String
    Dim varFilePath As Variant
    Dim filePath As String
    
    safeSheetName = SanitizeFileName(ws.Name)
    baseName = GetBaseName(wb.Name)
    defaultFileName = baseName & "_" & safeSheetName & ".pdf"
    
    If Len(wb.Path) > 0 Then
        defaultFolder = wb.Path & Application.PathSeparator
    Else
        defaultFolder = Application.DefaultFilePath & Application.PathSeparator
    End If
    defaultFullPath = defaultFolder & defaultFileName
    
    varFilePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultFullPath, _
        FileFilter:="PDF 文件 (*.pdf), *.pdf", _
        FilterIndex:=1, _
        Title:="请选择【" & ws.Name & "】导出 PDF 的保存路径与名称")
        
    ' 用户取消操作
    If VarType(varFilePath) = vbBoolean Then
        If varFilePath = False Then
            If wasHidden Then ws.Visible = origVisible
            Exit Sub
        End If
    End If
    
    filePath = CStr(varFilePath)
    If Len(Trim(filePath)) = 0 Or filePath = "False" Then
        If wasHidden Then ws.Visible = origVisible
        Exit Sub
    End If
    
    If LCase(Right(filePath, 4)) <> ".pdf" Then
        filePath = filePath & ".pdf"
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error GoTo SingleExportError
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=openAfter
        
    If wasHidden Then ws.Visible = origVisible
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "导出成功！" & vbCrLf & _
           "工作表【" & ws.Name & "】已成功导出为 PDF 文件。" & vbCrLf & _
           "保存路径: " & filePath, vbInformation, "导出成功"
    Exit Sub

SingleExportError:
    If wasHidden Then ws.Visible = origVisible
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "导出工作表【" & ws.Name & "】失败: " & Err.Description, vbCritical, "导出错误"
End Sub

' 模式：合并所有可见工作表导出为一个完整的 PDF 文件
Public Sub ExportSheetsToMergedPdf(Optional ByVal openAfter As Boolean = False)
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "当前没有打开的活动工作簿！", vbExclamation, "提示"
        Exit Sub
    End If
    
    Dim visibleNames() As String
    Dim visibleCount As Long
    visibleCount = GetVisibleWorksheetNames(wb, visibleNames)
    
    If visibleCount = 0 Then
        MsgBox "当前工作簿中没有可见的工作表！", vbExclamation, "提示"
        Exit Sub
    End If
    
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
    
    varFilePath = Application.GetSaveAsFilename( _
        InitialFileName:=defaultFullPath, _
        FileFilter:="PDF 文件 (*.pdf), *.pdf", _
        FilterIndex:=1, _
        Title:="请选择合并后 PDF 文件的保存路径与名称")
        
    If VarType(varFilePath) = vbBoolean Then
        If varFilePath = False Then Exit Sub
    End If
    
    filePath = CStr(varFilePath)
    If Len(Trim(filePath)) = 0 Or filePath = "False" Then
        Exit Sub
    End If
    
    If LCase(Right(filePath, 4)) <> ".pdf" Then
        filePath = filePath & ".pdf"
    End If
    
    Dim origActiveSheet As Object
    Set origActiveSheet = wb.ActiveSheet
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    On Error GoTo MergedError
    wb.Worksheets(visibleNames).Select
    
    wb.ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=openAfter
        
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

' 模式：各个可见工作表分别导出为独立的 PDF 文件
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
    
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "请选择单独导出 PDF 的保存文件夹"
        If Len(wb.Path) > 0 Then
            .InitialFileName = wb.Path & Application.PathSeparator
        End If
        If .Show = -1 Then
            exportFolder = .SelectedItems(1)
        Else
            Exit Sub
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

' 降级交互选择框（当用户窗体不可用时的备用方案）
Private Sub FallbackExportDialog(ByVal wb As Workbook)
    Dim choice As VbMsgBoxResult
    choice = MsgBox("请选择 PDF 导出模式：" & vbCrLf & vbCrLf & _
                    "【是 (Yes)】  ：导出指定/当前工作表为独立 PDF" & vbCrLf & _
                    "【否 (No)】   ：合并所有可见工作表为一个 PDF" & vbCrLf & _
                    "【取消 (Cancel)】：取消并退出", _
                    vbYesNoCancel + vbQuestion + vbDefaultButton1, "导出 PDF - 选择导出模式")
                    
    Select Case choice
        Case vbYes
            Dim sheetList As String
            Dim ws As Worksheet
            Dim i As Long
            i = 1
            sheetList = "请输入要导出的工作表序号或名称：" & vbCrLf & vbCrLf
            For Each ws In wb.Worksheets
                If ws.Visible = xlSheetVisible Then
                    sheetList = sheetList & i & ". " & ws.Name
                    If ws.Name = wb.ActiveSheet.Name Then
                        sheetList = sheetList & " (当前活动表)"
                    End If
                    sheetList = sheetList & vbCrLf
                    i = i + 1
                End If
            Next ws
            
            Dim inputVal As Variant
            inputVal = Application.InputBox(sheetList, "选择工作表导出 PDF", wb.ActiveSheet.Name, Type:=2)
            If VarType(inputVal) = vbBoolean Then
                If inputVal = False Then Exit Sub
            End If
            Dim strVal As String
            strVal = Trim(CStr(inputVal))
            If Len(strVal) = 0 Then Exit Sub
            
            Dim targetName As String
            targetName = ""
            If IsNumeric(strVal) Then
                Dim idx As Long
                idx = CLng(strVal)
                Dim curIdx As Long
                curIdx = 1
                For Each ws In wb.Worksheets
                    If ws.Visible = xlSheetVisible Then
                        If curIdx = idx Then
                            targetName = ws.Name
                            Exit For
                        End If
                        curIdx = curIdx + 1
                    End If
                Next ws
            Else
                targetName = strVal
            End If
            
            If Len(targetName) > 0 Then
                ExportSingleSheetToPdf targetName, False
            Else
                MsgBox "未找到指定序号/名称的工作表: " & strVal, vbExclamation, "提示"
            End If
            
        Case vbNo
            ExportSheetsToMergedPdf False
            
        Case vbCancel
            Exit Sub
    End Select
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