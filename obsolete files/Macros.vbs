Sub CEP()
    Dim fd As FileDialog, vrtSelectedItem As Variant
    'Dim iFile As Document
    Dim wdApp As Word.Application
    Dim iFile As Word.Document
    sht = ActiveSheet.Name

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    With fd
        .AllowMultiSelect = False
        .InitialFileName = ActiveWorkbook.Path
        .Filters.Add "Documents", "*.doc; *.docx; *.rtf", 1
        .FilterIndex = 2
        .Title = "150kHz～30MHz电源端子骚扰电压"
        If .Show <> -1 Then
            MsgBox "未选择文件", vbCritical
            Exit Sub
        Else
            'UserForm1.Show
            For Each vrtSelectedItem In .SelectedItems
                wdApp.Documents.Open vrtSelectedItem
                wdApp.Activate
                wdApp.Documents(1).Tables(1).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(5, 1).Select
                ActiveSheet.Paste
                
                wdApp.Activate
                wdApp.Documents(1).Tables(2).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(16, 1).Select
                ActiveSheet.Paste
                
                ActiveWorkbook.Activate
                On Error Resume Next
                ActiveWorkbook.Sheets("Y-CEP").Pictures(1).Delete
                
                wdApp.Activate
                wdApp.Documents(1).Paragraphs(3).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("Y-CEP").Select
                ActiveSheet.Cells(21, 2).Select
                ActiveSheet.Paste
                
                Set pic_range = Range("B21:F27")
                ActiveSheet.Pictures(1).Select
                With Selection.ShapeRange
                    .Top = pic_range.Top
                    .Height = pic_range.Height
                End With
                ActiveSheet.Pictures(1).Copy
                
                On Error Resume Next
                ActiveWorkbook.Sheets("B-CEP").Pictures(1).Delete
                
                ActiveWorkbook.Sheets("B-CEP").Select
                ActiveSheet.Cells(55, 3).Select
                ActiveSheet.Paste
                
                wdApp.ActiveDocument.Close
                
            Next vrtSelectedItem
            ActiveWorkbook.Activate
        End If
    End With
    wdApp.Quit
    MsgBox "Done!", , "Powered By herryqyh"
End Sub

Sub CE_five()
    Dim fd As FileDialog, vrtSelectedItem As Variant
    'Dim iFile As Document
    Dim wdApp As Word.Application
    Dim iFile As Word.Document
    sht = ActiveSheet.Name

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    With fd
        .AllowMultiSelect = False
        .InitialFileName = ActiveWorkbook.Path
        .Filters.Add "Documents", "*.doc; *.docx; *.rtf", 1
        .FilterIndex = 2
        .Title = "电信端口的传导共模骚扰 CAT5"
        If .Show <> -1 Then
            MsgBox "未选择文件", vbCritical
            Exit Sub
        Else
            For Each vrtSelectedItem In .SelectedItems
                wdApp.Documents.Open vrtSelectedItem
                wdApp.Activate
                wdApp.Documents(1).Tables(1).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(30, 1).Select
                ActiveSheet.Paste
                
                wdApp.Activate
                wdApp.Documents(1).Tables(2).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(41, 1).Select
                ActiveSheet.Paste
                
                
                ActiveWorkbook.Activate
                On Error Resume Next
                ActiveWorkbook.Sheets("Y-CE56").Pictures(1).Delete
                
                wdApp.Activate
                wdApp.Documents(1).Paragraphs(3).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("Y-CE56").Select
                ActiveSheet.Cells(15, 2).Select
                ActiveSheet.Paste
                
                Set pic_range = Range("B15:F21")
                ActiveSheet.Pictures(1).Select
                With Selection.ShapeRange
                    .Top = pic_range.Top
                    .Height = pic_range.Height
                End With
                ActiveSheet.Pictures(1).Copy
                
                On Error Resume Next
                ActiveWorkbook.Sheets("B-CE56").Pictures(1).Delete
                
                ActiveWorkbook.Sheets("B-CE56").Select
                ActiveSheet.Cells(48, 4).Select
                ActiveSheet.Paste
                
                wdApp.ActiveDocument.Close
                
            Next vrtSelectedItem
             ActiveWorkbook.Activate
        End If
    End With
    wdApp.Quit
    MsgBox "Done!", , "Powered By herryqyh"
End Sub

Sub CE_six()
    Dim fd As FileDialog, vrtSelectedItem As Variant
    'Dim iFile As Document
    Dim wdApp As Word.Application
    Dim iFile As Word.Document
    sht = ActiveSheet.Name

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    With fd
        .AllowMultiSelect = False
        .InitialFileName = ActiveWorkbook.Path
        .Filters.Add "Documents", "*.doc; *.docx; *.rtf", 1
        .FilterIndex = 2
        .Title = "电信端口的传导共模骚扰 CAT6"
        If .Show <> -1 Then
            MsgBox "未选择文件", vbCritical
            Exit Sub
        Else
            For Each vrtSelectedItem In .SelectedItems
                wdApp.Documents.Open vrtSelectedItem
                
                wdApp.Activate
                wdApp.Documents(1).Tables(1).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(54, 1).Select
                ActiveSheet.Paste
                
                ActiveWorkbook.Activate
                On Error Resume Next
                ActiveWorkbook.Sheets("Y-CE56").Pictures(2).Delete

                wdApp.Activate
                wdApp.Documents(1).Tables(2).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(65, 1).Select
                ActiveWorkbook.Sheets("D-1#").Paste
                
                wdApp.Activate
                wdApp.Documents(1).Paragraphs(3).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("Y-CE56").Select
                ActiveSheet.Cells(39, 2).Select
                ActiveSheet.Paste
                
                Set pic_range = Range("B39:F45")
                ActiveSheet.Pictures(2).Select
                With Selection.ShapeRange
                    .Top = pic_range.Top
                    .Height = pic_range.Height
                End With
                ActiveSheet.Pictures(2).Copy
                
                On Error Resume Next
                ActiveWorkbook.Sheets("B-CE56").Pictures(2).Delete
                
                ActiveWorkbook.Sheets("B-CE56").Select
                ActiveSheet.Cells(74, 4).Select
                ActiveSheet.Paste
                
                wdApp.ActiveDocument.Close
                
            Next vrtSelectedItem
            ActiveWorkbook.Activate
        End If
    End With
    wdApp.Quit
    MsgBox "Done!", , "Powered By herryqyh"
End Sub

Sub LV()
    Dim fd As FileDialog, vrtSelectedItem As Variant
    'Dim iFile As Document
    Dim wdApp As Word.Application
    Dim iFile As Word.Document
    sht = ActiveSheet.Name

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    With fd
        .AllowMultiSelect = False
        .InitialFileName = ActiveWorkbook.Path
        .Filters.Add "Documents", "*.doc; *.docx; *.rtf", 1
        .FilterIndex = 2
        .Title = "30MHz～1000MHz 辐射骚扰 LV"
        If .Show <> -1 Then
            MsgBox "未选择文件", vbCritical
            Exit Sub
        Else
            For Each vrtSelectedItem In .SelectedItems
                wdApp.Documents.Open vrtSelectedItem
                wdApp.Activate
                wdApp.Documents(1).Tables(1).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(79, 1).Select
                ActiveSheet.Paste
                
                ActiveWorkbook.Activate
                On Error Resume Next
                ActiveWorkbook.Sheets("LVH").Pictures(1).Delete

                wdApp.Activate
                wdApp.Documents(1).Paragraphs(4).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("LVH").Select
                ActiveSheet.Cells(14, 2).Select
                ActiveSheet.Paste
                
                Set pic_range = Range("B14:D20")
                ActiveSheet.Pictures(1).Select
                With Selection.ShapeRange
                    .Left = pic_range.Left
                    .Top = pic_range.Top
                    .Height = pic_range.Height
                End With
                ActiveSheet.Pictures(1).Copy
                
                On Error Resume Next
                ActiveWorkbook.Sheets("B-LVH").Pictures(1).Delete
                
                ActiveWorkbook.Sheets("B-LVH").Select
                ActiveSheet.Cells(45, 5).Select
                ActiveSheet.Paste
                
                wdApp.ActiveDocument.Close
                
            Next vrtSelectedItem
            ActiveWorkbook.Activate
        End If
    End With
    wdApp.Quit
    MsgBox "Done!", , "Powered By herryqyh"
End Sub


Sub LH()
    Dim fd As FileDialog, vrtSelectedItem As Variant
    'Dim iFile As Document
    Dim wdApp As Word.Application
    Dim iFile As Word.Document
    sht = ActiveSheet.Name

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    With fd
        .AllowMultiSelect = False
        .InitialFileName = ActiveWorkbook.Path
        .Filters.Add "Documents", "*.doc; *.docx; *.rtf", 1
        .FilterIndex = 2
        .Title = "30MHz～1000MHz 辐射骚扰 LH"
        If .Show <> -1 Then
            MsgBox "未选择文件", vbCritical
            Exit Sub
        Else
            For Each vrtSelectedItem In .SelectedItems
                wdApp.Documents.Open vrtSelectedItem
                wdApp.Activate
                wdApp.Documents(1).Tables(1).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(89, 1).Select
                ActiveSheet.Paste
                
                ActiveWorkbook.Activate
                On Error Resume Next
                ActiveWorkbook.Sheets("LVH").Pictures(2).Delete
                
                wdApp.Activate
                wdApp.Documents(1).Paragraphs(4).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("LVH").Select
                ActiveSheet.Cells(22, 2).Select
                ActiveSheet.Paste
                
                Set pic_range = Range("B22:D28")
                ActiveSheet.Pictures(2).Select
                With Selection.ShapeRange
                    .Left = pic_range.Left
                    .Top = pic_range.Top
                    .Height = pic_range.Height
                End With
                ActiveSheet.Pictures(2).Copy
                
                On Error Resume Next
                ActiveWorkbook.Sheets("B-LVH").Pictures(1).Delete
                
                ActiveWorkbook.Sheets("B-LVH").Select
                ActiveSheet.Cells(52, 5).Select
                ActiveSheet.Paste
                
                wdApp.ActiveDocument.Close
                
            Next vrtSelectedItem
            ActiveWorkbook.Activate
        End If
    End With
    wdApp.Quit
    MsgBox "Done!", , "Powered By herryqyh"
End Sub

Sub HV()
    Dim fd As FileDialog, vrtSelectedItem As Variant
    'Dim iFile As Document
    Dim wdApp As Word.Application
    Dim iFile As Word.Document
    sht = ActiveSheet.Name

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    With fd
        .AllowMultiSelect = False
        .InitialFileName = ActiveWorkbook.Path
        .Filters.Add "Documents", "*.doc; *.docx; *.rtf", 1
        .FilterIndex = 2
        .Title = "1GHz以上辐射骚扰 HV"
        If .Show <> -1 Then
            MsgBox "未选择文件", vbCritical
            Exit Sub
        Else
            For Each vrtSelectedItem In .SelectedItems
                wdApp.Documents.Open vrtSelectedItem
                wdApp.Activate
                wdApp.Documents(1).Tables(1).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(100, 1).Select
                ActiveSheet.Paste
                
                wdApp.Documents.Open vrtSelectedItem
                wdApp.Activate
                wdApp.Documents(1).Tables(2).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(111, 1).Select
                ActiveSheet.Paste
                
                wdApp.Activate
                wdApp.Documents(1).Paragraphs(3).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("HVH").Select
                ActiveSheet.Cells(38, 3).Select
                ActiveSheet.Paste
                
                Set pic_range = Range("C38:G47")
                ActiveSheet.Pictures(1).Select
                With Selection.ShapeRange
                    .Left = pic_range.Left
                    .Top = pic_range.Top
                    .Height = pic_range.Height
                End With
                ActiveSheet.Pictures(1).Copy
                ActiveWorkbook.Sheets("B-HVH").Select
                ActiveSheet.Cells(65, 5).Select
                ActiveSheet.Paste
                
                wdApp.ActiveDocument.Close
                
            Next vrtSelectedItem
            ActiveWorkbook.Activate
        End If
    End With
    wdApp.Quit
    MsgBox "Done!", , "Powered By herryqyh"
End Sub

Sub HH()
    Dim fd As FileDialog, vrtSelectedItem As Variant
    'Dim iFile As Document
    Dim wdApp As Word.Application
    Dim iFile As Word.Document
    sht = ActiveSheet.Name

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    With fd
        .AllowMultiSelect = False
        .InitialFileName = ActiveWorkbook.Path
        .Filters.Add "Documents", "*.doc; *.docx; *.rtf", 1
        .FilterIndex = 2
        .Title = "1GHz以上辐射骚扰 HH"
        If .Show <> -1 Then
            MsgBox "未选择文件", vbCritical
            Exit Sub
        Else
            For Each vrtSelectedItem In .SelectedItems
                wdApp.Documents.Open vrtSelectedItem
                wdApp.Activate
                wdApp.Documents(1).Tables(1).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(124, 1).Select
                ActiveSheet.Paste
                
                wdApp.Documents.Open vrtSelectedItem
                wdApp.Activate
                wdApp.Documents(1).Tables(2).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(135, 1).Select
                ActiveSheet.Paste
                
                wdApp.Activate
                wdApp.Documents(1).Paragraphs(3).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("HVH").Select
                ActiveSheet.Cells(49, 3).Select
                ActiveSheet.Paste
                
                Set pic_range = Range("C49:G58")
                ActiveSheet.Pictures(2).Select
                With Selection.ShapeRange
                    .Left = pic_range.Left
                    .Top = pic_range.Top
                    .Height = pic_range.Height
                End With
                ActiveSheet.Pictures(2).Copy
                ActiveWorkbook.Sheets("B-HVH").Select
                ActiveSheet.Cells(73, 5).Select
                ActiveSheet.Paste
                
                wdApp.ActiveDocument.Close
                
            Next vrtSelectedItem
            ActiveWorkbook.Activate
        End If
    End With
    wdApp.Quit
    MsgBox "Done!", , "Powered By herryqyh"
End Sub

Sub Export_Origin()
    Dim new_array() As String
    ReDim new_array(8)
    Dim i As Integer
    i = 0
    sName = ActiveWorkbook.Path & "\\" & Worksheets("D-1#").Range("N10").Value & " 原始记录"
    
    OriginalSheetsArray = Array("封面", "仪器", "标准", "Y-CEP", "Y-CE56", "LVH", "HVH", "布置图")
    For Each SingleSheet In OriginalSheetsArray
        On Error Resume Next
        If ActiveWorkbook.Sheets(SingleSheet) Is Nothing Then
            '
        Else
            new_array(i) = SingleSheet
            i = i + 1
        End If
    Next SingleSheet
    ReDim Preserve new_array(i - 1)
    
    i = 0
    For Each SingleSheet In Sheets(new_array)
             i = i + SingleSheet.PageSetup.Pages.Count
    Next SingleSheet
    Sheets("封面").Cells(10, 12) = i
    
    Sheets(new_array).Select
    Sheets("封面").Activate
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=sName, Quality:=xlQualityStandard, OpenAfterPublish:=True
    Sheets("D-1#").Activate
End Sub

Sub Export_DE()
    'MsgBox ActiveWorkbook.Path
    sName = ActiveWorkbook.Path & "\\" & Worksheets("D-1#").Range("N10").Value & "-D-E"
    Sheets(Array("B1", "B2", "B-CEP", "B-CE56", "B-LVH", "B-HVH", "B-HD", "B-3")).Select
    Sheets("B1").Activate
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=sName, Quality:=xlQualityStandard, OpenAfterPublish:=True
    Sheets("D-1#").Activate
End Sub
Sub HeaderRename()
    SheetsArray = Array("B1", "B2", "B-CEP", "B-CE56", "B-LVH", "B-HVH", "B-HD", "B-3")
    For Each SingleSheet In SheetsArray
        Sheets(SingleSheet).Activate
        ActiveSheet.PageSetup.LeftHeader = "申请编号：" & Worksheets("D-1#").Range("N10").Value
        ActiveSheet.PageSetup.RightHeader = "报告编号：" & Worksheets("D-1#").Range("N11").Value & "-D-E"
    Next SingleSheet
    Sheets("D-1#").Activate
End Sub
