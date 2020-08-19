# VBA Macro tools & Excel templates for CCC reports

## Feature

- [ ] 导入EMC32等软件测试的出的EUT表格、图片至Excel
  - [x] 传导骚扰 电源端（电压）
  - [x] 传导骚扰 电信端口（CAT 5&6）
  - [x] 低频辐射
  - [x] 高频辐射
  - [ ] 谐波电流
- [ ] excel中原始表格数据同步至原始报告以及数据报告
  - [x] 传导骚扰 电源端（电压）
  - [x] 传导骚扰 电信端口（CAT 5&6）
  - [x] 低频辐射
  - [x] 高频辐射
  - [ ] 谐波电流
- [ ] 试验照片导入
- [x] 数据报告（-D-E）文档页眉重命名
- [x] 导出PDF格式报告并按照统一命名格式
- [x] 锁定excel可编辑区域
- [ ] 多样品数量模板支持
- [ ] 导入数据文件类别错误检查



## 实现

### 导入EMC32等软件测试的出的EUT表格、图片至Excel

以导入传导骚扰 电源端（电压）为例

导入程序接受的格式为`doc docx rtf`

```vbscript
Sub CEP()
    Dim fd As FileDialog, vrtSelectedItem As Variant
    Dim wdApp As Word.Application
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
            'actually only one file was selected
            For Each vrtSelectedItem In .SelectedItems
                wdApp.Documents.Open vrtSelectedItem
                
                'copy table 1
                wdApp.Activate
                wdApp.Documents(1).Tables(1).Range.Copy
                ActiveWorkbook.Activate
                'desination sheet 
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(5, 1).Select
                ActiveSheet.Paste
                
                'copy table 2        
                wdApp.Activate
                wdApp.Documents(1).Tables(2).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("D-1#").Select
                ActiveSheet.Cells(16, 1).Select
                ActiveSheet.Paste
                
                'delete existed pic from original
                ActiveWorkbook.Activate
                On Error Resume Next
                ActiveWorkbook.Sheets("Y-CEP").Pictures(1).Delete
                
                'copy pic 1 to orignal report               
                wdApp.Activate
                wdApp.Documents(1).Paragraphs(3).Range.Copy
                ActiveWorkbook.Activate
                ActiveWorkbook.Sheets("Y-CEP").Select
                ActiveSheet.Cells(21, 2).Select
                ActiveSheet.Paste
                                        
                'set position and resize
                Set pic_range = Range("B21:F27")
                ActiveSheet.Pictures(1).Select
                With Selection.ShapeRange
                    .Top = pic_range.Top
                    .Height = pic_range.Height
                End With
                
                'copy to -D-E
                ActiveSheet.Pictures(1).Copy
                                            
                'delete existed pic from -D-E
                On Error Resume Next
                ActiveWorkbook.Sheets("B-CEP").Pictures(1).Delete
                
                'paste pic                           
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
```

tips: 

- VBA函数命名不可采用字母+数字形式，会被认作为某一个单元格的宏 

​		etc. `Sub CAT5 ()`-> `Sub CAT_five ()`

- 新建宏时选择 `录制宏->保存在个人工作簿` 创建个人工作簿，保存位置在

  ```
  C:\Users\{你的Windows用户名}\AppData\Roaming\Microsoft\Excel\XLSTART
  ```

  以便在所有工作簿中调用该宏

  

 ### excel中原始表格数据同步至原始报告以及数据报告

使用excel函数 `=` 实现

i.e. `='D-1#'!N9`



###  试验照片导入

### 数据报告（-D-E）文档页眉重命名

### 导出PDF格式报告并按照统一命名格式

### 锁定excel可编辑区域

### 多样品数量模板支持

### 导入数据文件类别错误检查

## 导出和使用

