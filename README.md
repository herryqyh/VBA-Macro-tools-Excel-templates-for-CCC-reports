# VBA Macro tools & Excel templates for CCC reports

//todo 

推迟*尝试解决某些情况下__表格图片复制后无法黏贴的问题__，可以参考[这里](https://stackoverflow.com/questions/10714251/how-to-avoid-using-select-in-excel-vba)重写*

​	*https://zhidao.baidu.com/question/257974352.html 自动添加加载项*



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

- [ ] 试验照片导入（试验布置图）

- [x] 数据报告（-D-E）文档页眉重命名

- [x] 导出PDF格式报告并按照统一格式命名

- [x] 清除试验结果表格以及图片数据（实拍图片除外）

- [x] 锁定excel可编辑区域

- [x] 多样品数量模板支持

- [ ] 导入数据文件类别错误检查

    

## 在撰写此文档时的程序版本

- Office 2016





## 导出和使用

- 需要excel模板与VBA宏配合使用，VBA宏通过直接导出PERSONAL.XLSB后，放到其他用户的

  ```powershell
  C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Excel\XLSTART
  ```

  文件夹中即可使用

- 本文所用office版本为2016版本，不同版本excel间可能有不兼容情况导致PERSONAL.XLSB无法在启动时自动加载，可以选择手动打开PERSONAL.XLSB工作簿后再打开要编辑的文档

  - 对于dll库缺失/dll库版本不一致，需要在VBA的*工具->引用* 内重新选择正确的版本

   - 对于无法创建/修改在XLSTART内的文件，需要在其他位置编辑完成后手动移入该文件夹内

     或着，在360中关闭*office宏病毒免疫*  //360 R U kidding??? 因噎废食

- 使用时导入word文档前，确保被导入文档是关闭状态

- 各组数据需要按照报告呈现顺序依次导入，以确保导入图片顺序呈现正确

- 调用方法

  - 对于所有版本excel，可以在*视图 -> 宏*  中选择对应的宏执行
  - 对于老版本excel，可以将宏添加到自定义快速访问工具栏中
  - 对于可自定义选项卡的新版本，可以在*文件 -> 选项 -> 自定义功能区* 中将宏添加到选项卡中==（单个样品时推荐方法）==
  - 以上两种方案操作，[具体操作参考这里](https://jingyan.baidu.com/article/4dc40848753509c8d946f1a7.html)
  - 在开发工具，插入按钮控件并且绑定带有`_Click`后缀的宏，[具体操作参考这里](https://jingyan.baidu.com/article/3a2f7c2e32340e26afd61180.html)，（开发工具选项卡可以在excel选项中开启）==（多个样品时推荐方法）==

- 宏功能说明

  - ```CEP CE_five CE_six LV LH HV HH```

    分别导入电源端传导骚扰、电信端口传导骚扰（CAT5、6）、低频辐射（垂直、水平方向）、高频辐射（垂直、水平方向）

  - ```HeaderRename```

    按照工作表```D-1#``` 中```N9 N10``` 单元格内容重命名-D-E报告的页眉

  - ```Export_Origin Export_DE```
    导出原始报告以及数据报告

  - ```Data_CLR Pic_CLR```

    清除试验数据和曲线示意图，```布置图```工作表内的除外==（导入多个样品数据时建议先执行此命令）==

  - 带有`_Click`后缀的模块

    被多样品数据模板内的按钮调用，点击后在当前sheet黏贴表格，并且在原始记录与数据报告sheet中，通过鼠标选取相应图片位置。

    


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

​		i.e. `Sub CAT5 ()`-> `Sub CAT_five ()`

- 新建宏时选择 `录制宏->保存在个人工作簿` 创建个人工作簿PERSONAL.XLSB，其保存位置在

  ```powershell
  C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Excel\XLSTART
  ```

  保存在此处便于在所有工作簿中调用该宏

- 导出宏时可以选择直接导出PERSONAL.XLSB，放到其他用户对应位置

- 不同版本excel间可能有不兼容情况导致PERSONAL.XLSB无法在启动时自动加载，可能的解决方法包括

  - 可以选择手动打开PERSONAL.XLSB工作簿
  - 重新封装代码，使用对应版本的dll的引用

  

### 多个样品的模板支持

使用经过修改的带有```_Click```后缀的代码实现，绑定到按钮控件

通过```ActiveSheet``` 指定激活的工作表为黏贴位置

通过 单元格内值 D-n# 指定图片黏贴位置

 //todo ```InputBox``` 点击取消时的错误处理



 ### excel中原始表格数据同步至原始报告以及数据报告

使用excel函数 `=` 实现

i.e. `='D-1#'!N9`

使用该功能的部分包括试验数据表格、原始记录报告编号、仪器设备清单、试验标准、试验样机配置说明、温度湿度气压



###  试验照片导入（试验布置图）

推迟实现



### 数据报告（-D-E）文档页眉重命名

```vbscript
Sub HeaderRename()
    SheetsArray = Array("B1", "B2", "B-CEP", "B-CE56", "B-LVH", "B-HVH", "B-HD", "B-3")
    For Each SingleSheet In SheetsArray
        Sheets(SingleSheet).Activate
        ActiveSheet.PageSetup.LeftHeader = "申请编号：" & Worksheets("D-1#").Range("N10").Value
        ActiveSheet.PageSetup.RightHeader = "报告编号：" & Worksheets("D-1#").Range("N11").Value & "-D-E"
    Next SingleSheet
    Sheets("D-1#").Activate
End Sub
```



### 导出PDF格式报告并按照统一命名格式、清除试验数据及示意图

以导出原始记录为例

```vbscript
Sub Export_Origin()
    Dim new_array() As String
    ReDim new_array(8)
    Dim i As Integer
    i = 0
    'export to the current folder as a pdf file
    sName = ActiveWorkbook.Path & "\\" & Worksheets("D-1#").Range("N10").Value & " 原始记录"
    'array of all sheets
    OriginalSheetsArray = Array("封面", "仪器", "标准", "Y-CEP", "Y-CE56", "LVH", "HVH", "布置图")
    
    'check if existed
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
    
    'calculate total pages
    i = 0
    For Each SingleSheet In Sheets(new_array)
             i = i + SingleSheet.PageSetup.Pages.Count
    Next SingleSheet
    Sheets("封面").Cells(10, 12) = i
    
    'select all existed sheets
    Sheets(new_array).Select
    Sheets("封面").Activate
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=sName, Quality:=xlQualityStandard, OpenAfterPublish:=True
    Sheets("D-1#").Activate
End Sub
```



### 锁定excel可编辑区域

**选取允许编辑的单元格**
选中单元格并右键，在弹出的菜单中点击*设置单元格格式*，在弹出的*自定义序列* 窗口中点击*保护* 选项卡，然后将*锁定*复选框取消选中并点击*确定*   

**保护其他单元格不能编辑**
选择*审阅* 选项卡中*保护工作表*，在弹出窗口中点击*确定*
tips: 不能保护需要导入图片的工作表，会导致图片无法插入（仅部分工作表会有bug？？？ 未解决）



### 导入数据文件类别错误检查

 推迟实现





