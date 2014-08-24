Excel-and-SQL
=============
## VBA for Excel

### 1 复制粘贴实现

```
Sub copysheets()
Dim i As Integer
Dim j As Integer
For i = 8 To 20 Step 2
    j = i / 2 - 3
    Sheets("Sheet1").Cells(10, i).Copy
    Sheets("sheet3").Activate
    Sheets("Sheet3").Cells(22, j).Select
    ActiveSheet.Paste
Next
End Sub
```

### 2 分类筛选后复制粘贴

```
Sub SelectCategory()
Dim str As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim n As Integer
Dim x As Integer
Dim y As Integer
Dim A()

Dim Class(2)
Class(0) = "A"
Class(1) = "B"
Class(2) = "C"

c = 2

n = 1

For k = 0 To 2
j = 0
m = 0
Erase A
For i = 1 To 10 Step 1
    Set c = Sheet5.Cells(i, 1).Find(Class(k))

    If Not c Is Nothing Then
        x = c.Row
        y = c.Column
        ReDim Preserve A(0 To c * (m + 1) - 1)
        A(j) = Cells(x, y).Address
        A(j + 1) = Cells(x, y + 1).Address
        j = j + c
        m = m + 1
    End If
Next
MsgBox Join(A, ",")
Sheet5.Range(Join(A, ",")).Copy
Paste Destination:=Sheets("sheet6").Cells(n, 1)
n = n + m
Next
End Sub
```

### 3  隔列进行查找匹配后复制粘贴下一列内容
```
Sub Findstr()
Dim i As Integer
Dim str As String
    For i = 3 To 6 Step 2
        str = Sheet2.Range(Cells(3, i).Address).Value
        MsgBox str
        Set c = Sheet1.Range("B1", "B10").Find(str)
        If Not c Is Nothing Then
            x = c.Row
            y = c.Column + 1
            Sheet1.Cells(x, y).Copy
            Sheet1.Paste Destination:=Sheet2.Cells(2, i)
        End If
    Next
End Sub
```

### 4 复制某个单元格图像到合并单元格
```
Sub xxx()
Sheet1.Cells(1, 1).Copy
Sheet2.Range(Cells(1, 1).Address).MergeArea.Select
Sheet2.Paste
End Sub
```

### 5 提取不同表格中的版型、面料、卖点信息
```
Sub zzzz()

Dim s As String
Dim t As String
Dim desc As String
Dim str As String
Dim tname As String
Dim i As Integer

'搜索每张表里面的关键词:“卖点”
s = "卖点"

For i = 3 To 5
    '表名
    tname = Sheets(i).Name
    '在表里面查找关键词
    Set md = Sheets(tname).Cells.Find(s)
    '保存关键词的位置
    x = md.Row
    y = md.Column
    
    '在目标表格查找款号,同上述表名
    Set xxd = Sheets("Sheet5").Cells.Find(tname)
    If Not xxd Is Nothing Then
        '保存款号的位置
        nx = xxd.Row
        ny = xxd.Column
    
        '保存卖点的描述文字
        desc = Sheets(tname).Cells(x, y + 1).Value
        
        '在描述中查找版型
        '“版”的位置
        id1 = InStr(desc, "版型")
        If id1 >= 1 Then
            '查找换行符
            id11 = InStr(id1 + 1, desc, Chr(10))
            '跳过前导n个字符(版型和空格字符=3),提取id11-id1+1-n-1个字符
            str = Mid(desc, id1 + 3, id11 - id1 - 3)
            '保存到目标表格的对应位置
            Sheet5.Cells(nx, ny + 1) = str
        End If
        
        
        id2 = InStr(desc, "面料")
        If id2 >= 1 Then
            id22 = InStr(id2 + 1, desc, Chr(10))
            str = Mid(desc, id2 + 3, id22 - id2 - 3)
            Sheet5.Cells(nx, ny + 2) = str
        End If
        
        id3 = InStr(desc, "设计")
        If id3 >= 1 Then
            id33 = InStr(id3 + 1, desc, Chr(10))
            str = Mid(desc, id3 + 3, id33 - id3 - 3)
            Sheet5.Cells(nx, ny + 3) = str
        End If
        
        
        id4 = InStr(desc, "工艺")
        If id4 >= 1 Then
            '保留最后len-id4+1-n
            str = Right(desc, Len(desc) - id4 - 2)
            Sheet5.Cells(nx, ny + 4) = str
        End If
    End If
Next

End Sub
```

### 6 分类后的品类提取并放在对应的单元格
```
Sub kuanhaofenlei()
Dim i As Integer
Dim n As Integer
Dim j As Integer
Dim str As String
j = 2
'n定义每一类的款号数目
n = 10
For i = 2 To (2 + 4 * (n - 1)) Step 4   '定义需要在哪些单元格进行操作
    str = Sheets(6).Cells(j, 1).Value  '取单元格的值
    Sheets(7).Cells(3, i) = str        '将对应款号写入到对应商品分类的行中
    j = j + 1
Next
End Sub
```
 
### 7 批量清除合并单元格内容

```
'清除上一周新品分析中商品分类中每个类目的内容
Sub clearcon()
Dim i As Integer
Dim n As Integer

'n定义每一类的款号数目
n = 10
For i = 2 To (2 + 4 * (n - 1)) Step 4
Sheet7.Range(Cells(3, i).Address).MergeArea.ClearContents
Next
End Sub

```

### 8 批量清除正常单元格中内容
```
'清除正常单元格的内容
Sub clearcon1()
Dim i As Integer
Dim n As Integer

'n定义每一类的款号数目
n = 10
For i = 2 To (2 + 4 * (n - 1)) Step 4
Sheet7.Cells(10, i).ClearContents
Next
End Sub
```

### 9 批量匹配图片
```
Sub matchpic()
Dim n As Integer
Dim kh As String
n = 10
For i = 2 To 2 + 4 * (n - 1) Step 4
    kh = Sheet7.Cells(3, i).Value
    MsgBox kh
    '款号必须具有唯一性
    Set C = Sheet6.Range("A2:A11").Find(kh)
    x = C.Row
    y = C.Column + 6
    'MsgBox x & "-" & y
    Sheet6.Cells(x, y).Copy
    '图像大小必须小于单元格
    Sheet7.Range(Cells(2, i - 1).Address).MergeArea.Select
    Sheet7.Paste
Next
End Sub
```

### 如何选择表，区域，以及单元格
```
'see: http://support.microsoft.com/kb/291308

'如何在活动工作表上选择单元格
ActiveSheet.Cells(5, 4).Select
ActiveSheet.Range("D5").Select

'如何在同一工作簿中的另一工作表上选择单元格
Application.Goto ActiveWorkbook.Sheets("Sheet2").Cells(6, 5)
Application.Goto (ActiveWorkbook.Sheets("Sheet2").Range("E6"))
'or
Sheets("Sheet2").Activate
ActiveSheet.Cells(6, 5).Select

'如何在另外一个工作簿中的工作表上选择单元格
Application.Goto Workbooks("BOOK2.XLS").Sheets("Sheet1").Cells(7, 6)
Application.Goto Workbooks("BOOK2.XLS").Sheets("Sheet1").Range("F7")
'也可以激活工作表，然后使用上面的方法 1 来选择单元格
Workbooks("BOOK2.XLS").Sheets("Sheet1").Activate
ActiveSheet.Cells(7, 6).Select

'如何在活动工作表上选择单元格区域
ActiveSheet.Range(Cells(2, 3), Cells(10, 4)).Select
ActiveSheet.Range("C2:D10").Select
ActiveSheet.Range("C2", "D10").Select

'如何在同一工作簿中另一工作表上选择单元格区域
Application.Goto ActiveWorkbook.Sheets("Sheet3").Range("D3:E11")
Application.Goto ActiveWorkbook.Sheets("Sheet3").Range("D3", "E11")
'也可以激活工作表，然后使用上面的方法来选择范围
Sheets("Sheet3").Activate
ActiveSheet.Range(Cells(3, 4), Cells(11, 5)).Select

'如何在另外一个工作簿中的工作表上选择单元格区域
Application.Goto Workbooks("BOOK2.XLS").Sheets("Sheet1").Range("E4:F12")
Application.Goto Workbooks("BOOK2.XLS").Sheets("Sheet1").Range("E4", "F12")
'也可以激活工作表，然后使用上面的方法来选择范围
Workbooks("BOOK2.XLS").Sheets("Sheet1").Activate
ActiveSheet.Range(Cells(4, 5), Cells(12, 6)).Select

'如何在活动工作表上选择命名区域
Range("Test").Select				
Application.Goto "Test"

'如何在另一工作簿中的工作表上选择命名区域
Application.Goto Workbooks("BOOK2.XLS").Sheets("Sheet2").Range("Test")
'也可以激活工作表，然后使用上面的方法
Workbooks("BOOK2.XLS").Sheets("Sheet2").Activate
Range("Test").Select

'如何相对于活动单元格选择单元格
ActiveCell.Offset(5, -4).Select
ActiveCell.Offset(-2, 3).Select


'如何相对于另一（非活动）单元格选择单元格
ActiveSheet.Cells(7, 3).Offset(5, 4).Select
ActiveSheet.Range("C7").Offset(5, 4).Select

'如何相对于指定区域选择单元格区域偏移
ActiveSheet.Range("Test").Offset(4, 3).Select
'命名区域在另一个（非活动）工作表上，请首先激活该工作表，接着使用以下示例选择该区域
Sheets("Sheet3").Activate
ActiveSheet.Range("Test").Offset(4, 3).Select

'如何选择指定的区域并调整所选区域的大小
Range("Database").Select
Selection.Resize(Selection.Rows.Count + 5, Selection.Columns.Count).Select

'如何选择指定的区域、使之偏移然后调整其大小
Range("Database").Select
Selection.Offset(4, 3).Resize(Selection.Rows.Count + 2, Selection.Columns.Count + 1).Select

'如何选择两个或更多指定区域的联合
Application.Union(Range("Test"), Range("Sample")).Select
'两个区域必须位于同一个工作表上才行。 另请注意，Union 方式不可以跨工作表使用
Set y = Application.Union(Range("Sheet1!A1:B2"), Range("Sheet1!C3:D4"))

'如何选择两个或更多指定区域的交集
Application.Intersect(Range("Test"), Range("Sample")).Select

'如何选择一列连续数据的最后一个单元格
'选择一个连续列中的最后一个单元格
ActiveSheet.Range("a1").End(xlDown).Select

'如何选择一列连续数据底部的空白单元格
ActiveSheet.Range("a1").End(xlDown).Offset(1,0).Select

'如何在一列中选择整个相邻单元格区域
ActiveSheet.Range("a1", ActiveSheet.Range("a1").End(xlDown)).Select
' or
ActiveSheet.Range("a1:" & ActiveSheet.Range("a1").End(xlDown).Address).Select


'如何在一列中选择整个非连续单元格区域
ActiveSheet.Range("a1",ActiveSheet.Range("a65536").End(xlUp)).Select
ActiveSheet.Range("a1:" & ActiveSheet.Range("a65536").End(xlUp).Address).Select

'如何选择矩形单元格区域
ActiveSheet.Range("a1").CurrentRegion.Select

```

### Worksheet使用方法
```
'see: http://blog.csdn.net/znyang/article/details/14120817
```

### Cells字体设置
```
'see: http://www.feiesoft.com/vba/excel/xlproCells.htm

'将 Sheet1 中单元格 C5 的字体大小设置为 14 磅
Worksheets("Sheet1").Cells(5, 3).Font.Size = 14

'本示例清除 Sheet1 上第一个单元格的公式
Worksheets("Sheet1").Cells(1).ClearContents

'本示例将 Sheet1 上所有单元格的字体设置为 8 磅的“Arial”字体
With Worksheets("Sheet1").Cells.Font
    .Name = "Arial"
    .Size = 8
End With

'在 Sheet1 上的单元格区域 A1:J4 中循环，将其中小于 0.001 的值替换为 0（零）。
For rwIndex = 1 to 4
    For colIndex = 1 to 10
        With Worksheets("Sheet1").Cells(rwIndex, colIndex)
            If .Value < .001 Then .Value = 0
        End With
    Next colIndex
Next rwIndex

'将 Sheet1 上单元格区域 A1: C5 的字体样式设置为斜体。
Worksheets("Sheet1").Activate
Range(Cells(1, 1), Cells(5, 3)).Font.Italic = True

'搜索列“myRange”中的数据
'如果发现某单元格的值与上面的一个单元格的值相等，则将显示这个包含重复数据的单元格的地址
Set r = Range("myRange")
For n = 1 To r.Rows.Count
    If r.Cells(n, 1) = r.Cells(n + 1, 1) Then
        MsgBox "Duplicate data in " & r.Cells(n + 1, 1).Address
    End If
Next n

'用union方法联合选取不连续的若干非空的行：
sub aa()
	dim a as range
	for i = 1 to 2000
		if cells(i,13)  <> "" then
			if a is nothing then
				set a=rows(i)
			else
				set a =union(a,rows(i))
			end if
		end if
	next i
	if not a is nothing then
		a.select
	end if
end sub

```
