## VBA基础：单元格（Range/Cells）操作



**1. 选择单元格：**

```vba
Sub SelectCell()
    ' 选择单元格 A1
    Range("A1").Select
End Sub
```

这个示例代码选择了单元格 A1。

```
Sub SelectCellRange()
    ' 选择工作表中的单元格范围（例如，A1 到 B5）
    Range("A1:B5").Select
End Sub
```

这个示例代码选择工作表中的单元格范围（A1 到 B5）。

```
Sub SelectCellRangeOtherSheet()
    ' 选择工作表Sheet1中的单元格范围（A1）
    Sheet2.Activate
    Sheet2.Range("A2").Select
    
'    Worksheets("工资表").Select
'    Worksheets("工资表").Range("A2").Select
    
End Sub
```

这个示例代码选择工作表“ Sheet2”(Sheet2的名称为“工资表”)中的单元格范围（A2）,注意，如果跨sheet选取，需要先设指定的sheet为活跃状态。

**2. 设置单元格的值：**

```vba
Sub SetCellValue()
    ' 设置单元格 A1 的值为 "Hello, World!"
    Range("A1").Value = "Hello, World!"
End Sub
```

这个示例代码将单元格 A1 的值设置为 "Hello, World!"。

**3. 读取单元格的值：**

```vba
Sub ReadCellValue()
    ' 读取单元格 A2 的值并显示在消息框中
    Dim cellValue As String
    cellValue = Range("A2").Value
    MsgBox "单元格 A2 的值是：" & cellValue
    
    'Cells方式获取值
    cellValue = Cells(2, 1).Value
    MsgBox "单元格 2行 1列 的值是：" & cellValue
    
    'Sheet2的range方式获取值
    cellValue = Sheet2.range("A2").Value
    MsgBox "Sheet2 单元格 A2 的值是：" & cellValue
    
    'Sheet名为“工资表”的Cells方式获取值
    cellValue = Worksheets("工资表").range("A2").Value
    MsgBox "工资表 单元格 A2 的值是：" & cellValue
    
    'Sheet2的Cells方式获取值
    cellValue = Sheet2.Cells(2, 1).Value
    MsgBox "Sheet2 单元格 2行 1列 的值是：" & cellValue
    
    'Sheet名为“工资表”的Cells方式获取值
    cellValue = Worksheets("工资表").Cells(2, 1).Value
    MsgBox "工资表 单元格 2行 1列 的值是：" & cellValue
End Sub
```

这个示例代码读取单元格 A2 及 Sheet2的单元格A2 的值并显示在消息框中。

**4. 复制和粘贴单元格：**

```vba
Sub CopyPasteCell()
    ' 复制单元格 A1 的值到 B1
    Range("A1:D7").Copy Destination:=Range("H1")
End Sub
```

这个示例代码复制单元格 A1 的值并粘贴到单元格 B1。

**5. 设置单元格格式：**

```vba
Sub SetCellFormat()
    ' 设置单元格 A1 的字体颜色为红色
    Range("A1").Font.Color = RGB(255, 0, 0) ' 红色
End Sub
```

这个示例代码设置单元格 A1 的字体颜色为红色。

**6. 合并和拆分单元格：**

```vba
Sub MergeAndUnmergeCells()
    ' 合并单元格 A1 到 B1
    Range("A1:B1").Merge
    
    ' 拆分合并的单元格
    Range("A1:B1").UnMerge
End Sub
```

这个示例代码合并单元格 A1 到 B1，然后拆分这些合并的单元格。

**7. 设置单元格公式：**

```vba
Sub SetCellFormula()
    ' 设置单元格 E2 的公式为求和 B2 到 C2 的值
    Range("E2").Formula = "=SUM(B2:C2)"
End Sub
```

这个示例代码设置单元格 E2 的公式为求和 B2 到 C2 的值。

**8. 遍历处理单元格：**
```vba
Sub ProcessSelectedRange()
    Dim selectedRange As Range
    Dim cell As Range

    ' 获取选定的范围
    Set selectedRange = Range("B2:D7")
    
    ' 遍历选定范围的每一个单元格
    For Each cell In selectedRange
        ' 检查单元格所在的行和列
        If cell.Row Mod 2 = 0 And cell.Column Mod 2 = 0 Then
            ' 偶数行、偶数列单元格设为红色
            cell.Font.Color = RGB(255, 0, 0) ' 红色
        End If
    Next cell
End Sub
```
**9. 插入和删除单元格：**

```vba
Sub InsertAndDeleteCells()
    ' 在单元格 B1 前插入一列
    Range("B1").EntireColumn.Insert
    
    ' 删除单元格 B1
    Range("B1").Delete
End Sub
```

这个示例代码在单元格 B1 后插入一列，并删除单元格 B1。

**10. 清除单元格：**

```vba
Sub DeleteCells()
    ' 清除单元格 E2（清除内容和格式）
    Range("E2").Clear
    
    ' 清除单元格 E2（只清除内容，保留格式）
    Range("E2").ClearContents
End Sub
```

**11. 未知空白行行号动态循环单元格：**

1）通过预设大行数，再判断新行的开头单元格是否有值

```vba
Sub UnKnownRows1()

    For Each cell In Range("A2:B20")
        If cell.Column = 1 And cell.Value = "" Then
            Exit For
        End If
        
        MsgBox cell.Row & " " & cell.Column & " " & cell.Value
    Next

End Sub
```

2）通过Do While循环+自增行号，判断新行开头单元格是否有值

```vba
Sub UnKnownRows2()
    
    Dim intRow As Integer
    intRow = 2
    
    Do While Range("A" & intRow) <> ""
       
        MsgBox intRow & " " & CStr(1) & " " & Range("A" & intRow).Value
        MsgBox intRow & " " & CStr(2) & " " & Range("B" & intRow).Value
        
        intRow = intRow + 1
    Loop

End Sub
```

这些示例代码演示了Excel VBA中常见的单元格操作，包括选择、设置值、读取值、复制粘贴、设置格式、合并拆分、设置公式以及插入和删除单元格。您可以根据需要自定义这些示例代码，并将它们应用于您的Excel VBA项目中。
