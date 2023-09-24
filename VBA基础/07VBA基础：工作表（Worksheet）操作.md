## VBA基础：工作表（Worksheet）操作



**1. 创建新工作表：**

```vba
Sub CreateNewWorksheet()
    ' 在活动工作簿中创建一个新工作表
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "新工作表"
End Sub
```

这个示例代码在活动工作簿中创建一个新工作表，并将其命名为 "新工作表"。

**2. 选择工作表：**

```vba
Sub SelectWorksheet()
    ' 选择指定名称的工作表
    ThisWorkbook.Sheets("工作表1").Select
End Sub
```

这个示例代码选择指定名称的工作表。

**3. 复制工作表：**

```vba
Sub CopyWorksheet()
    ' 复制指定名称的工作表并将其插入到末尾
    ThisWorkbook.Sheets("工作表1").Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
End Sub
```

这个示例代码复制指定名称的工作表并将其插入到末尾。

**4. 删除工作表：**

```vba
Sub DeleteWorksheet()
    ' 删除指定名称的工作表
    Application.DisplayAlerts = False ' 禁用警告
    ThisWorkbook.Sheets("工作表1").Delete
    Application.DisplayAlerts = True ' 启用警告
End Sub
```

这个示例代码删除指定名称的工作表。注意，我们在删除工作表之前禁用了警告。

**5. 重命名工作表：**

```vba
Sub RenameWorksheet()
    ' 重命名指定名称的工作表
    ThisWorkbook.Sheets("工作表1").Name = "新名称"
End Sub
```

这个示例代码重命名指定名称的工作表。

**6. 隐藏和显示工作表：**

```vba
Sub HideAndUnhideWorksheet()
    ' 隐藏指定名称的工作表
    ThisWorkbook.Sheets("工作表1").Visible = xlSheetHidden
    
    ' 显示指定名称的工作表
    ThisWorkbook.Sheets("工作表1").Visible = xlSheetVisible
End Sub
```

这个示例代码演示如何隐藏和显示指定名称的工作表。

**7. 循环遍历工作表：**

```vba
Sub LoopThroughWorksheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        ' 在这里对每个工作表执行操作
        MsgBox "工作表名称：" & ws.Name
    Next ws
End Sub
```

这个示例代码循环遍历工作簿中的所有工作表，并对每个工作表执行操作。

**8. 移动工作表：**

```
Sub MoveWorksheet()
    ' 将当前工作表移到工作簿的末尾
    Sheets("Sheet1").Move After:=Sheets(Sheets.Count)
End Sub
```

**9. 创建新工作簿并复制工作表：**

```
Sub CreateAndCopyWorksheet()
    ' 创建一个新的工作簿
    Dim NewWb As Workbook
    Set NewWb = Workbooks.Add
    
    ' 复制当前工作表到新工作簿
    ThisWorkbook.Sheets("Sheet1").Copy Before:=NewWb.Sheets(1)
End Sub
```

这些示例代码演示了Excel VBA中常见的工作表操作，包括创建、选择、复制、删除、重命名、隐藏和显示工作表，以及循环遍历工作表。您可以根据需要自定义这些示例代码，并将它们应用于您的Excel VBA项目中。