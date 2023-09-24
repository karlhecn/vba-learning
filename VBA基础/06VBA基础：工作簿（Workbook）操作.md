## VBA基础：工作簿（Workbook）操作



**1. 打开工作簿：**

```vba
Sub OpenWorkbook()
    ' 打开指定的工作簿
    Workbooks.Open "C:\Path\To\Your\File.xlsx"
End Sub
```

这个示例代码演示如何打开指定路径的工作簿。

**2. 创建新工作簿：**

```vba
Sub CreateNewWorkbook()
    ' 创建一个新的工作簿
    Workbooks.Add
End Sub
```

这个示例代码创建一个新的工作簿。

**3. 保存工作簿：**

```vba
Sub SaveWorkbook()
    ' 保存当前工作簿
    ThisWorkbook.Save
End Sub
```

这个示例代码保存当前工作簿。

**4. 关闭工作簿：**

```vba
Sub CloseWorkbook()
    ' 关闭当前工作簿
    ThisWorkbook.Close
    
    ' ThisWorkbook.Close SaveChanges:=False ' 如果不保存更改
End Sub
```

这个示例代码关闭当前工作簿。

**5. 复制工作簿：**

```vba
Sub CopyWorkbook()
    ' 复制当前工作簿
    ThisWorkbook.Copy
End Sub
```

这个示例代码复制当前工作簿。

**6. 激活工作簿：**

```vba
Sub ActivateWorkbook()
    ' 激活指定名称的工作簿
    Workbooks("YourWorkbook.xlsx").Activate
End Sub
```

这个示例代码激活指定名称的工作簿。

**7. 循环遍历所有打开的工作簿：**

```vba
Sub LoopThroughWorkbooks()
    Dim wb As Workbook
    For Each wb In Workbooks
        ' 在这里对每个打开的工作簿执行操作
        MsgBox "工作簿名称：" & wb.Name
    Next wb
End Sub
```

这个示例代码循环遍历所有打开的工作簿，并对每个工作簿执行操作。

**8. 删除工作簿：**

```vba
Sub DeleteWorkbook()
    ' 删除指定名称的工作簿
    Workbooks("YourWorkbook.xlsx").Close SaveChanges:=False
    Kill "C:\Path\To\Your\File.xlsx" ' 删除文件
End Sub
```

这个示例代码关闭并删除指定名称的工作簿。

**9. 合并工作簿：**

```
vbaCopy codeSub MergeWorkbooks()
    ' 打开第一个工作簿
    Workbooks.Open "C:\Path\Workbook1.xlsx"
    
    ' 打开第二个工作簿
    Workbooks.Open "C:\Path\Workbook2.xlsx"
    
    ' 合并工作簿
    Workbooks("Workbook1.xlsx").Sheets(1).Copy Before:=Workbooks("Workbook2.xlsx").Sheets(1)
    
    ' 保存并关闭第二个工作簿
    Workbooks("Workbook2.xlsx").Save
    Workbooks("Workbook2.xlsx").Close
End Sub
```

这些示例代码演示了Excel VBA中常见的工作簿操作，包括打开、创建、保存、关闭、复制、激活、循环遍历和删除工作簿等。您可以根据需要自定义这些示例代码，并将它们应用于您的Excel VBA项目中。