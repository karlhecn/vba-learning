## VBA基础：使用内置函数和扩充自定义函数

**1. 使用 SUM 函数：**

```vba
Sub SumFunctionExample()
    ' 使用 SUM 函数计算单元格 A1 到 A10 的总和
    Dim total As Double
    total = Application.WorksheetFunction.Sum(Range("A1:A10"))
    MsgBox "总和是：" & total
End Sub
```

在此示例中，我们使用了 SUM 函数来计算单元格 A1 到 A10 的总和。

**2. 使用 VLOOKUP 函数：**

```vba
Sub VLookupFunctionExample()
    ' 使用 VLOOKUP 函数查找数据
    Dim valueToFind As Variant
    Dim result As Variant
    valueToFind = "John"
    result = Application.WorksheetFunction.VLookup(valueToFind, Range("A1:B10"), 2, False)
    MsgBox "找到的结果是：" & result
End Sub
```

在此示例中，我们使用 VLOOKUP 函数来查找数据。

**3. 使用 AVERAGE 函数：**

```vba
Sub AverageFunctionExample()
    ' 使用 AVERAGE 函数计算平均值
    Dim average As Double
    average = Application.WorksheetFunction.Average(Range("A1:A10"))
    MsgBox "平均值是：" & average
End Sub
```

在此示例中，我们使用了 AVERAGE 函数来计算单元格 A1 到 A10 的平均值。

这些示例代码演示了Excel VBA中常见的使用内置函数的操作，包括 SUM、VLOOKUP和 AVERAGE 函数。您可以根据需要自定义这些示例代码，并将它们应用于您的Excel VBA项目中。内置函数可以帮助您执行各种计算和操作，提高Excel的功能。



### 扩充自定义函数

要创建一个可以在Excel单元格中直接调用的自定义公式，您需要创建一个VBA函数，并将其注册为Excel自定义函数。以下是一个示例代码，演示如何创建和使用这样的自定义公式：

```vba
Function MyCustomFormula(value1 As Double, value2 As Double) As Double
    ' 自定义公式，计算两个值的和
    MyCustomFormula = value1 + value2
End Function
```

在这个示例中，我们创建了一个名为 `MyCustomFormula` 的自定义函数，它接受两个参数 `value1` 和 `value2`，并返回它们的和。

接下来，您需要将这个自定义函数注册为Excel自定义函数。请按照以下步骤操作：

1. 打开Excel，然后按下 `Alt` + `F11` 打开VBA编辑器。

2. 在VBA编辑器中，选择插入 > 模块，以创建一个新的VBA模块。

3. 将上面的自定义函数代码粘贴到新模块中。

4. 在Excel工作表中，您可以在单元格中使用以下方式调用这个自定义公式：`=MyCustomFormula(A1, B1)`，其中 `A1` 和 `B1` 是您要相加的值。

5. 在单元格中输入该公式后，按下 `Enter` 键即可得到计算结果。

这样，您就创建了一个可以在Excel单元格中直接调用的自定义公式。当您在单元格中使用这个公式时，Excel会自动调用VBA函数并返回结果。这对于执行自定义计算或操作非常有用，可以增强Excel的功能。