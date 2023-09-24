## VBA基础：函数和子过程

**1. 函数的示例：**

```vba
Function AddNumbers(num1 As Double, num2 As Double) As Double
    ' 声明函数并定义参数及返回值的数据类型
    ' 计算两个数字的和并返回结果
    AddNumbers = num1 + num2
End Function

Sub CallFunctionExample()
    ' 调用函数并获取返回值
    Dim result As Double
    result = AddNumbers(5, 3)
    MsgBox "结果是：" & result
End Sub
```

在这个示例中，我们定义了一个名为 `AddNumbers` 的函数，它接受两个参数 `num1` 和 `num2`，并返回它们的和。我们使用 `CallFunctionExample` 子过程来调用这个函数并获取返回值。

**2. 子过程的示例：**

```vba
Sub PrintMessage(message As String)
    ' 声明子过程并定义参数
    ' 打印传入的消息
    MsgBox message
End Sub

Sub CallSubroutineExample()
    ' 调用子过程并传递参数
    Call PrintMessage("Hello, World!")
End Sub
```

在这个示例中，我们定义了一个名为 `PrintMessage` 的子过程，它接受一个参数 `message` 并在消息框中显示传入的消息。我们使用 `CallSubroutineExample` 子过程来调用这个子过程并传递参数。

**3. 作用域的示例：**

在 Excel VBA 中，函数和子过程可以具有不同的作用域，包括模块级作用域和过程级作用域。以下是示例代码：

```vba
' 模块级变量
Dim globalVariable As Integer

Sub ProcedureWithGlobalVariable()
    ' 过程级变量
    Dim localVariable As Integer
    localVariable = 10
    
    ' 修改全局变量
    globalVariable = 20
    
    MsgBox "局部变量的值：" & localVariable & vbCrLf & "全局变量的值：" & globalVariable
End Sub
```

在这个示例中，我们定义了一个模块级变量 `globalVariable` 和一个子过程 `ProcedureWithGlobalVariable`。在子过程内部，我们还声明了一个过程级变量 `localVariable`。模块级变量在整个模块中可见，而过程级变量仅在子过程内部可见。

这些示例演示了Excel VBA中函数和子过程的基本概念，包括参数、返回值和作用域。函数用于执行特定的计算并返回结果，而子过程用于执行一系列操作而不返回值。作用域决定了变量的可见性范围。您可以根据需要自定义这些示例代码，并将它们应用于您的Excel VBA项目中。