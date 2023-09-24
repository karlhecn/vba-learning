## 10VBA基础：消息弹窗（MsgBox）与输入弹窗(InputBox)

`MsgBox` 和 `InputBox` 是Excel VBA中常用的用于与用户进行交互的函数，它们用于显示消息框和输入框。下面是它们的示例代码以及讲解：

**1. MsgBox 示例：**

`MsgBox` 用于显示消息框，向用户显示一条消息，等待用户的响应。以下是一个示例：

```vba
Sub ShowMessage()
    MsgBox "Hello, Excel VBA!", vbInformation, "Greeting"
End Sub
```

- `"Hello, Excel VBA!"` 是要显示的消息文本。
- `vbInformation` 指定消息框的图标类型（信息图标）。还可以选择其他图标类型，如感叹号（vbExclamation）
- `"Greeting"` 是消息框的标题。

用户会看到一个带有消息文本 "Hello, Excel VBA!"、信息图标和标题 "Greeting" 的消息框。用户可以点击消息框上的按钮来响应消息。

**2. InputBox 示例：**

`InputBox` 用于显示一个输入框，允许用户输入文本。以下是一个示例：

```vba
Sub GetUserInput()
    Dim userInput As String
    userInput = InputBox("请输入您的姓名:", "输入框示例")
    If userInput <> "" Then
        MsgBox "您输入的姓名是: " & userInput, vbInformation
    Else
        MsgBox "您没有输入姓名。", vbExclamation
    End If
End Sub
```

- `"请输入您的姓名:"` 是输入框中显示的提示文本。
- `"输入框示例"` 是输入框的标题。

在这个示例中，用户会看到一个输入框，要求输入姓名。用户输入的文本会被存储在 `userInput` 变量中。然后，根据用户的输入，程序会显示相应的消息框。

这些示例演示了如何在Excel VBA中使用 `MsgBox` 和 `InputBox` 函数与用户进行交互，将它们应用于您的VBA项目中，以实现更复杂的用户交互和信息提示功能。