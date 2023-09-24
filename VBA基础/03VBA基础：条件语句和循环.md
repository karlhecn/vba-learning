## VBA基础：条件语句和循环



**1. 条件语句 - If...Then...Else：**

```vba
Sub ConditionalStatement()
    ' 声明变量并赋值
    Dim score As Integer
    score = 85
    
    ' 使用条件语句判断分数等级
    If score >= 90 Then
        MsgBox "优秀"
    ElseIf score >= 70 Then
        MsgBox "良好"
    Else
        MsgBox "需要改进"
    End If
End Sub
```

**2. 条件语句 - Select Case：**

```vba
Sub SelectCaseStatement()
    ' 声明变量并赋值
    Dim dayOfWeek As Integer
    dayOfWeek = 3
    
    ' 使用Select Case语句判断星期几
    Select Case dayOfWeek
        Case 1
            MsgBox "星期一"
        Case 2
            MsgBox "星期二"
        Case 3
            MsgBox "星期三"
        Case Else
            MsgBox "其他天"
    End Select
End Sub
```

**3. 循环 - For 循环：**

```vba
Sub ForLoop()
    ' 使用For循环输出1到10的数字
    Dim i As Integer
    For i = 1 To 10
        MsgBox "当前数字是：" & i
    Next i
End Sub
```

**4. 循环 - While 循环：**

```vba
Sub WhileLoop()
    ' 使用While循环输出1到10的偶数
    Dim i As Integer
    i = 2
    While i <= 10
        MsgBox "当前偶数是：" & i
        i = i + 2
    Wend
End Sub
```

**5. 循环 - Do While 循环：**

```vba
Sub DoWhileLoop()
    ' 使用Do While循环输出1到10的奇数
    Dim i As Integer
    i = 1
    Do While i <= 10
        MsgBox "当前奇数是：" & i
        i = i + 2
    Loop
End Sub
```

**6. 循环 - For Each 循环：**

```vba
Sub ForEachLoop()
    ' 声明数组并初始化
    Dim fruits() As String
    fruits = Split("Apple, Banana, Orange", ", ")
    
    ' 使用For Each循环遍历数组元素
    Dim fruit As Variant
    For Each fruit In fruits
        MsgBox "当前水果是：" & fruit
    Next fruit
End Sub
```

**7. 循环 - While 循环：**

```vba
Sub WhileLoopExample()
    ' 使用 While 循环输出 1 到 10 的数字
    Dim i As Integer
    i = 1
    While i <= 10
        MsgBox "当前数字是：" & i
        i = i + 1
    Wend
End Sub
```

