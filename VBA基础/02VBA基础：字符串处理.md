## VBA基础：字符串处理



**1. 字符串长度 (Len)：**

```vba
Sub StringLength()
    Dim text As String
    text = "Hello, World!"
    
    ' 获取字符串长度
    Dim length As Integer
    length = Len(text)
    
    MsgBox "字符串长度为：" & length
End Sub
```

**2. 字符串连接 (Concatenation)：**

```vba
Sub StringConcatenation()
    Dim str1 As String
    Dim str2 As String
    
    str1 = "Hello, "
    str2 = "World!"
    
    ' 连接两个字符串
    Dim combinedString As String
    combinedString = str1 & str2
    
    MsgBox combinedString
End Sub
```

**3. 字符串截取 (Left 和 Right)：**

```vba
Sub StringSubstring()
    Dim text As String
    text = "Hello, World!"
    
    ' 截取左边的字符
    Dim leftSubstring As String
    leftSubstring = Left(text, 5)
    
    ' 截取右边的字符
    Dim rightSubstring As String
    rightSubstring = Right(text, 6)
    
    MsgBox "左边截取：" & leftSubstring & vbCrLf & "右边截取：" & rightSubstring
End Sub
```

**4. 字符串查找 (InStr)：**

```vba
Sub StringSearch()
    Dim text As String
    text = "Hello, World!"
    
    ' 查找子字符串的位置
    Dim position As Integer
    position = InStr(text, "World")
    
    MsgBox "子字符串的位置：" & position
End Sub
```

**5. 字符串替换 (Replace)：**

```vba
Sub StringReplace()
    Dim text As String
    text = "Hello, World!"
    
    ' 替换字符串中的子字符串
    Dim newText As String
    newText = Replace(text, "World", "Universe")
    
    MsgBox newText
End Sub
```

**6. 字符串分割 (Split)：**

```vba
Sub StringSplit()
    Dim text As String
    text = "Apple, Banana, Orange"
    
    ' 使用逗号分割字符串
    Dim parts() As String
    parts = Split(text, ", ")
    
    ' 遍历分割后的数组
    Dim i As Integer
    For i = LBound(parts) To UBound(parts)
        MsgBox "分割后的部分 " & (i + 1) & ": " & parts(i)
    Next i
End Sub
```

