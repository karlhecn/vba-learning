## VBA基础：数组



**1. 声明和初始化数组：**

```vba
Sub DeclareAndInitializeArray()
    ' 声明整数数组并初始化
    Dim numbers(4) As Integer
    numbers(0) = 1
    numbers(1) = 2
    numbers(2) = 3
    numbers(3) = 4
    numbers(4) = 5
    
    ' 访问数组元素并显示
    Dim i As Integer
    For i = 0 To UBound(numbers)
        MsgBox "元素 " & i & ": " & numbers(i)
    Next i
End Sub
```

**2. 动态数组：**

```vba
Sub DynamicArray()
    ' 声明动态整数数组
    Dim dynamicArray() As Integer
    
    ' 重新调整数组大小
    ReDim dynamicArray(4)
    
    ' 初始化数组
    dynamicArray(0) = 10
    dynamicArray(1) = 20
    dynamicArray(2) = 30
    dynamicArray(3) = 40
    dynamicArray(4) = 50
    
    ' 访问数组元素并显示
    Dim i As Integer
    For i = LBound(dynamicArray) To UBound(dynamicArray)
        MsgBox "元素 " & i & ": " & dynamicArray(i)
    Next i
End Sub
```

**3. 数组函数 - LBound 和 UBound：**

```vba
Sub ArrayBounds()
    ' 声明整数数组并初始化
    Dim numbers(3 To 7) As Integer
    numbers(3) = 10
    numbers(4) = 20
    numbers(5) = 30
    numbers(6) = 40
    numbers(7) = 50
    
    ' 使用LBound和UBound函数获取数组的下限和上限
    Dim lowerBound As Integer
    Dim upperBound As Integer
    lowerBound = LBound(numbers)
    upperBound = UBound(numbers)
    
    MsgBox "数组下限：" & lowerBound & vbCrLf & "数组上限：" & upperBound
End Sub
```

**4. 数组函数 - Split：**

```vba
Sub SplitArray()
    ' 声明字符串
    Dim text As String
    text = "Apple, Banana, Orange"
    
    ' 使用Split函数将字符串分割为数组
    Dim fruits() As String
    fruits = Split(text, ", ")
    
    ' 遍历数组并显示
    Dim i As Integer
    For i = LBound(fruits) To UBound(fruits)
        MsgBox "水果 " & (i + 1) & ": " & fruits(i)
    Next i
End Sub
```

**5. 多维数组：**

```vba
Sub MultidimensionalArray()
    ' 声明二维整数数组并初始化
    Dim matrix(1 To 3, 1 To 3) As Integer
    matrix(1, 1) = 1
    matrix(1, 2) = 2
    matrix(1, 3) = 3
    matrix(2, 1) = 4
    matrix(2, 2) = 5
    matrix(2, 3) = 6
    matrix(3, 1) = 7
    matrix(3, 2) = 8
    matrix(3, 3) = 9
    
    ' 访问数组元素并显示
    Dim i As Integer, j As Integer
    For i = 1 To 3
        For j = 1 To 3
            MsgBox "元素 (" & i & ", " & j & "): " & matrix(i, j)
        Next j
    Next i
End Sub
```
