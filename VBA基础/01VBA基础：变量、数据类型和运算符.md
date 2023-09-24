## VBA基础：变量、数据类型和运算符



**1. 声明变量和赋值：**

```vba
Sub DeclareAndAssignVariables()
    ' 声明整数类型变量并赋值
    Dim intValue As Integer
    intValue = 10
    
    ' 声明字符串类型变量并赋值
    Dim strValue As String
    strValue = "Hello, World!"
    
    ' 声明浮点数类型变量并赋值
    Dim floatValue As Double
    floatValue = 3.14

    ' 声明布尔类型变量并赋值
    Dim boolValue As Boolean
    boolValue = False
    
    ' 声明日期类型变量并赋值
    Dim dateValue As Date
    dateValue = #9/24/2023 10:00:00 AM#
End Sub
```

**2. 数据类型转换：**

```vba
Sub DataTypeConversion()
    ' 声明变量并赋值
    Dim intValue As Integer
    intValue = 10
    
    ' 将整数类型转换为浮点数类型
    Dim floatValue As Double
    floatValue = CDbl(intValue)
    
    ' 将浮点数类型转换为字符串类型
    Dim strValue As String
    strValue = CStr(floatValue)
End Sub
```

**3. 基本算术运算：**

```vba
Sub BasicArithmeticOperations()
    ' 声明变量并赋值
    Dim num1 As Double
    num1 = 10
    
    Dim num2 As Double
    num2 = 5
    
    ' 加法
    Dim resultAddition As Double
    resultAddition = num1 + num2
    
    ' 减法
    Dim resultSubtraction As Double
    resultSubtraction = num1 - num2
    
    ' 乘法
    Dim resultMultiplication As Double
    resultMultiplication = num1 * num2
    
    ' 除法
    Dim resultDivision As Double
    resultDivision = num1 / num2
End Sub
```

**4. 字符串操作：**

```vba
Sub StringOperations()
    ' 声明字符串变量
    Dim str1 As String
    Dim str2 As String
    
    ' 字符串连接
    str1 = "Hello, "
    str2 = "World!"
    Dim concatenatedString As String
    concatenatedString = str1 & str2
    
    ' 字符串长度
    Dim strLength As Integer
    strLength = Len(concatenatedString)
    
    ' 字符串截取
    Dim subString As String
    subString = Left(concatenatedString, 5) ' 截取前5个字符
End Sub
```

**5. 逻辑运算：**

```vba
Sub LogicalOperations()
    ' 声明变量并赋值
    Dim x As Integer
    x = 10
    
    Dim y As Integer
    y = 5
    
    ' 等于
    Dim isEqual As Boolean
    isEqual = (x = y)
    
    ' 大于
    Dim isGreaterThan As Boolean
    isGreaterThan = (x > y)
    
    ' 逻辑与
    Dim logicalAnd As Boolean
    logicalAnd = (x > 0 And y > 0)
    
    ' 逻辑或
    Dim logicalOr As Boolean
    logicalOr = (x > 0 Or y > 0)
    
    ' 逻辑非
    Dim logicalNot As Boolean
    logicalNot = Not (x > y)
End Sub
```

