Attribute VB_Name = "模块10"
''第一讲 什么VBA

Sub test()
 Range("a1") = 100
End Sub


Sub 输入100()
''
'' 输入100 Macro
'' 宏由 Lenovo User 录制，时间: 2011-4-22
''

''
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "100"
    Range("B4").Select
End Sub
Sub 删除A1的内容()
''
'' 删除A1的内容 Macro
'' 宏由 Lenovo User 录制，时间: 2011-4-22
''

''
    Range("A1").Select
    Selection.ClearContents
End Sub

Sub 输入()
  Range("b2") = ""
End Sub

’第二讲 VBA语句对象方法和属性

''VBA对象

  ''VBA中的对象其实就是我们操作的具有方法、属性的excel中支持的对象

''Excel中的几个常用对象表示方法

 ''1、工作簿
 
      '' Workbooks 代表工作簿集合，所有的工作簿,Workbooks(N)，表示已打开的第N个工作簿
      '' Workbooks ("工作簿名称")
      '' ActiveWorkbook 正在操作的工作簿
      '' ThisWorkBook ''代码所在的工作簿
      
  ''2、工作表
    '' '''sheets("工作表名称")
      '''sheet1 表示第一个插入的工作表,Sheet2表示第二个插入的工作表....
      '''sheets(n) 表示按排列顺序，第n个工作表
      ''ActiveSheet 表示活动工作表，光标所在工作表
      ''worksheet 也表示工作表，但不包括图表工作表、宏工作表等。

  ''3、单元格
       ''cells 所有单元格
       ''Range ("单元格地址")
       ''Cells(行数,列数)
       ''Activecell 正在选中或编辑的单元格
       '''selection 正被选中或选取的单元格或单元格区域
''一、VBA属性

    ''VBA属性就是VBA对象所具有的特点
    ''表示某个对象的属性的方法是
        
        ''对象.属性=属性值
        
    Sub ttt()
      Range("a1").Value = 100
    End Sub

    Sub ttt1()
      Sheets(1).Name = "工作表改名了"
    End Sub

    Sub ttt2()
    
       Sheets("Sheet2").Range("a1").Value = "abcd"
    
    End Sub
    
    
    Sub ttt3()
     
      Range("A2").Interior.ColorIndex = 3
      
    End Sub


''二 、VBA方法

   ''VBA方法是作用于VBA对象上的动作
     
     ''表示用某个方法作用于VBA的对象上，可以用下面的格式：

        
  Sub ttt4()
  
      牛排.做 熟的程度:=七成熟
     
      Range("A1").Copy Range("A2")
  End Sub
   
  Sub ttt5()
  
    Sheet1.Move before:=Sheets("Sheet3")
    
  End Sub
        
''VBA中的代码的基本结构与组成部分




''VBA语句
''一、宏程序语句
  ''运行后可以完成一个功能

Sub test()  ''开始语句
  
  Range("a1") = 100

End Sub   ''结束语句


''二、函数程序语句
   
   ''运行后可以返回一个值
   
Function shcount()

  shcount = Sheets.Count
  
End Function


''三、在程序中应用的语句

  Sub test2()
    
    Call test
    
  End Sub

 Sub test3()
 
   For x = 1 To 100   ''for next 循环语句
      Cells(x, 1) = x
   Next x
 
 End Sub
 

''第三讲 判断语句

Sub 判断1() ''单条件判断
  If Range("a1").Value > 0 Then
     Range("b1") = "正数"
  Else
     Range("b1") = "负数或0"
  End If
End Sub

Sub 判断2() ''多条件判断
  If Range("a1").Value > 0 Then
     Range("b1") = "正数"
  ElseIf Range("a1") = 0 Then
     Range("b1") = "等于0"
  ElseIf Range("B1") <= 0 Then
     Range("b1") = "负数"
  End If
End Sub

Sub 多条件判断2()
 If Range("a1") <> "" And Range("a2") <> "" Then
   Range("a3") = Range("a1") * Range("a2")
 End If
End Sub

 Sub 判断4()
  Range("a3") = IIf(Range("a1") <= 0, "负数或零", "负数")
End Sub



Sub 判断1() ''单条件判断
  Select Case Range("a1").Value
  Case Is > 0
     Range("b1") = "正数"
  Case Else
     Range("b1") = "负数或0"
  End Select
End Sub

Sub 判断2() ''多条件判断
  Select Case Range("a1").Value
  Case Is > 0
     Range("b1") = "正数"
  Case Is = 0
     Range("b1") = "0"
  Case Else
     Range("b1") = "负数"
  End Select
End Sub

Sub 判断3()
 If Range("a3") < "G" Then
   MsgBox "A-G"
 End If
End Sub



Sub if区间判断()
If Range("a2") <= 1000 Then
  Range("b2") = 0.01
ElseIf Range("a2") <= 3000 Then
  Range("b2") = 0.03
ElseIf Range("a2") > 3000 Then
  Range("b2") = 0.05
End If
End Sub

Sub select区间判断()
 Select Case Range("a2").Value
 Case 0 To 1000
   Range("b2") = 0.01
 Case 1001 To 3000
   Range("b2") = 0.03
 Case Is > 3000
   Range("b2") = 0.05
 End Select
End Sub



''第四讲 循环语句

Sub t1()
  Range("d2") = Range("b2") * Range("c2")
  Range("d3") = Range("b3") * Range("c3")
  Range("d4") = Range("b4") * Range("c4")
  Range("d5") = Range("b5") * Range("c5")
  Range("d6") = Range("b6") * Range("c6")
End Sub

Sub t2()
Dim x As Integer
 For x = 10000 To 2 Step -3
  Range("d" & x) = Range("b" & x) * Range("c" & x)
 Next x
End Sub


Sub t3()
Dim rg As Range
 For Each rg In Range("d2:d18")
  rg = rg.Offset(0, -1) * rg.Offset(0, -2)
 Next rg
End Sub


Sub t4()
Dim x As Integer
 x = 1
 Do
   x = x + 1
   Cells(x, 4) = Cells(x, 2) * Cells(x, 3)
 Loop Until x = 18
End Sub

Sub t5()
 x = 1
 Do While x < 18
   x = x + 1
   Cells(x, 4) = Cells(x, 2) * Cells(x, 3)
 Loop
End Sub
    
  
Sub s1()
 Dim rg As Range
 For Each rg In Range("a1:b7,d5:e9")
   If rg = "" Then
     rg = 0
   End If
  Next rg
End Sub

Sub s2()
 Dim x As Integer
 Do
   x = x + 1
   If Cells(x + 1, 1) <> Cells(x, 1) + 1 Then
      Cells(x, 2) = "断点"
      Exit Do
   End If
 Loop Until x = 14
End Sub
  


''第五讲 变量

Dim m As Integer
''变量
''一、什么是变量？
''所谓变量，就是可变的量。就好象在内存中临时存放的一个小盒子，这个小盒子放的什么物体不固定。
Sub t1()
  Dim X As Integer ''x就是一个变量
  For X = 1 To 10
    Cells(X, 1) = X
  Next X
End Sub

''二、小盒子里可以放什么？
   ''1 放数字
     ''如t1
     
    ''2 放文本
    Sub t2()
     Dim st As String
     Dim X As Integer
     For X = 1 To 10
      st = st & "Excel精英培训"
     Next X
    End Sub
    
     ''3 放对象
     
      Sub t3()
        Dim rg As Range
        Set rg = Range("a1")
        rg = 100
      End Sub
      
      
      ''4 放数组
       Sub t4()
       
        Dim arr(1 To 10) As Integer, X As Integer
        For X = 1 To 10
          arr(X) = X
        Next X
        
       End Sub
 

''三、变量的类型和声明
 
   ''1 变量的类型
      
       ''详见帮助文件
       
   ''2 为什么要声明变量
   
    
   ''3 声明变量
      
      ''dim public
   
    
''四、变量的存活周期
   
   ''1 过程级变量:过程结束，变量值释放
       
       ''如t1
   
   ''2 模块级变量:变量的值只在本模块中保持，工作簿关闭时随时释放
       ''例5
         Sub t6()
            m = 1
         End Sub
         Sub t5()
          MsgBox m
          m = 7
         End Sub
   ''3 全局级变量: 在所有的模块中都可以调用，值会保存到EXCEL关闭时才会被释放。
       
       '' public 变量
       
         Sub t7()
           MsgBox qq
         End Sub
''五 变量的释放
    
     ''一般情况下，过程级变量在过程运行结束后就会自动从内存中释放，而只有一些从外部借用的对象变量才需要使用set 变量=nothing进行释放。



Public qq As Integer

Sub DD()
  qq = 12
End Sub


Option Explicit


’第六讲 公式与函数

Option Explicit

''一、在单元格中输入公式

''1、用VBA在单元格中输入普通公式

     Sub t1()
       Range("d2") = "=b2*c2"
     End Sub
     
     Sub t2()
      Dim x As Integer
      For x = 2 To 6
       Cells(x, 4) = "=b" & x & "*c" & x
      Next x
     End Sub

''2、用VBA在单元格输入带引号的公式
     Sub t3()
     
     Range("c16") = "=SUMIF(A2:A6,""b"",B2:B6)" ''遇到单引号就把单引号加倍
     
     End Sub
      
''3、用VBA在单元格中输入数组公式

    Sub t4()
      Range("c9").FormulaArray = "=SUM(B2:B6*C2:C6)"
    End Sub
    
''二、利用单元格公式返回值

     Sub t5()
         Range("d16") = Evaluate("=SUMIF(A2:A6,""b"",B2:B6)")
         Range("d9") = Evaluate("=SUM(B2:B6*C2:C6)")
     End Sub
  
''三、借用工作表函数
    
     Sub t6()
        
        Range("d8") = Application.WorksheeFunction.CountIf(Range("A1:A10"), "B")
        
     End Sub

''四、利用VBA函数

     Sub t7()
     
      Range("C20") = VBA.InStr(Range("a20"), "E")

     End Sub
     
   

''五、编写自定义函数

      Function wn()
         wn = Application.Caller.Parent.Name
      End Function
     

''第七讲 VBE编辑器

''VBA第七集：VBE编辑器

''一、VBE的窗口
 ''1、工程窗口
 
    ''A 显示工作簿工作表对象
    ''B 窗体
    ''C 模块
    ''D 类模块

 ''range("a1")=10
    
    ''对应工程窗口的对象和模板，显示其所具体的一些特征。
     
 ''3、代码窗口
    ''A 注释文字的设置
    ''B 代码缩进的设置
    ''C 代码强制转行的设置
    ''D 代码运行和调试
         ''逐句运行
         ''设置断点
    ''E 对象列表框和过程列表框
 ''4、立即窗口
 
 ''立即窗口可以把运行过程中的值立即显示出来，主要用于程序的调试

Sub d()
 Dim x As Integer, st As String
 For x = 1 To 10
    st = st & Cells(x, 1)
    Debug.Print "第" & x & "次运行结果:" & st
 Next x
End Sub

 ''5、本地窗口

   ''在本地窗口中可以显示运行中断时对象信息、变量值、数组信息等。

 Sub d1()
 Dim x As Integer, k As Integer
 For x = 1 To 10
   k = k + Cells(x, 1)
  '' If k > 26 Then
  '' Stop
  '' End If
 Next x
 End Sub

''第八讲 分支与end语句

''VBA第七集：VBE编辑器

''一、VBE的窗口
 ''1、工程窗口
 
    ''A 显示工作簿工作表对象
    ''B 窗体
    ''C 模块
    ''D 类模块

 ''range("a1")=10
    
    ''对应工程窗口的对象和模板，显示其所具体的一些特征。
     
 ''3、代码窗口
    ''A 注释文字的设置
    ''B 代码缩进的设置
    ''C 代码强制转行的设置
    ''D 代码运行和调试
         ''逐句运行
         ''设置断点
    ''E 对象列表框和过程列表框
 ''4、立即窗口
 
 ''立即窗口可以把运行过程中的值立即显示出来，主要用于程序的调试

Sub d()
 Dim x As Integer, st As String
 For x = 1 To 10
    st = st & Cells(x, 1)
    Debug.Print "第" & x & "次运行结果:" & st
 Next x
End Sub

 ''5、本地窗口

   ''在本地窗口中可以显示运行中断时对象信息、变量值、数组信息等。

 Sub d1()
 Dim x As Integer, k As Integer
 For x = 1 To 10
   k = k + Cells(x, 1)
  '' If k > 26 Then
  '' Stop
  '' End If
 Next x
 End Sub

Option Explicit

''Goto语句,跳转到指定的地方

Sub t1()
  Dim x As Integer
  Dim sr
100:
  sr = Application.InputBox("请输入数字", "输入提示")
  If Len(sr) = 0 Or Len(sr) = 5 Then GoTo 100
  
End Sub

''gosub..return ,跳过去,再跳回来

Sub t2()
  Dim x As Integer
  For x = 1 To 10
     If Cells(x, 1) Mod 2 = 0 Then GoSub 100
  Next x
Exit Sub
100:
   Cells(x, 1) = "偶数"
   Return
End Sub

''on error resume next ''遇到错误,跳过继续执行下一句

 Sub t3()
  On Error Resume Next
  Dim x As Integer
  For x = 1 To 10
    Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
  Next x
 End Sub
 
''on error goto  ''出错时跳到指定的行数
 
  Sub t4()
  On Error GoTo 100
  Dim x As Integer
  For x = 1 To 10
    Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
  Next x
   Exit Sub
100:
   MsgBox "在第" & x & "行出错了"
  End Sub
 
 ''on error goto 0 ''取消错误跳转
 
  Sub t5()
  On Error Resume Next
  Dim x As Integer
  For x = 1 To 10
    If x > 5 Then On Error GoTo 0
    Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
  Next x
   Exit Sub

  End Sub

''第9集 excel文件操作

Option Explicit

''excel文件和工作簿

  ''excel文件就是excel工作簿，excel文件打开需要excel程的支持
  
   ''Workbooks  工作簿集合，泛指excel文件或工作簿
   
   ''Workbooks("A.xls")，名称为A的excel工作簿
     Sub t1()
        Workbooks("A.xls").Sheets(1).Range("a1") = 100
     End Sub
   
   ''workbooks(2)，按打开顺序，第二个打开的工作簿。
      Sub t2()
        Workbooks(2).Sheets(2).Range("a1") = 200
     End Sub
   ''ActiveWorkbook ，当打开多个excel工作簿时，你正在操作的那个就是ActiveWorkbook（活动工作簿）
   
   ''Thisworkbook，VBA程序所在的工作簿，无论你打开多少个工作簿，无论当前是哪个工作簿是活动的,thisworkbook就是指它所在的工作簿。

''工作簿窗口

    ''Windows("A.xls"),A工作簿的窗口，使用windows可以设置工作簿窗口的状态，如是否隐藏等。
     Sub t3()
        Windows("A.xls").Visible = False
     End Sub
     Sub t4()
        Windows(2).Visible = True
     End Sub
    

   


Option Explicit

''Goto语句,跳转到指定的地方

Sub t1()
  Dim x As Integer
  Dim sr
100:
  sr = Application.InputBox("请输入数字", "输入提示")
  If Len(sr) = 0 Or Len(sr) = 5 Then GoTo 100
  
End Sub

''gosub..return ,跳过去,再跳回来

Sub t2()
  Dim x As Integer
  For x = 1 To 10
     If Cells(x, 1) Mod 2 = 0 Then GoSub 100
  Next x
Exit Sub
100:
   Cells(x, 1) = "偶数"
   Return
End Sub

''on error resume next ''遇到错误,跳过继续执行下一句

 Sub t3()
  On Error Resume Next
  Dim x As Integer
  For x = 1 To 10
    Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
  Next x
 End Sub
 
''on error goto  ''出错时跳到指定的行数
 
  Sub t4()
  On Error GoTo 100
  Dim x As Integer
  For x = 1 To 10
    Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
  Next x
   Exit Sub
100:
   MsgBox "在第" & x & "行出错了"
  End Sub
 
 ''on error goto 0 ''取消错误跳转
 
  Sub t5()
  On Error Resume Next
  Dim x As Integer
  For x = 1 To 10
    If x > 5 Then On Error GoTo 0
    Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
  Next x
   Exit Sub

  End Sub


''第10集  excel工作表操作

Option Explicit

''excel工作表的分类


  ''excel工作表有两大类，一类是我们平常用的工作表(worksheet)，另一类是图表、宏表等。这两类的统称是sheets
  
   '''sheets  工作表集合，泛指excel各种工作表
   
   '''sheets("A")，名称为A的excel工作表
     Sub t1()
        Sheets("A").Range("a1") = 100
     End Sub
   
   ''workbooks(2)，按打开顺序，第二个打开的工作簿。
      Sub t2()
        Sheets(2).Range("a1") = 200
     End Sub
   ''ActiveSheet ，当打开多个excel工作簿时，你正在操作的那个就是ActiveSheet
   
   
Option Explicit

''1 判断A工作表文件是否存在
    Sub s1()
     Dim X As Integer
      For X = 1 To Sheets.Count
        If Sheets(X).Name = "A" Then
          MsgBox "A工作表存在"
          Exit Sub
        End If
      Next
      MsgBox "A工作表不存在"
    End Sub
   
''2 excel工作表的插入

  Sub s2()
     Dim sh As Worksheet
     Set sh = Sheets.Add
       sh.Name = "模板"
       sh.Range("a1") = 100
  End Sub

''3 excel工作表隐藏和取消隐藏
  
 Sub s3()
    Sheets(2).Visible = True
 End Sub

''4 excel工作表的移动

   Sub s4()
     Sheets("Sheet2").Move before:=Sheets("sheet1") '''sheet2移动到sheet1前面
     Sheets("Sheet1").Move after:=Sheets(Sheets.Count) '''sheet1移动到所有工作表的最后面
   End Sub
  
''6 excel工作表的复制
   Sub s5() ''在本工作簿中
      Dim sh As Worksheet
      Sheets("模板").Copy before:=Sheets(1)
       Set sh = ActiveSheet
          sh.Name = "1日"
          sh.Range("a1") = "测试"
   End Sub
   
   Sub s6() ''另存为新工作簿
      Dim wb As Workbook
       Sheets("模板").Copy
       Set wb = ActiveWorkbook
          wb.SaveAs ThisWorkbook.Path & "/1日.xls"
          wb.Sheets(1).Range("b1") = "测试"
          wb.Close True
   End Sub
''7 保护工作表
   Sub s7()
      Sheets("sheet2").Protect "123"
   End Sub
   Sub s8() ''判断工作表是否添加了保护密码
      If Sheets("sheet2").ProtectContents = True Then
        MsgBox "工作簿保护了"
      Else
        MsgBox "工作簿没有添加保护"
      End If
   End Sub
   
 ''8 工作表删除
     Sub s9()
       Application.DisplayAlerts = False
         Sheets("模板").Delete
       Application.DisplayAlerts = True
     End Sub
''9 工作表的选取
     Sub s10()
       Sheets("sheet2").Select
     End Sub

  
''习题

Option Explicit

Sub 日报表格式生成()
      Dim sh As Worksheet
      Dim co As String
      Sheets("日报表模板").Visible = True
      Sheets("日报表模板").Copy after:=Sheets(Sheets.Count)
     co = Sheets.Count - 2
      If co > 31 Then co = 1
    Sheets("日报表模板 (2)").Name = co & "日报表"
    Sheets("日报表模板").Visible = False
End Sub

Sub 日报表格式生成1()
''来自网络
Dim i As Integer

Dim ws As Worksheet
Dim a As String

Set ws = Sheets("日报表模板")

ws.Visible = -1

i = Val(Sheets(Sheets.Count).Name)
a = Sheets(Sheets.Count).Name
ws.Copy after:=Sheets(Sheets.Count)

If i Then

ActiveSheet.Name = i + 1 & "日报表"

Else

ActiveSheet.Name = "1日报表"

End If

ws.Visible = 0

Sheets(1).Select

End Sub


Sub 另存报表()
      On Error Resume Next
      Application.DisplayAlerts = False
      Dim wb As Workbook
      Dim x As Integer
      Dim i As String
      x = 1
      Do While x < Sheets.Count
        i = CStr(x) & "日报表"
       Sheets(i).Select
       Sheets(i).Copy
       Set wb = ActiveWorkbook
       wb.SaveAs ThisWorkbook.Path & "/" & Sheets(i).Name & ".xls"
    x = x + 1
     wb.Close True
    Loop
    Application.DisplayAlerts = True
End Sub


Sub 另存报2()
''来自网络
Dim i As Integer

Dim sh As Worksheet

For i = 1 To Sheets.Count

If Sheets(i).Name Like "*日报表" Then

Sheets(i).Copy

Set sh = ActiveSheet

sh.SaveAs ThisWorkbook.Path & "\" & sh.Name ''& ".xls"

ActiveWorkbook.Close True

End If

Next

End Sub


''第十一讲 单元格的选取

Option Explicit


''1 表示一个单元格(a1)
 Sub s()
   Range("a1").Select
   Cells(1, 1).Select
   Range("A" & 1).Select
   Cells(1, "A").Select
   Cells(1).Select
   [a1].Select
 End Sub


''2 表示相邻单元格区域
   
   
   Sub d() ''选取单元格a1:c5
''     Range("a1:c5").Select
''     Range("A1", "C5").Select
''     Range(Cells(1, 1), Cells(5, 3)).Select
     ''Range("a1:a10").Offset(0, 1).Select
      Range("a1").Resize(5, 3).Select
   End Sub
   
''3 表示不相邻的单元格区域
   
    Sub d1()
    
      Range("a1,c1:f4,a7").Select
      
      ''Union(Range("a1"), Range("c1:f4"), Range("a7")).Select
      
    End Sub
    
    Sub dd() ''union示例
      Dim rg As Range, x As Integer
      For x = 2 To 10 Step 2
        If x = 2 Then Set rg = Cells(x, 1)
        
        Set rg = Union(rg, Cells(x, 1))
      Next x
      rg.Select
    End Sub
    
''4 表示行
  
    Sub h()
    
      ''Rows(1).Select
      ''Rows("3:7").Select
      ''Range("1:2,4:5").Select
       Range("c4:f5").EntireRow.Select
       
    End Sub
    
''5 表示列
    
   Sub L()
    
      '' Columns(1).Select
      '' Columns("A:B").Select
      '' Range("A:B,D:E").Select
      Range("c4:f5").EntireColumn.Select ''选取c4:f5所在的行
       
   End Sub

''6 重置坐标下的单元格表示方法

    Sub cc()
    
      Range("b2").Range("a1") = 100
      
    End Sub
    
''7 表示正在选取的单元格区域

   Sub d2()
     Selection.Value = 100
   End Sub

''习题

''Option Explicit

Sub 任选条件填充()
Dim A As Range '', B
For Each A In Selection
    If IsNumeric(A) And A > 0 Then
      A = "正数"
    End If
Next
End Sub

Sub 选取()
Dim myrange As Range
Dim i As Integer
Dim j As Integer
Dim n As Integer
For i = 1 To 12
    For j = 1 To 3
        If IsNumeric(Cells(i, j)) And Cells(i, j).Value > 0 Then
            n = n + 1
            If n = 1 Then Set myrange = Cells(i, j).EntireRow
            Set myrange = Union(myrange, Cells(i, j).EntireRow)
        End If
    Next j
Next i
myrange.Select
End Sub

Sub db() ''bajifeng
Dim rng As Range
For i = 1 To 12
    For j = 1 To 3
        If IsNumeric(Cells(i, j)) And Cells(i, j).Value > 0 Then
            n = n + 1
            If n = 1 Then Set rng = Cells(i, 1).Resize(1, 3)
            Set rng = Union(rng, Cells(i, 1).Resize(1, 3))
        End If
    Next
Next
rng.Select
End Sub


''第12集 特殊单元格定位

Option Explicit


''1 已使用的单元格区域

  Sub d1()
  
    Sheets("sheet2").UsedRange.Select
    
    ''wb.Sheets(1).Range("a1:a10").Copy Range("i1")
    
  End Sub


''2 某单元格所在的单元格区域

   Sub d2()
    
      Range("b8").CurrentRegion.Select
    
   End Sub
   
   
''3 两个单元格区域共同的区域

    Sub d3()
     
    Intersect(Columns("b:c"), Rows("3:5")).Select
  
    End Sub
   
''4 调用定位条件选取特殊单元格
  
    Sub d4()
  
       Range("A1:A6").SpecialCells(xlCellTypeBlanks).Select
       
    End Sub
    
''5 端点单元格
 
   Sub d5()
   
     Range("a65536").End(xlUp).Offset(1, 0) = 1000
     
   End Sub
  
   Sub d6()
   
     Range(Range("b6"), Range("b6").End(xlToRight)).Select
     
   End Sub
    
''实例
Option Explicit

Sub t()
 Dim x As Integer
  For x = 2 To 6
    If Cells(x, 2) > 0 Then
      Cells(x, "N") = "1月"
    Else
      Cells(x, "N") = Range("b" & x).End(xlToRight).Column - 1 & "月"
    End If
  Next x
  
End Sub


 ''习题
 Option Explicit

''题目1:
''   B:D列各行单元格,如果为非空,则在该行A列填充数字1,如A列所示,
''
''要求: 不得使用循环
''题目2:
''    打开本路径下的A.Xls文件 , 并把文件中的所有工作表的明细数据合并到本表中, 上下排列
''
''注:     A.xls文件中工作表数量和明细表行数和列数均不定.但各个工作表中的行列数量相同
Sub 第一题()
''    Intersect(Columns(1), Range("B:D"). _  ''分隔符的用法，连到下一行
''    SpecialCells(xlCellTypeConstants).EntireRow) = 1
    Intersect(Columns(1), Range("B:D").SpecialCells(xlCellTypeConstants).EntireRow) = 1
End Sub
Sub 第2题()
Dim i As Integer, wbk As Workbook
Set wbk = Workbooks.Open(ThisWorkbook.Path & "\A.xls")
With ThisWorkbook.Sheets("第2题")
For i = 1 To wbk.Sheets.Count
    If i = 1 Then
       wbk.Sheets(i).UsedRange.Copy .Range("A1") ''把标题考虑进去了
    Else
       wbk.Sheets(i).UsedRange.Offset(1, 0).Copy .Range("A" & .[A65536].End(xlUp).Row + 1) ''直接在单元格里面把坐标做好
    End If
Next i
wbk.Close True
End With
End Sub
Sub se2()
    Dim wb As Workbook, i As Integer
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\A.xls")
    For i = 1 To wb.Sheets.Count
        wb.Sheets(i).UsedRange.Offset(1, 0).Copy ThisWorkbook.Sheets("第2题").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0) ''往下移动一格，不然覆盖了
    Next i
    wb.Close
End Sub
Sub copy用法()
    ''Worksheets(1).Range("A1:D4").Copy Destination:=Worksheets(2).Range("E1")  ''Destination:=可以省略
    Worksheets(1).Range("A1:D4").Copy Worksheets(2).Range("E1")
End Sub
Sub fuzhi()
''      Dim sh As Worksheet
''      Workbooks("第12集练习题.xls").Sheets("第2题").Copy after:=Workbooks("第12集练习题0.xls").Sheets(1)

    Workbooks(1).Sheets(2).UsedRange.Offset(0, 0).Select
End Sub



''第13讲 Option Explicit

''1 单元格的值

   Sub x1()
    Range("b7") = Range("I3").Value
    Range("b8") = Range("c2").Text
    Range("b9") = "''" & Range("I3").Formula
   End Sub

 ''2 单元格的地址
   
    Sub x2()
     With Range("b2").CurrentRegion
       [b12] = .Address
       [c12] = .Address(0, 0)
       [d12] = .Address(1, 0)
       [e12] = .Address(0, 1)
       [f12] = .Address(1, 1)
     End With
    End Sub
 
 ''3 单元格的行列信息
    Sub x3()
      With Range("b2").CurrentRegion
        [b13] = .Row
        [b14] = .Rows.Count
        [b15] = .Column
        [b16] = .Columns.Count
        [b17] = .Range("a1").Address
      End With
    End Sub
     
 ''4、单元格的格式信息
    Sub x4()
      With Range("b2")
        [b19] = .Font.Size
        [b20] = .Font.ColorIndex
        [b21] = .Interior.ColorIndex
        [b22] = .Borders.LineStyle
      End With
    End Sub
       
  ''5、单元格批注信息
     Sub x5()
        [B24] = Range("I2").Comment.Text
     End Sub

  ''6 单元格的位置信息
     Sub x6()
        With Range("B2")
          [b26] = .Top
          [b27] = .Left
          [b28] = .Height
          [b29] = .Width
        End With
     End Sub
  ''7 单元格的上级信息
    Sub x7()
      With Range("b3")
        [b31] = .Parent.Name
        [b32] = .Parent.Parent.Name
      End With
    End Sub
   ''8 内容判断
      Sub x8()
       With Range("i3")
        [b34] = .HasFormula
        [b35] = .Hyperlinks.Count
       End With
      End Sub
    ''9 单元格数据类型（另讲）
      
    
''习题
Sub 第一题()
      Dim x, y As Integer
      With Range("c4").CurrentRegion
      x = .Rows.Count
      y = .Columns.Count
      [a1] = .Cells(x, y).Address(0, 0)
      End With
End Sub
Sub 第二题()
    Range("F3").Comment.Shape.Left = Range("E1").Left
End Sub


''第14集 单元格格式


''一、判断数值的格式
  ''1 判断是否为空单元格
    Sub d1()
       [b1] = ""
       ''If Range("a1") = "" Then
       ''If Len([a1]) = 0 Then
       If VBA.IsEmpty([a1]) Then
          [b1] = "空值"
        End If
    End Sub
  ''2 判断是否为数字
    Sub d2()
      [b2] = ""
      ''If VBA.IsNumeric([a2]) And [a2] <> "" Then
      ''If Application.WorksheetFunction.IsNumber([a2]) Then
        [b2] = "数字"
      End If
    End Sub
  ''3 判断是否为文本
    Sub d3()
      [b3] = ""
      ''If Application.WorksheetFunction.IsText([A3]) Then
       If VBA.TypeName([a3].Value) = "String" Then
         [b3] = "文本"
      End If
    End Sub
  ''4 判断是否为汉字
     Sub d4()
        [b4] = ""
        If [a4] > "z" Then
          [b4] = "汉字"
        End If
     End Sub
  ''5 判断错误值
  Sub d10()
      [b5] = ""
      ''If VBA.IsError([a5]) Then
      If Application.WorksheetFunction.IsError([a5]) Then
         [b5] = "错误值"
      End If
  End Sub
    Sub d11()
      [b6] = ""
      If VBA.IsDate([a6]) Then
         [b6] = "日期"
      End If
  End Sub

''二、设置单元格自定义格式
   Sub d30()
        Range("d1:d8").NumberFormatLocal = "0.00"
   End Sub


''三、按指定格式从单元格返回数值
   
   ''Format函数语法(和工作表数Text用法基本一致)
   
    ''Format(数值,自定义格式代码)
    

    
    

	Option Explicit
''Excel中的颜色
   
    ''Excel中的颜色可以用两种方式获取，一种是EXCEL内置颜色，另一种是利用QBCOLOR函数返回
  Sub y1()
   Dim x As Integer
    Range("a1:b60").Clear
    For x = 1 To 56
      Range("a" & x) = x
     '' Range("b" & x).Font.ColorIndex = 3
      Range("b" & x).Interior.ColorIndex = x
    Next x
  End Sub

   Sub y2()
    Dim x As Integer
     For x = 0 To 15
      Range("d" & x + 1) = x
      Range("e" & x + 1).Interior.Color = QBColor(x)
     Next x
   End Sub

  Sub y3()
    Dim 红 As Integer, 绿 As Integer, 蓝 As Integer
    红 = 255
    绿 = 123
    蓝 = 100
    Range("g1").Interior.Color = RGB(红, 绿, 蓝)
  End Sub

  
''单元格合并

  Sub h1()
    
    Range("g1:h3").Merge
    
  End Sub
  
  ''合并区域的返回信息
  Sub h2()
   
   Range("e1") = Range("b3").MergeArea.Address ''返回单元格所在的合并单元格区域
   
  End Sub
  
  ''判断是否含合并单元格
  Sub h3()
   ''MsgBox Range("b2").MergeCells
    ''MsgBox Range("A1:D7").MergeCells
    Range("e2") = IsNull(Range("a1:d7").MergeCells)
    Range("e3") = IsNull(Range("a9:d72").MergeCells)
  End Sub
  
 ''综合示例
 
   ''合并H列相同单元格
   
     Sub h4()
      Dim x As Integer
      Dim rg As Range
      Set rg = Range("h1")
       Application.DisplayAlerts = False
      For x = 1 To 13
        If Range("h" & x + 1) = Range("h" & x) Then
          Set rg = Union(rg, Range("h" & x + 1))
        Else
         
           rg.Merge
          
          Set rg = Range("h" & x + 1)
        End If
      Next x
      Application.DisplayAlerts = True
     End Sub

	

Option Explicit
Dim a, b, c
Private Sub ScrollBar1_Change()
a = ScrollBar1.Value
 Me.BackColor = RGB(a, b, c)
End Sub

Private Sub ScrollBar2_Change()
 c = ScrollBar2.Value
 Me.BackColor = RGB(a, b, c)
End Sub

Private Sub ScrollBar3_Change()
 b = ScrollBar3.Value
 Me.BackColor = RGB(a, b, c)
End Sub


Private Sub UserForm_Click()

End Sub


	
''第15集

Option Explicit

Sub c1()
  Rows(4).Insert
End Sub

Sub c2() ''插入行并复制公式
  Rows(4).Insert
  Range("3:4").FillDown
  Range("4:4").SpecialCells(xlCellTypeConstants) = ""
End Sub

Sub c3()
  Dim x As Integer
  For x = 2 To 20
    If Cells(x, 3) <> Cells(x + 1, 3) Then
      Rows(x + 1).Insert
      x = x + 1
    End If
  Next x
End Sub

Sub c4()
  Dim x As Integer, m1 As Integer, m2 As Integer
  Dim k As Integer
  m1 = 2
  For x = 2 To 1000
    If Cells(x, 1) = "" Then Exit Sub
    If Cells(x, 3) <> Cells(x + 1, 3) Then
      m2 = x
      Rows(x + 1).Insert
      Cells(x + 1, "c") = Cells(x, "c") & " 小计"
      Cells(x + 1, "h") = "=sum(h" & m1 & ":h" & m2 & ")"
      Cells(x + 1, "h").Resize(1, 4).FillRight
      Cells(x + 1, "i") = ""
      x = x + 1
      m1 = m2 + 2
    End If
  Next x
End Sub

Sub dd() ''删除小计行
 Columns(1).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub

''
Option Explicit

 
''1 单元格输入
 
    Sub t1()
      Range("a1") = "a" & "b"
      Range("b1") = "a" & Chr(10) & "b" ''换行答输入
    End Sub
    
''2 单元格复制和剪切
    
      Sub t2()
        Range("a1:a10").Copy Range("e1") ''A1：A10的内容复制到C1
      End Sub
    
      Sub t3()
        Range("a1:a10").Copy
        ActiveSheet.Paste Range("d1") ''粘贴至D1
      End Sub
      
      Sub t4()
        Range("a1:a10").Copy
        Range("e1").PasteSpecial (xlPasteValues) ''只粘贴为数值
      End Sub
      Sub t5()
        Range("a1:a10").Cut
        ActiveSheet.Paste Range("f1") ''粘贴到f1
      End Sub

      Sub t6()
        Range("c1:c10").Copy
        Range("a1:a10").PasteSpecial Operation:=xlAdd ''选择粘贴-加
      End Sub
      
      Sub T7()
          Range("G1:G10") = Range("A1:A10").Value
      End Sub
''3 填充公式
    Sub T8()
      Range("b1") = "=a1*10"
      Range("b1:b10").FillDown ''向下填充公式
    End Sub
    
''第16集

Option Explicit

''1 使用循环查找 (在单元格中查找效率太低)

''2 调用工作表函数
  
    Sub c1() ''判断是否存在,并查找所在行数
      Dim hao As Integer
      Dim icount As Integer
      icount = Application.WorksheetFunction.CountIf(Sheets("库存明细表").[b:b], [g3])
      If icount > 0 Then
       MsgBox "该入库单号码已经存在，请不要重复录入"
       MsgBox Application.WorksheetFunction.Match([g3], Sheets("库存明细表").[b:b], 0)
      End If
    End Sub
    
    
''3 使用Find方法

    Sub c2()
      Dim r As Integer, r1 As Integer
      Dim icount As Integer
      icount = Application.WorksheetFunction.CountIf(Sheets("库存明细表").[b:b], [g3])
      If icount > 0 Then
       r = Sheets("库存明细表").[b:b].Find(Range("G3"), Lookat:=xlWhole).Row ''查找号码第一次出现的位置
       r1 = Sheets("库存明细表").[b:b].Find([g3], , , , , xlPrevious).Row
       MsgBox r & ":" & r1
      End If
    End Sub
 

   Sub c3() ''返回最下一行非空行的行数
    
      MsgBox Sheets("库存明细表").Cells.Find("*", , , , , xlPrevious).Row
    
   End Sub

   
   
   
   Option Explicit
Sub 输入()
  Dim c As Integer   ''号码在库存表中的个数
  Dim r As Integer   ''入库单的数据行数
  Dim cr As Integer  ''库存明细表中第一个空行的行数
With Sheets("库存明细表")
    c = Application.CountIf(.[b:b], Range("g3"))
    If c > 0 Then
       MsgBox "该单据号码已经存在！,请不要重复录入"
       Exit Sub
    Else
       r = Application.CountIf(Range("b6:b10"), "<>")
       cr = .[b65536].End(xlUp).Row + 1
       .Cells(cr, 1).Resize(r, 1) = Range("e3")
       .Cells(cr, 2).Resize(r, 1) = Range("g3")
       .Cells(cr, 3).Resize(r, 1) = Range("c3")
       .Cells(cr, 4).Resize(r, 6) = Cells(6, 2).Resize(r, 6).Value
       MsgBox "输入已完成"
    End If
 End With
End Sub

Sub 查找()
  Dim c As Integer   ''号码在库存表中的个数
  Dim r As Integer   ''入库单的数据行数
  
With Sheets("库存明细表")
    c = Application.CountIf(.[b:b], Range("g3"))
    If c = 0 Then
       MsgBox "该单据号码不存在！"
       Exit Sub
    Else
        r = .[b:b].Find(Range("g3"), , , , , xlNext).Row
        Range("c3") = .Cells(r, 3)
        Range("e3") = .Cells(r, 1)
        Cells(6, 2).Resize(c, 5) = .Cells(r, 4).Resize(c, 5).Value
       MsgBox "查询已完成"
    End If
 End With
End Sub

Sub 删除()
 Dim c As Integer   ''号码在库存表中的个数
  Dim r As Integer   ''入库单的数据行数
  
With Sheets("库存明细表")
    c = Application.CountIf(.[b:b], Range("g3"))
    If c = 0 Then
       MsgBox "该单据号码不存在！"
       Exit Sub
    Else
        r = .[b:b].Find(Range("g3"), , , , , xlNext).Row
        .Range(r & ":" & c + r - 1).Delete
       MsgBox "删除已完成"
    End If
 End With
End Sub
Sub 修改()
  Call 删除
  Call 输入
End Sub


''第17集


Private Sub Worksheet_Activate()
 If ActiveSheet.name = "Sheet2" Then
    Sheets(1).Select
End If
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
 
End Sub

Private Sub Worksheet_Calculate()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
 Application.EnableEvents = False
  Target = Target * 2
 Application.EnableEvents = True
End Sub



Option Explicit

Private Sub Worksheet_Calculate()
 MsgBox "公式的值发生了改变"
End Sub

Private Sub Worksheet_Deactivate()
  MsgBox "谢谢使用sheet3"
End Sub

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
 MsgBox Target.Address
End Sub

Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
  
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub


''第18集

Private Sub Workbook_Deactivate()

End Sub

Private Sub Workbook_Open()
  UserForm1.Show
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
   ''Cancel = True
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)

End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
 If Sh.name = "Sheet2" Then
  MsgBox Target.Address
  MsgBox Sh.name
 End If
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
  MsgBox "本工作簿禁止插入新工作表"
  Application.DisplayAlerts = False
   Sh.Delete
  Application.DisplayAlerts = True
End Sub

Private Sub Workbook_BeforePrint(Cancel As Boolean)
 MsgBox "此excel文件禁止打印，如需打印请与管理员联系"
 Cancel = True
End Sub


Private Sub Workbook_Activate()
 
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
 MsgBox "你点击保存按钮了"
End Sub

’第19集

Public WithEvents app As Excel.Application

Private Sub app_NewWorkbook(ByVal Wb As Workbook)
 
End Sub

Private Sub app_SheetActivate(ByVal Sh As Object)

End Sub

Private Sub app_WorkbookNewSheet(ByVal Wb As Workbook, ByVal Sh As Object)

End Sub

Private Sub app_WorkbookOpen(ByVal Wb As Workbook)
'' a = Application.InputBox("请输入打开excel程序口令", "安全提示")
'' If a <> 123 Then
''   Wb.Close False
''End If
End Sub

Private Sub Workbook_Open()
 Set app = Excel.Application
End Sub

’第20集

''****************************************************************************************************
''*                              VBA数组教程                                                         *
''*                                       --------excel精英培训网:兰色幻想                           *
''****************************************************************************************************


Sub v4() ''运行时间0.01秒
 Dim t
 t = Timer
 For x = 1 To 100000
   m = m + 1000 ''真接调用内存中的值
 Next x
 MsgBox Timer - t
End Sub

Sub v5() ''运行时间0.5秒
 Dim t
 t = Timer
 For x = 1 To 100000
   m = m + Cells(1, 1) ''调用单元格中的值
 Next x
  MsgBox Timer - t
End Sub


''1、什么是VBA数组呢？
    
    ''VBA数组就是储存一组数据的数据空间?数据类型可以数字,可以是文本,可以是对象,也可以是VBA数组.
    
''2 VBA数组存在形态
     '' VBA数组是以变量形式存放的一个空间,它也有行有列，也可以是三维空间。
           
    ''1) 常量数组
          ''array(1,2)
          ''array(array(1,2,4),array("a","b","c"))
    ''2) 静态数组
          ''x(4) ''有5个位置，编号从0~4
          ''arr(1 to 10) ''有10个位置，编号1~10
          ''arr(1 to 10,1 to 2) ''10行2列的空间，总共20个位置，这是二维数组
          ''arr(1 to 10,1 to 2,1 to 3) ''三维数组，总10*2*3=60个位置。这是三维数组
    ''3)动态数组
          ''arr() ''不知道有多少行多少列


Option Explicit

''向VBA数组中写入数据
   
   ''1、按编号(标)写入和读取
   
     Sub t1() ''写入一维数组
     Dim x As Integer
     Dim arr(1 To 10)
   arr(2) = 190
   arr(10) = 5
     End Sub
  
    Sub t2() ''向二维数组写入数据和读取
     Dim x As Integer, y As Integer
     Dim arr(1 To 5, 1 To 4)
     For x = 1 To 5
       For y = 1 To 4
         arr(x, y) = Cells(x, y)
       Next y
     Next x
    MsgBox arr(3, 1)
    End Sub
    
   ''2、动态数组
       Sub t3()
        Dim arr()
        Dim row
        row = Sheets("sheet2").Range("a65536").End(xlUp).row - 1
        ReDim arr(1 To row)
        For x = 1 To row
           arr(x) = Cells(x, 1)
        Next x
        Stop
       End Sub
       
   ''3、批量写入
    
      Sub t4() ''由常量数组导入
      Dim arr
      arr = Array(1, 2, 3, "a")
      Stop
      End Sub
    
     Sub t5() ''由单元格区域导入
       Dim arr
       arr = Range("a1:d5")
       Stop
     End Sub
	 
	 
''第21集

Option Explicit

''VBA数组
  ''1、在内存中读取
      
       ''在内存中读取后用于继续运算，直接用下面的格式
        
        ''数组变量(5)
        ''数组变量(3,2)
     ''例:
        Sub d1()
         Dim arr, arr1()
         Dim x As Integer, k As Integer, m As Integer
         arr = Range("a1:a10") ''把单元格区域导入内存数组中
         m = Application.CountIf(Range("a1:a10"), ">10") ''计算大于10的个数
         ReDim arr1(1 To m)
         For x = 1 To 10
           If arr(x, 1) > 10 Then
              k = k + 1
              arr1(k) = arr(x, 1)
           End If
         Next x
        End Sub
        
        
  ''2、读取存入单元格中
          
      Sub d2() ''二维数组存入单元格
        Dim arr, arr1(1 To 5, 1 To 1)
        Dim x As Integer
        arr = Range("b2:c6")
        For x = 1 To 5
          arr1(x, 1) = arr(x, 1) * arr(x, 2)
        Next x
        Range("d2").Resize(10) = arr1
      End Sub
      
      Sub d3() ''一维数组存入单元格
        Dim arr, arr1(1 To 5)
        Dim x As Integer
        arr = Range("b2:c6")
        For x = 1 To 5
          arr1(x) = arr(x, 1) * arr(x, 2)
        Next x
        ''Range("a13").Resize(1, 5) = arr1
        Range("d2").Resize(5) = Application.Transpose(arr1)
      End Sub
       
      Sub d4() ''数组部分存入
        Dim arr, arr1(1 To 10000, 1 To 1)
        Dim x As Integer
        arr = Range("b2:c6")
        For x = 1 To 5
          arr1(x, 1) = arr(x, 1) * arr(x, 2)
        Next x
        Range("d2").Resize(5) = arr1
      End Sub

''第22集

Option Explicit

''1、数组的大小
''数组是用编号排序的，那么如何获得一个数组的大小呢

 ''Lbound(数组) 可以获取数组的最小下标(编号)
 ''Ubound(数组) 可以获取数组的最大上标(编号)
 ''Ubound(数组,1) 可以获得数组的行方面(第1维)最大上标
 ''Ubound(数组,2) 可以获得数组的列方向(第2维)的最大上标

Sub d6()
 Dim arr
 Dim k, m
 arr = Range("a2:d5")
 For x = 1 To UBound(arr, 1)
   
 Next x
End Sub


''2、动态数组的动态扩充
   
     ''如果一个数组无法或不方便计算出总的大小，而在一些特殊情况下又不允许有空位。这时我们就需要用动态的导入方法
  ''
     ''ReDim Preserve arr() 可以声明一个动态大小的数组，而且可以保留原来的数值，就相当于厂房小了，可以改扩建增大，但是它只能
        ''让最未维实现动态，如果是一维不存在最未维，只有一维
    
     例子1见sheet1工作表
     
   Sub d7()
   Dim arr, arr1()
     arr = Range("a1:d6")
     Dim x, k
     For x = 1 To UBound(arr)
      If arr(x, 1) = "B" Then
         k = k + 1
         ReDim Preserve arr1(1 To 4, 1 To k)
         arr1(1, k) = arr(x, 1)
         arr1(2, k) = arr(x, 2)
         arr1(3, k) = arr(x, 3)
         arr1(4, k) = arr(x, 4)
      End If
     Next x
    Range("a8").Resize(k, 4) = Application.Transpose(arr1)
   End Sub
   
   Sub d8()
   Dim arr, arr1(1 To 100000, 1 To 4)
     arr = Range("a1:d6")
     Dim x, k
     For x = 1 To UBound(arr)
      If arr(x, 1) = "B" Then
         k = k + 1
         arr1(k, 1) = arr(x, 1)
         arr1(k, 2) = arr(x, 2)
         arr1(k, 3) = arr(x, 3)
         arr1(k, 4) = arr(x, 4)
      End If
     Next x
    Range("a15").Resize(k, 4) = arr1
   End Sub
   
''3 清空数组
     ''清空数组使用earse语句
  Sub d9()
   Dim arr, arr1(1 To 1000, 1 To 1)
   Dim x, m, k
   arr = Range("a1:a16")
   For x = 1 To UBound(arr)
     If arr(x, 1) <> "" Then
        k = k + 1
        arr1(k, 1) = arr(x, 1)
     Else
        m = m + 1
        Range("c1").Offset(0, m).Resize(k) = arr1
        Erase arr1
        k = 0
     End If
   Next x
  End Sub
     

''第23集

Option Explicit

''
''1 数组的最值
    Sub s()
    Dim arr1()
    
    arr1 = Array(1, 12, 4, 5, 19)
    
    MsgBox "1, 12, 4, 5, 19最大值" & Application.Max(arr1)
    MsgBox "1, 12, 4, 5, 19最小值:" & Application.Min(arr1)
    MsgBox "1, 12, 4, 5, 19第二大值：" & Application.Large(arr1, 2)
    MsgBox "1, 12, 4, 5, 19第二小值：" & Application.Small(arr1, 2)
    
    End Sub
    
 ''2、求和

    ''用application.Sum (数组)
    
''3 统计个数

  ''counta和count函数可以统计VBA数组的数字个数及所有已填充内容的个数
  
     Sub s1()
     
     Dim arr1, arr2(0 To 10), x
     arr1 = Array("a", "3", "", 4, 6)
     For x = 0 To 4
       arr2(x) = arr1(x)
     Next x
     
     MsgBox "数组1的数字个数：" & Application.Count(arr2)
     
     MsgBox "数组2的已填充数值的个数" & Application.CountA(arr2)
     
     End Sub
 
 ''3 在数组里查找
     
     Sub s2()
      Dim arr
      On Error Resume Next
      arr = Array("a", "c", "b", "f", "d")
      MsgBox Application.Match("f", arr, 0)
     If Err.Number = 13 Then
        MsgBox "查找不到"
      End If
     End Sub

Option Explicit

'' 1、split函数
     ''按分隔符把字符串截取成VBA数组,该数组是一维数组，编号从0开始
 
     '''split(字符串,分隔符)
   
    Sub t1()
      Dim sr, arr
      sr = "A-BC-FGR-H"
      arr = VBA.Split(sr, "-")
      MsgBox Join(arr, ",")
    End Sub
     
'' 2、Filter函数：
     ''按条件筛选符合条件的值组成一个新的数组

     ''Filter(数组,筛选条件,是/否)
     
     ''注：如果是（true）则返回包含的数组，如果否则返回非包含的数组
    Sub t2()
     Dim arr, arr1, arr2
     arr = Application.Transpose(Range("A2:A10"))
     arr1 = VBA.Filter(arr, "W", True)
     arr2 = VBA.Filter(arr, "W", False)
     Range("B2").Resize(UBound(arr1) + 1) = Application.Transpose(arr1)
     Range("C2").Resize(UBound(arr2) + 1) = Application.Transpose(arr2)
    End Sub
    
''3、index函数：
    ''调用该工作表函数可以把二维数组的某一列或某一行截取出来，构成一个新的数组。
     '' Application.Index(二维数组,0,列数)) 返回二维数组
     '' Application.Index(二维数组,行数,0)) 返回一维数组
    Sub t3()
     Dim arr, arr1, arr2
      arr = Range("a2:d6")
      arr1 = Application.Index(arr, , 1)
      arr2 = Application.Index(arr, 4, 0)
      Stop
    End Sub

''4、vlookup函数
      ''Vlookup函数的第一个参数可以用VBA数组，返回的也是一个VBA数组
    Sub t4()
    Dim arr, arr1
      arr = Range("a2:d6")
      arr1 = Application.VLookup(Array("B", "C"), arr, 4, 0)
    End Sub
''5 Sumif函数和Countif函数
     ''Countif和sumif函数的第二个参数都可以使用数组，所以也可以返回一个VBA数组，如：
     Sub t5()
     Dim T
     T = Timer
       Dim arr
       arr = Application.SumIf(Range("a2:a10000"), Array("B", "C", "G", "R"), Range("B2:B10000"))
     MsgBox Timer - T
     Stop
     End Sub
 
   Sub t55()
     Dim T
     T = Timer
      Dim arr, arr1(1 To 4, 1 To 2), x
      arr1(1, 1) = "B"
      arr1(2, 1) = "C"
      arr1(3, 1) = "G"
      arr1(4, 1) = "R"
     '' arr = Range("a1:d10000")
      For x = 2 To 10000
         Select Case Cells(x, 1)
         Case "B"
            arr1(1, 2) = arr1(1, 2) + Cells(x, 2)
         Case "C"
            arr1(2, 2) = arr1(2, 2) + Cells(x, 2)
         Case "G"
            arr1(3, 2) = arr1(3, 2) + Cells(x, 2)
         Case "R"
            arr1(4, 2) = arr1(4, 2) + Cells(x, 2)
         End Select
      Next x
     MsgBox Timer - T
   End Sub



 

 
''第24集

Option Explicit

Sub 单元格循环()
 Dim x As Integer
 Dim t
 清除颜色
 t = Timer
 For x = 2 To Range("a65536").End(xlUp).Row
   If Range("d" & x) > 500 Then
     Range(Cells(x, 1), Cells(x, 4)).Interior.ColorIndex = 3
   End If
 Next x
 MsgBox Timer - t
End Sub

Sub 清除颜色()
 Range("a:d").Interior.ColorIndex = xlNone
End Sub

Sub 数组方法()
 Dim arr, t
 Dim x As Integer
 Dim sr As String, sr1 As String
 清除颜色
 t = Timer
 arr = Range("d2:d" & Range("a65536").End(xlUp).Row)
 For x = 1 To UBound(arr)
   If x = UBound(arr) And sr <> "" Then Range(Left(sr, Len(sr) - 1)).Interior.ColorIndex = 3
   If arr(x, 1) > 500 Then
      sr1 = sr
      sr = sr & "A" & x + 1 & ":D" & x + 1 & ","
      If Len(sr) > 255 Then
        sr = sr1
        Range(Left(sr, Len(sr) - 1)).Interior.ColorIndex = 3
        sr = ""
      End If
   End If
 Next x
 MsgBox Timer - t
End Sub
Sub 数组方法2()
Dim arr, t
 Dim x As Integer, x1 As Integer
 Dim sr As String, sr1 As String
 清除颜色
 t = Timer
 arr = Range("d2:d" & Range("a65536").End(xlUp).Row)
 For x = 1 To UBound(arr)
   If x = UBound(arr) Then Range(Left(sr, Len(sr) - 1)).Interior.ColorIndex = 3
   If arr(x, 1) > 500 Then
      sr1 = sr
      x1 = x + 1
      Do
        x = x + 1
      Loop Until arr(x, 1) <= 500
      
      sr = sr & "A" & x1 & ":D" & x & ","
      If Len(sr) > 255 Then
        sr = sr1
        x = x1 - 1
        Range(Left(sr, Len(sr) - 1)).Interior.ColorIndex = 3
        sr = ""
      End If
      x = x - 1
   End If
 Next x
 MsgBox Timer - t


End Sub
Sub 数组方法3()
Dim arr, t
 Dim x As Integer, x1 As Integer
 Dim sr As String, sr1 As String
 清除颜色
 t = Timer
 arr = Range("d2:d" & Range("a65536").End(xlUp).Row)
 For x = 1 To UBound(arr)
   If x = UBound(arr) Then Application.Intersect(Range("a:d"), Range(Left(sr, Len(sr) - 1))).Interior.ColorIndex = 3
   If arr(x, 1) > 500 Then
      sr1 = sr
      x1 = x + 1
      Do
        x = x + 1
      Loop Until arr(x, 1) <= 500
      
      sr = sr & x1 & ":" & x & ","
      If Len(sr) > 255 Then
        sr = sr1
        x = x1 - 1
        Application.Intersect(Range("a:d"), Range(Left(sr, Len(sr) - 1))).Interior.ColorIndex = 3
        sr = ""
      End If
      x = x - 1
   End If
 Next x
 MsgBox Timer - t
End Sub


Option Explicit
''数组也可以设置格式？
   ''数组除了数字类型外，当然没有颜色、字体等格式，但是别忘了range对象可以表示多个连续或不连续的单元格区域
   ''利用上述特点，我们就是要数组构造单元格地址串，然后批量对单元格进行格式设置。
   ''注意，单元格地址串不能>255，所以如果单元格操作过多，我们还需要分次分批设置单元格格式
   
Sub 填充颜色()
 Range("a2:d2,a7:d7,a10:d10").Interior.ColorIndex = 3
End Sub




’第25集
Option Explicit

Sub 冒泡排序()
 Dim arr, temp, x, y, t, k
 t = Timer
 arr = Range("a1:a10")
 For x = 1 To UBound(arr) - 1
   For y = x + 1 To UBound(arr) ''只和当前数字下面的数进行比较
     If arr(x, 1) > arr(y, 1) Then ''如果它大于它下面某一个数字
       temp = arr(x, 1)
       arr(x, 1) = arr(y, 1)
       arr(y, 1) = temp
     End If
       
   Next y
 Next x
 Range("b3").Resize(x) = ""
 Range("b3").Resize(x) = arr
 ''Range("b2") = Timer - t
 MsgBox k
End Sub


Sub 冒泡排序演示()
 Dim arr, temp, x, y, t, k
 For x = 1 To 9
                         Range("a" & x).Interior.ColorIndex = 3
   For y = x + 1 To 10  ''只和当前数字下面的数进行比较
                         Range("a" & y).Interior.ColorIndex = 4
     If Cells(x, 1) > Cells(y, 1) Then ''如果它大于它下面某一个数字
       temp = Cells(x, 1)
       Cells(x, 1) = Cells(y, 1)
       Cells(y, 1) = temp
2     End If
                         Range("a" & y).Interior.ColorIndex = xlNone
   Next y
                         Range("a" & x).Interior.ColorIndex = xlNone
                         
 Next x

End Sub

Option Explicit


Sub 选择排序()
  Dim arr, temp, x, y, t, iMax, k, k1, k2
  t = Timer
  arr = Range("a1:a10")
  For x = UBound(arr) To 1 + 1 Step -1
     iMax = 1 ''最大的索引
     For y = 1 To x
          If arr(y, 1) > arr(iMax, 1) Then iMax = y
     Next y
     temp = arr(iMax, 1)
     arr(iMax, 1) = arr(x, 1)
     arr(x, 1) = temp
  Next x
  
  ''Range("c3").Resize(UBound(arr)) = ""
  ''Range("c3").Resize(UBound(arr)) = arr
  ''Range("c2") = Timer - t
  ''MsgBox k1
End Sub

Sub 选择排序单元格演示()
  Dim arr, temp, x, y, t, iMax, k, k1, k2

  For x = 10 To 2 Step -1
     iMax = 1
                       Range("a" & x).Interior.ColorIndex = 3
     For y = 1 To x
                       Range("a" & y).Interior.ColorIndex = 4
          If Cells(y, 1) > Cells(iMax, 1) Then
                       Range("a" & iMax).Interior.ColorIndex = xlNone
           iMax = y
          End If
                       Range("a" & y).Interior.ColorIndex = xlNone
                       Range("a" & iMax).Interior.ColorIndex = 6
                       
     Next y
     temp = Cells(iMax, 1)
     Cells(iMax, 1) = Cells(x, 1)
     Cells(x, 1) = temp
     Range("a" & x).Interior.ColorIndex = xlNone
     Range("a" & iMax).Interior.ColorIndex = xlNone
  Next x

End Sub
    
''第26集

Option Explicit


Sub 插入排序()
Dim arr, temp, x, y, t, iMax, k, k1, k2
  t = Timer
  arr = Range("a1:a10")
  For x = 2 To UBound(arr)
  
     temp = arr(x, 1) ''记得要插入的值
     
     For y = x - 1 To 1 Step -1
       If arr(y, 1) <= temp Then Exit For
       arr(y + 1, 1) = arr(y, 1)
       ''k1 = k1 + 1
     Next y
     arr(y + 1, 1) = temp
     ''k2 = k2 + 1
  Next
 '' Range("d3").Resize(UBound(arr)) = ""
 '' Range("d3").Resize(UBound(arr)) = arr
 ''Range("d2") = Timer - t
 MsgBox k1
End Sub

Sub 插入排序单元格演示()
On Error Resume Next
  Dim arr, temp, x, y, t, iMax, k
  For x = 2 To 10
  
     temp = Cells(x, 1) ''记得要插入的值
               Range("A" & x).Interior.ColorIndex = 3
     For y = x - 1 To 1 Step -1
               Range("A" & y).Interior.ColorIndex = 4
       If Cells(y, 1) <= temp Then Exit For
               Cells(y + 1, 1) = Cells(y, 1)
               Range("A" & y).Interior.ColorIndex = xlNone
     Next y
     Cells(y + 1, 1) = temp
               Range("A" & y).Interior.ColorIndex = xlNone
               Range("A" & x).Interior.ColorIndex = xlNone
  Next

End Sub


Option Explicit

Sub dd()
    Dim arr1(0 To 4999) As Long, arr, x, t
    t = Timer
    arr = Range("a1:a5000")
    For x = 1 To 5000
      arr1(x - 1) = arr(x, 1)
    Next x
    QuickSort arr1()
    Range("f2") = Timer - t
End Sub
Public Sub QuickSort(ByRef lngArray() As Long)

    Dim iLBound As Long

    Dim iUBound As Long

    Dim iTemp As Long

    Dim iOuter As Long

    Dim iMax As Long
   

    iLBound = LBound(lngArray)

    iUBound = UBound(lngArray)

    

    ''若只有一个值，不排序

    If (iUBound - iLBound) Then

        For iOuter = iLBound To iUBound

            If lngArray(iOuter) > lngArray(iMax) Then iMax = iOuter

        Next iOuter

        

        iTemp = lngArray(iMax)

        lngArray(iMax) = lngArray(iUBound)

        lngArray(iUBound) = iTemp

    

        ''开始快速排序

        InnerQuickSort lngArray, iLBound, iUBound

    End If
    Range("f3").Resize(5000) = Application.Transpose(lngArray)

End Sub

 

Private Sub InnerQuickSort(ByRef lngArray() As Long, ByVal iLeftEnd As Long, ByVal iRightEnd As Long)

    Dim iLeftCur As Long

    Dim iRightCur As Long

    Dim iPivot As Long

    Dim iTemp As Long

    

    If iLeftEnd >= iRightEnd Then Exit Sub

    

    iLeftCur = iLeftEnd

    iRightCur = iRightEnd + 1

    iPivot = lngArray(iLeftEnd)

    

    Do

        Do

            iLeftCur = iLeftCur + 1

        Loop While lngArray(iLeftCur) < iPivot

        

        Do

            iRightCur = iRightCur - 1

        Loop While lngArray(iRightCur) > iPivot

        

        If iLeftCur >= iRightCur Then Exit Do

        

        ''交换值

        iTemp = lngArray(iLeftCur)

        lngArray(iLeftCur) = lngArray(iRightCur)

        lngArray(iRightCur) = iTemp

    Loop

    

    ''递归快速排序

    lngArray(iLeftEnd) = lngArray(iRightCur)

    lngArray(iRightCur) = iPivot

    

    InnerQuickSort lngArray, iLeftEnd, iRightCur - 1

    InnerQuickSort lngArray, iRightCur + 1, iRightEnd

End Sub






Sub 希尔排序()
  Dim arr
  Dim 总大小, 间隔, x, y, temp, t
  t = Timer
  arr = Range("a1:a30")
  总大小 = UBound(arr) - LBound(arr) + 1
  间隔 = 1
  If 总大小 > 13 Then
     Do While 间隔 < 总大小
       间隔 = 间隔 * 3 + 1
     Loop
     间隔 = 间隔 \ 9
  End If
''  Stop
  Do While 间隔
     For x = LBound(arr) + 间隔 To UBound(arr)
      temp = arr(x, 1)
      For y = x - 间隔 To LBound(arr) Step -间隔
         If arr(y, 1) <= temp Then Exit For
         arr(y + 间隔, 1) = arr(y, 1)
        '' k1 = k1 + 1
      Next y
      arr(y + 间隔, 1) = temp
     Next x
    间隔 = 间隔 \ 3
   Loop
  '' MsgBox k1
   ''Range("e3").Resize(5000) = ""
    Range("d1").Resize(UBound(arr)) = arr
   ''Range("e2") = Timer - t
End Sub
Sub 打乱顺序()
 Dim arr, temp, x
 arr = Range("a1:a" & Range("a65536").End(xlUp).Row)
 For x = 1 To UBound(arr)
   num = Int(Rnd() * UBound(arr) + 1)
   temp = arr(num, 1)
   arr(num, 1) = arr(x, 1)
   arr(x, 1) = temp
 Next x
 Range("a1").Resize(x - 1) = arr
End Sub
Sub 希尔排序单元格演示()
  Dim arr
  Dim 总大小, 间隔, x, y, temp, t
  t = Timer
  arr = Range("a1:a" & Range("a65536").End(xlUp).Row)
  总大小 = UBound(arr) - LBound(arr) + 1
  间隔 = 1
  If 总大小 > 13 Then
     Do While 间隔 < 总大小
       间隔 = 间隔 * 3 + 1
     Loop
     间隔 = 间隔 \ 9
  End If
''  Stop
  Do While 间隔
     For x = LBound(arr) + 间隔 To UBound(arr)
      temp = Cells(x, 1)
      Range("a" & x).Interior.ColorIndex = 3
      For y = x - 间隔 To LBound(arr) Step -间隔
          Range("a" & y).Interior.ColorIndex = 6
         If Cells(y, 1) <= temp Then Exit For
         Cells(y + 间隔, 1) = Cells(y, 1)
        '' k1 = k1 + 1
      Next y
      Cells(y + 间隔, 1) = temp
      Range("a1:a30").Interior.ColorIndex = xlNone
     Next x
    间隔 = 间隔 \ 3
   Loop
  '' MsgBox k1
   ''Range("e3").Resize(5000) = ""
   '' Range("d1").Resize(UBound(arr)) = arr
   ''Range("e2") = Timer - t
End Sub


Option Explicit

Sub 工作表排序之冒泡法()
Dim arr(1 To 1000), i As Integer, j As Integer, k
Application.ScreenUpdating = False
For i = 1 To Sheets.Count
    arr(i) = Sheets(i).Name
Next i
For i = 1 To Sheets.Count - 1
    For j = i To Sheets.Count
        If arr(j) < arr(i) Then
           k = arr(i)
           arr(i) = arr(j)
           arr(j) = k
        End If
    Next j
Next i
For i = 1 To Sheets.Count
    Sheets(arr(i)).Move before:=Sheets(i)
Next i
Sheets("王6").Activate
Application.ScreenUpdating = True

End Sub
Sub 工作表排序之选择法()
Dim arr(1 To 1000), i As Integer, j As Integer, k, m As Integer
Application.ScreenUpdating = False
For i = 1 To Sheets.Count
    arr(i) = Sheets(i).Name
Next i
For i = 1 To Sheets.Count - 1
    m = i
    For j = i To Sheets.Count
        If arr(j) < arr(m) Then m = j
    Next j
    k = arr(i)
    arr(i) = arr(m)
    arr(m) = k
Next i
For i = 1 To Sheets.Count
    Sheets(arr(i)).Move before:=Sheets(i)
Next i
Sheets("王6").Activate
Application.ScreenUpdating = True
End Sub
Sub 工作表排序之插入法()
Dim arr(1 To 1000), i As Integer, j As Integer, k
Application.ScreenUpdating = False
For i = 1 To Sheets.Count
    arr(i) = Sheets(i).Name
Next i
For i = 2 To Sheets.Count
    k = arr(i)
    For j = i - 1 To 1 Step -1
        If arr(j) <= k Then Exit For
        arr(j + 1) = arr(j)
    Next j
    arr(j + 1) = k
Next i
For i = 1 To Sheets.Count
    Sheets(arr(i)).Move before:=Sheets(i)
Next i
Sheets("王6").Activate
Application.ScreenUpdating = True
End Sub

Sub 打乱顺序()
Dim arr(1 To 1000), i As Integer, j As Integer, k
Application.ScreenUpdating = False
For i = 1 To Sheets.Count
    arr(i) = Sheets(i).Name
Next i
For i = 1 To Sheets.Count
    j = Int(Rnd() * Sheets.Count + 1)
    k = arr(i)
    arr(i) = arr(j)
    arr(j) = k
Next i
For i = 1 To Sheets.Count
    Sheets(arr(i)).Move before:=Sheets(i)
Next i
Sheets("王6").Activate
Application.ScreenUpdating = True
End Sub

Sub 希尔排序()

Dim arr(1 To 1000), i As Integer, j As Integer, k, m As Integer, n As Integer
Application.ScreenUpdating = False
For i = 1 To Sheets.Count
    arr(i) = Sheets(i).Name
Next i
i = Sheets.Count
j = 1
If i > 13 Then
   Do While j < i
      j = j * 3 + 1
   Loop
   j = j \ 9
End If
Do While j
   For m = j + 1 To Sheets.Count
       k = arr(m)
       For n = m - j To 1 Step -j
           If arr(n) <= k Then Exit For
           arr(n + j) = arr(n)
       Next n
       arr(n + j) = k
   Next m
   j = j \ 3
Loop
For m = 1 To Sheets.Count
    Sheets(arr(m)).Move before:=Sheets(m)
Next m
Sheets("王6").Activate
Application.ScreenUpdating = True

End Sub


Option Explicit

Sub 工作表排序之希尔排序()
Dim arr(1 To 1000), i As Integer, j As Integer, k, m As Integer, n As Integer
Application.ScreenUpdating = False
For i = 1 To Sheets.Count
    arr(i) = Sheets(i).Name
Next i
i = Sheets.Count
j = 1
If i > 13 Then
   Do While j < i
      j = j * 3 + 1
   Loop
   j = j \ 9
End If
Do While j
   For m = j + 1 To Sheets.Count
       k = arr(m)
       For n = m - j To 1 Step -j
           If arr(n) <= k Then Exit For
           arr(n + j) = arr(n)
       Next n
       arr(n + j) = k
   Next m
   j = j \ 3
Loop
For m = 1 To Sheets.Count
    Sheets(arr(m)).Move before:=Sheets(m)
Next m
Sheets("王6").Activate
Application.ScreenUpdating = True

End Sub

''第27集

Option Explicit

''1 什么是VBA字典？
   ''字典（dictionary）是一个储存数据的小仓库。共有两列。
      ''第一列叫key , 不允许有重复的元素。
      ''第二列是item,每一个key对应一个item，本列允许为重复
            ''Key   item
             ''A     10
             ''B     20
             ''C     30
             ''Z     10

''2 即然有数组，为什么还要学字典？
   ''原因:提速，具体表现在
      ''1) A列只能装入非重复的元素，利用这个特点可以很方便的提取不重复的值
      ''2) 每一个key对应一个唯一的item，只要指点key的值，就可以马上返回其对应的item，利用字典可以实现快速的查找

''3 字典有什么局限？
    ''字典只有两列，如果要处理多列的数据，还需要通过字符串的组合和拆分来实现。
    ''字典调用会耗费一定时间，如果是数据量不大，字典的优势就无法体现出来。
    
''4 字典在哪里？如何创建字典？
    
    ''字典是由scrrun.dll链接库提供的，要调用字典有两种方法
      ''第一种方法：直接创建法
        '''set d = CreateObject("scripting.dictionary")
      ''第二种方法：引用法
        ''工具-引用-浏览-找到scrrun.dll-确定

		
Option Explicit
 
 ''1 装入数据
    Sub t1()
      Dim d As New Dictionary
      Dim x As Integer
      For x = 2 To 4
       d.Add Cells(x, 1).Value, Cells(x, 2).Value
      Next x
      MsgBox d.Keys(1)
      '''stop
    End Sub
 ''2 读取数据
    Sub t2()
      Dim d
      Dim arr
      Dim x As Integer
      Set d = CreateObject("scripting.dictionary")
      For x = 2 To 4
       d.Add Cells(x, 1).Value, Cells(x, 2).Value
      Next x
      ''MsgBox d("李四")
      ''MsgBox d.Keys(2)
      Range("d1").Resize(d.Count) = Application.Transpose(d.Keys)
      Range("e1").Resize(d.Count) = Application.Transpose(d.Items)
      arr = d.Items
    End Sub

  ''3 修改数据
    Sub t3()
      Dim d As New Dictionary
      Dim x As Integer
      For x = 2 To 4
       d.Add Cells(x, 1).Value, Cells(x, 2).Value
      Next x
      d("李四") = 78
      MsgBox d("李四")
      d("赵六") = 100
      MsgBox d("赵六")
    End Sub

  ''4 删除数据
    Sub t4()
      Dim d As New Dictionary
      Dim x As Integer
      For x = 2 To 4
        d(Cells(x, 1).Value) = Cells(x, 2).Value
      Next x
       d.Remove "李四"
     '' MsgBox d.Exists("李四")
      d.RemoveAll
      MsgBox d.Count
    End Sub
 
''区分大小写
    Sub t5()
      Dim d As New Dictionary
      Dim x
      For x = 1 To 5
        d(Cells(x, 1).Value) = ""
      Next x
      Stop
    End Sub

''第28集


	Option Explicit

Sub 多表双向查找()
 Dim d As New Dictionary
 Dim x, y
 Dim arr
 For x = 3 To 5
   arr = Sheets(x).Range("a2").Resize(Sheets(x).Range("a65536").End(xlUp).Row - 1, 2)
   For y = 1 To UBound(arr)
     d(arr(y, 1)) = arr(y, 2)
     d(arr(y, 2)) = arr(y, 1)
   Next y
 Next x
 MsgBox d("C1")
 MsgBox d("吴情")
End Sub


Option Explicit

Sub 汇总()
 Dim d As New Dictionary
 Dim arr, x
 arr = Range("a2:b10")
 For x = 1 To UBound(arr)
   d(arr(x, 1)) = d(arr(x, 1)) + arr(x, 2) ''key对应的item的值在原来的基础上加新的
 Next x
 Range("d2").Resize(d.Count) = Application.Transpose(d.Keys)
 Range("e2").Resize(d.Count) = Application.Transpose(d.Items)
End Sub

Option Explicit

Sub 提取不重复的产品()
 Dim d As New Dictionary
 Dim arr, x
 arr = Range("a2:a12")
 For x = 1 To UBound(arr)
      d(arr(x, 1)) = ""
 Next x
 Range("c2").Resize(d.Count) = Application.Transpose(d.Keys)
End Sub

''第29集

Option Explicit

Sub 下棋法之多列汇总()
 Dim 棋盘(1 To 10000, 1 To 3)
 Dim 行数
 Dim arr, x, k
 Dim d As New Dictionary
 arr = Range("a2:c" & Range("a65536").End(xlUp).Row)
 For x = 1 To UBound(arr)
   If d.Exists(arr(x, 1)) Then
      行数 = d(arr(x, 1))
      棋盘(行数, 2) = 棋盘(行数, 2) + arr(x, 2)
      棋盘(行数, 3) = 棋盘(行数, 3) + arr(x, 3)
   Else
      k = k + 1
      d(arr(x, 1)) = k
      棋盘(k, 1) = arr(x, 1)
      棋盘(k, 2) = arr(x, 2)
      棋盘(k, 3) = arr(x, 3)
   End If
 Next x
 Range("f2").Resize(k, 3) = 棋盘
End Sub


Option Explicit

Sub 下棋法之多条件多列汇总()
 Dim 棋盘(1 To 10000, 1 To 4)
 Dim 行数
 Dim arr, x As Integer, sr As String, k As Integer
 Dim d As New Dictionary
 arr = Range("a2:d" & Range("a65536").End(xlUp).Row)
 For x = 1 To UBound(arr)
    sr = arr(x, 1) & "-" & arr(x, 2)
    If d.Exists(sr) Then
      行数 = d(sr)
      棋盘(行数, 3) = 棋盘(行数, 3) + arr(x, 3)
      棋盘(行数, 4) = 棋盘(行数, 4) + arr(x, 4)
    Else
      k = k + 1
      d(sr) = k
      棋盘(k, 1) = arr(x, 1)
      棋盘(k, 2) = arr(x, 2)
      棋盘(k, 3) = arr(x, 3)
      棋盘(k, 4) = arr(x, 4)
    End If
 Next x
   Range("g2").Resize(k, 4) = 棋盘
End Sub


Option Explicit

Sub 下棋法之数据透视表式汇总()
 Dim d As New Dictionary
 Dim 棋盘(1 To 10000, 1 To 7)
 Dim 行数, 列数
 Dim arr, x, k
 
 arr = Range("a2:c" & Range("a65536").End(xlUp).Row)
 
 For x = 1 To UBound(arr)
   列数 = (InStr("1月2月3月4月5月6月", arr(x, 2)) + 1) / 2 + 1
   If d.Exists(arr(x, 1)) Then
      行数 = d(arr(x, 1))
      
      棋盘(行数, 列数) = 棋盘(行数, 列数) + arr(x, 3)
   Else
      k = k + 1
      d(arr(x, 1)) = k
      棋盘(k, 1) = arr(x, 1)
      棋盘(k, 列数) = arr(x, 3)
   End If
 Next x
 
 Range("f2").Resize(k, 7) = 棋盘

End Sub


''第30集

Option Explicit

''1 取得工作表总个数的自定义函数

Function shcount()

 shcount = Sheets.Count
 
End Function
Sub dd()
 MsgBox getv(Range("a7"))
End Sub


''2 取得单元格显示值的自定义函数
 
  Function getv(rg As Range)
  
    getv = rg.Text
    
  End Function
 
''3 截取字符串的函数
 
 Function jiequ(sr As String, fh As String, wz As Integer)
    
    Dim Arr
    Arr = Split(sr, fh)
    jiequ = Arr(wz - 1)
    
 End Function
  
''4 提取不重复值的个数

  Function 不重复个数(rg As Range)
   Dim d, Arr, ar
   Arr = rg
   Set d = CreateObject("scripting.dictionary")
   For Each ar In Arr
     d(ar) = ""
   Next ar
   不重复个数 = d.Count
  End Function
 Sub test()
  
  MsgBox jiequ("A-BRT-C-EF", "-", 2)
  
 End Sub

 
 Option Explicit

''1 什么是自定义函数？
  ''在VBA中有VBA函数，我们还可以调用工作表函数，我们能不能自已编写函数呢？可以，这就是本集所讲的自定义函数
  
''2 怎么编写自定义函数？
 
   ''我们可以按下面的结构编写自定义函数
  
    '' Function 函数名称(参数1,参数2....)
         
         ''代码
         ''函数名称=返回的值或数组
         
    '' End Function
    
    
Option Explicit

''1 怎么让自定义函数在所有工作簿中使用？
  
   ''答： 把含有自定义函数的文件另存为加截宏，然后通过工具-加截宏-浏览找到这个文件-确定。
   
''2 怎么给自定义函数添加说明

    ''工具-宏-宏名输入自定义函数的名称-选项--在说明栏中写入这个函数的名称
    
''3、怎么给自定义函数分类

    Sub 分类()
     Application.MacroOptions "不重复个数", Category:=4
    End Sub
     
   ''注:
         ''0 是全部
         ''1 财务
         ''2 日期和时间
         ''3 数学和三角
         ''4 统计
         ''5 查找和引用
         ''6 数据库
         ''7 文本
         ''8 逻辑
         ''9 信息
     



''第32集

Option Explicit

''一、什么MsgBox函数
   ''它可以弹出一个窗口，显示你设定的内容。并且窗口上有可以让你选择的按钮，点击不同的按钮会返回不同的数值。
 ''用msgbox信息窗口可以增加一个程序对话的机会，以告诉程序下一步应该怎么做
  
    Sub test1()
      MsgBox "大家好，我是msgbox窗口"
    End Sub

''二、基本语法
   
   ''Msgbox (窗口中显示的内容,按钮和图示类别,窗口标题,相关的帮助文件,帮助文件上下文的编号)
   
  
 Option Explicit

''按钮类型
   ''消息窗体由按钮显示,图标显示,缺省按钮和其他特殊功能组合,这些功能都可以随意组合,组合他们只需要用"+"号
   
  Sub test8()
    MsgBox "test", vbYesNoCancel + vbExclamation + vbDefaultButton2 + vbMsgBoxHelpButton ''显示确定和取消按钮并显示询问图标
  End Sub
  Sub test9()
    MsgBox "mytest", vbExclamation + vbYesNo ''显示危险图标和是否按钮
  End Sub
  Sub test10()
    MsgBox "测试窗体结构", vbYesNoCancel + vbMsgBoxHelpButton + vbCritical + vbDefaultButton3, "测试四个按钮的窗口"
  End Sub
 Sub dd()
   MsgBox "dd", vbYesNo + vbExclamation + vbMsgBoxHelpButton
 End Sub


 Option Explicit

''1、窗口显示的内容
    
     ''1) 基本显示:只需要给第一个参数设置一个字符串或生成字符串的表达式即或
        
        ''例:
        Sub test2()
          MsgBox "你好,欢迎你的使用"
          MsgBox "你好!,欢迎你使用" & ThisWorkbook.Name
        End Sub
      
      ''2) 换行显示。
            ''chr(10) 可以生成换行符
            ''chr(13) 可以生成回车符
            ''vbcrlf 换行符和回车符
            ''vbCr 等同于chr(10)
            ''vblf 等同于chr(13)
         ''例：
         Sub test3()
           MsgBox "我爱" & Chr(10) & "Excel精英培训"
          '' MsgBox "我爱你" & Chr(13) & "Excel"
          '' MsgBox "今天" & vbCrLf & "我是水王"

         End Sub
     
        ''3) 表格显示
          ''chr(9) 制表符
          Sub test4()
             MsgBox "姓名" & Chr(9) & "职业" & Chr(10) & "张三" & Chr(9) & "工程师" _
                     & Chr(10) & "于上伟" & Chr(9) & "教师"
          End Sub
     
          Sub test5()
            Dim sr, x, y
            For x = 1 To 5
              For y = 1 To 3
                sr = sr & Cells(x, y) & Chr(9) & Chr(9)
              Next y
              sr = sr & Chr(13)
            Next x
            MsgBox sr
          End Sub
         
          ''用空格键设置
            '' space(n) 可以产生N个空格
          Sub test6()
          Dim x, y, sr, k
             For x = 1 To 5
               For y = 1 To 3
                 If VBA.IsNumeric(Cells(x, y)) Then
                   k = 12 - Len(Cells(x, y))
                 Else
                  k = 12 - Len(Cells(x, y)) * 2
                 End If
                  sr = sr & Cells(x, y) & Space(k)
               Next y
               sr = sr & Chr(13)
             Next x
           MsgBox sr
          End Sub
 
 ''2  标题的显示文字
    Sub test7()
      MsgBox "核对关系出错了", , "系统提示"
    End Sub
   

 Option Explicit

''要想和消息框交流,还需要在我们点击窗体的按钮后能返回一个值,告诉程序我们点了哪个按钮.

 Sub test11()
  Dim k
  k = MsgBox("测试返回值", vbYesNoCancel)
  MsgBox "你点击了按钮:" & Choose(k, "确定", "取消", "终止", "重试", "忽略", "是", "否")
 End Sub

 ''应用示例
   Sub test12()
     If MsgBox("你确定要删除第15行吗?", vbQuestion + vbYesNo, "删除提示") = vbYes Then
       Rows(15).Delete
       MsgBox "删除成功"
     Else
       MsgBox "你取消了删除"
     End If
   End Sub


Option Explicit
 
  ''要添加帮助,需要设置msgbox 函数的第四和第五个参数
    ''第四个参数是帮助文件的路径,帮助文件要放在C:\WINDOWS\Help路径下
    ''第五个参数和帮助文件本身有关,是为了准备的打开帮助文件而设置的上下文编号,如果没有则设置为0
  Sub test13()
  Dim x
  x = MsgBox("测试添加帮助的效果", vbOKCancel + vbMsgBoxHelpButton, "测试帮助!", "D:/a.chm", 0) ''"C:\WINDOWS\Help\excel.chm", 0)
  End Sub



Option Explicit

''1 自动定时关闭消息框,可以用其他消息框完成

 Sub AA()
    Dim WshShell As Object
    Set WshShell = CreateObject("Wscript.Shell")
    WshShell.Popup "1秒后关闭！", 1, "提示！", 16
End Sub
 
 
''第33集


Option Explicit

''1.inpubox函数
  
  ''语法:
    ''inputbox(输入框显示内容,窗体标题,默认值,水平位置,垂直位置,帮助文件,帮助文件ID
    
''2.Application对象的Inputbox方法:显示一个接收用户输入的对话框。返回此对话框中输入的信息
  
   ''语法:
     ''Application.InputBox(对话框显示内容,输入框标题,文本框内默认值,x坐标,y坐标,帮助文件,帮助文件上下文ID,文本框内输入类型)
  
  ''最后一个参数数值说明:
 ''     值    含义
      ''0     公式
      ''1     数字
      ''2     文本 (字符串)
      ''4     逻辑值 (True 或 False)
      ''8     单元格引用，作为一个 Range 对象
      ''16    错误值，如 #N/A
      ''64    数值数组

 ''什么时候用方法,什么时候用函数
       ''从上面的参数可以看出inputbox函数和方法的不同之处是方法比函数多了后面几个参不数,如果只是简单的输入,可以用方法,
    ''如果需要添加帮助和设置输入类型,则用Application对象的Inputbox方法.

	
	
Option Explicit
  ''最后一个参数数值说明:
 ''     值    含义
      ''0     公式
      ''1     数字
      ''2     文本 (字符串)
      ''4     逻辑值 (True 或 False)
      ''8     单元格引用，作为一个 Range 对象
      ''16    错误值，如 #N/A
      ''64    数值数组
      
'' 1.引用单元格
     ''inputbox方法的最后个参数值为8的时候,可以用鼠标选择单元格的地址.使用变量是使用SET声明的对象变量,则返回的是一个单元格对象,
  ''否则反回的这个单元格区域的值,即VBA数组.
   Sub text5()
     Dim rg As Range
     Set rg = Application.InputBox("请选择单元格区域", "选取提示", , , , , , 8)
     MsgBox rg.Parent.Name & "!" & rg.Address
   End Sub
  
     Sub text6()
     Dim rg
      rg = Application.InputBox("请选择单元格区域", "选取提示", , , , , , 8)
     MsgBox rg(2, 1)
   End Sub

 ''2 公式引用
    ''当最后一个参数设置为0时,可以输入公式,返回的也是一个公式字符串,如果公式中含单元格引用,可以自动转换成rc引用格式(以当前活动单元格为参照)
    
    Sub test7()
      Dim r
      r = Application.InputBox("请输入公式", "输入提示", , , , , , 0)
      MsgBox r
    End Sub

 ''3 限制输入返回的数值格式
  Sub test8()
      Dim r
      r = Application.InputBox("请输入公式", "输入提示", , , , , , 1) ''输入非数字则会提示无效的数字
      MsgBox r
  End Sub
  Sub test9()
      Dim r
      r = Application.InputBox("请输入公式", "输入提示", , , , , , 2) ''可以输入字符,当然,文字型数字也符字符
      MsgBox TypeName(r)
  End Sub
 ''4.数值数组
    ''可以选取单元格区域的值作为数组,也可以输入以带有大括号的一维或二维数组
  Sub test10()
      Dim r
      r = Application.InputBox("请输入公式", "输入提示", , , , , , 64) ''可以输入字符,当然,文字型数字也符字符
      MsgBox r(2, 1)
  End Sub
 

 Option Explicit
''1 输入的内容返回给一个变量
Sub test1()
  Dim sr
  sr = InputBox("输入测试", "测试", 100)
    MsgBox sr
  sr = Application.InputBox("输入测试", "测试", 100)
    MsgBox sr
End Sub

''2 如果不输入直接点确定返回什么
 
Sub test2()
  Dim sr
  sr = InputBox("输入测试", "测试")
    MsgBox sr
  sr = Application.InputBox("输入测试", "测试")
    MsgBox sr
End Sub

        ''经过测试发现当不输入任何内容直接点确定都会返回空,所以我们就可以用空来判断是否输入了内容

 Sub test3()
  Dim sr
  sr = InputBox("输入测试", "测试")
  If sr = "" Then
    MsgBox "你没有输入就点了确定"
  End If
    
  sr = Application.InputBox("输入测试", "测试")
      If sr = "" Then
    MsgBox "你没有输入就点了确定"
      ElseIf sr = "False" Then
      
  End If
End Sub

''3 如果直接点了"退出"按钮会有什么值返回
  
Sub test4()
  Dim sr
  sr = InputBox("输入测试", "测试")
    MsgBox sr ''返回空
  sr = Application.InputBox("输入测试", "测试")
    MsgBox sr ''返回False
End Sub

         ''由上面2,3可以看出,如果需要判断是否输入了内容和是否点击了退出,用Inpubox函数时判断返回值是否为空就可以了,
   '' 如果是Inputbox方法,则需要进行两种判断.

   
   
 ''第34集 
 Option Explicit

''一 FileDialog 对象简介
 ''提供文件对话框，功能与 Microsoft Office 应用程序中标准的“打开”和“保存”对话框类似。
 ''利用这些对话框，解决方案的用户可以简便地指定解决方案中应该使用的文件和文件夹。

''
''“打开”对话框：让用户选择一个或多个可以在主机应用程序中使用 Execute 方法打开的文件。
''“另存为”对话框：让用户选择一个可以使用 Execute 方法保存当前文件的文件。
''“文件选取器”对话框：让用户选择一个或多个文件。用户选择的文件路径将捕获到 FileDialogSelectedItems 集合。
''“文件夹选取器”对话框：让用户选择一个路径。用户选择的文件路径将捕获到 FileDialogSelectedItems 集合。

''二 属性和方法
  
   ''1 AllowMultiSelect 如果允许用户从文件对话框中选择多个文件，则返回 True。Boolean 类型，可读写
   ''2 SelectedItems 选取的多个文件集合
   ''3 InitialFileName 属性:设置初始路径和文件名称
   ''4 InitialView 属性 :可以设置初始文件的显示样多
   ''5 show 可以判断用户是否点击了取消按钮,如果点击取消会返回0,否则返回-1
   
    ''选择并返回一组文件名和路径
      Sub f1()
        Dim f
        Dim dig As Object
        Set dig = Application.FileDialog(msoFileDialogOpen)
        With Application.FileDialog(msoFileDialogOpen)
           .AllowMultiSelect = True
           .Filters.Add "Excel文件", "*.xls", 1
           .InitialFileName = ThisWorkbook.FullName ''"d:\"
           .InitialView = msoFileDialogViewDetails
           .Title = "对话框测试"
           .Show
           MsgBox .Show
          For Each f In .SelectedItems
            MsgBox f
          Next f
        End With
        Set dig = Nothing
      End Sub
   ''选择并返回文件夹
     Sub F2()
      Dim dig As Object
      Set dig = Application.FileDialog(msoFileDialogFolderPicker)
       With dig
         .InitialFileName = "d:\"
         .Show
         MsgBox .SelectedItems(1)
       End With
     Set dig = Nothing
     End Sub
   ''

Sub t10()
Dim f
 With Application.FileDialog(msoFileDialogOpen)
     .AllowMultiSelect = True
     .Filters = "Excel表格,*.xls"
     .InitialFileName = "测试.xls"
     .FilterIndex = 1
     .Title = "测试"
  End With
End Sub


Option Explicit

'' 一、 概述基本语法

   ''GetOpenFilename相当于Excel打开窗口，通过该窗口选择要打开的文件，并可以返回选择的文件完整路径和文件名。
     ''注:此方法并不会真正打开文件?

  ''Application.GetOpenFilename(文件类型筛选规则,优先显示第几个类型的文件,标题,是否允许选择多个文件名)
  
     
  
 ''二、示例
   
      ''1 打开类型只限excel文件
      
        ''设置打开某类文件可以用下面的规则：
          
           ''"文件类型说明文字,*.文件类型后辍"
      Sub t1()
        Dim f
        f = Application.GetOpenFilename("Excel文件,*.xls")
        MsgBox f
      End Sub
         
      ''2、打开多种文件类型(word和excel)
        
       ''打开多种文件类型，只需要用","隔开，添加新的文件类型说明和文件类型。
       
      Sub t2()
        Dim f
        f = Application.GetOpenFilename("Excel2003文件,*.xls,Word文件,*.doc")
        MsgBox f
      End Sub
  
      ''3 打开多种文件类型,默认显示word文件
      
      Sub t3()
        Dim f
        f = Application.GetOpenFilename("Excel2003文件,*.xls,Word文件,*.doc,文本文件,*.txt", 2)
        MsgBox f
      End Sub
        
       ''4 设置对话框名称
       
       Sub t4()
          Dim f
           f = Application.GetOpenFilename("Excel2003文件,*.xls,Word文件,*.doc,文本文件,*.txt", 2, "选择要汇总的文件")
           MsgBox f
      End Sub
   
       ''5 选择多个文件,并以数组形式返回
      Sub t5()
        Dim f
        ChDrive "E"
        ChDir Application.Path
        ''ChDir ".."
        f = Application.GetOpenFilename("Excel2003文件,*.xls,Word文件,*.doc,文本文件,*.txt", 1, MultiSelect:=True)
        MsgBox f(1)
      End Sub
 



     ''GetSaveAsFilename语法:
         
         '' GetSaveAsFilename(默认显示的文件名,筛选条件,多个筛选类型时显示第几个,标题)
         ''注:该窗口也会有实质性的保存操作.只作为返回文件名的一个途径

       Sub t1()
        Dim f
        f = Application.GetSaveAsFilename("示例.xls", "excel表格,*.xls", , "保存示例")
        MsgBox f
      End Sub

	  Option Explicit
  
   ''chdrive 盘符 可以改变默认驱动器
   ''chdir  路径  可以改变默认路径
   
      Sub t6()
        Dim f
         ChDrive "E"
         ChDir ThisWorkbook.Path
        ''ChDir ".."
        f = Application.GetOpenFilename("Excel2003文件,*.xls,Word文件,*.doc,文本文件,*.txt", 1, MultiSelect:=True)
       '' MsgBox f(1)
      End Sub

	  
''第35集

Option Explicit
''字符串截取

''left,right,mid,Len
Sub z1()
  Dim sr
  sr = "Excel精英培训网"
  Debug.Print Left(sr, 5)
  Debug.Print Right(sr, 5)
  Debug.Print Mid(sr, 3, 5)
  Debug.Print Left(sr, Len(sr) - 1)
End Sub

'''split
 
Sub z2()
  Dim sr, arr
  sr = "Excel的精的英的培训网"
  arr = Split(sr, "的")
  Debug.Print UBound(arr)
End Sub


''val

 Sub z3()
  Dim sr
  sr = "89.90美元"
  Debug.Print Val(sr)
 End Sub

''字符串组合
 ''&
 Sub a4()
  Debug.Print "a" & "b"
 End Sub
 ''join
  
 Sub a5()
  Dim sr, arr
  sr = "Excel-精英-培训网"
  arr = Split(sr, "-")
  Debug.Print Join(arr, "+")
End Sub


Option Explicit

''instr 从前向后查

Sub c1()
  Dim sr
  sr = "Excel精英培训"
  Debug.Print InStr(sr, "精英") > 0
End Sub

''InStrRev 从后向前

Sub c2()
  Dim sr
  sr = "Excel精英培训培训论坛"
  Debug.Print InStr(sr, "培")
End Sub
''Replace替换

Sub c5()
 Dim sr
  sr = "Excel精英培训网"
  sr = Replace(sr, "培训网", "论坛")
  Debug.Print sr
End Sub

''mid语句替换

Sub c6()
 Dim sr
  sr = "Excel精英培训网"
  Mid(sr, 8, 2) = "论坛"
  Debug.Print sr
End Sub


Option Explicit

''LCase 转换成小写

Sub z1()
  Debug.Print LCase("ABC")
End Sub

''UCcae 转换成大写

Sub z2()

  Debug.Print UCase("Abc")
  
End Sub

'''strConv 函数

''常数 值 说明
''vbUpperCase 1 将字符串文字转成大写。
''vbLowerCase 2 将字符串文字转成小写。
''vbProperCase 3 将字符串中每个字的开头字母转成大写
Sub 转换()

  Debug.Print VBA.StrConv("wHo ARE you?", vbProperCase)
  
End Sub

Sub 转换2()
 Dim i As Long
Dim x() As Byte
x = StrConv("ABCDEFG", vbFromUnicode)    '' 转换字符串。
Debug.Print Application.Min(x)
For i = 0 To UBound(x)
    Debug.Print x(i)
Next

End Sub


''TRim删除两端空格
''Ltrim 删除左边空格
''Rtrim 删除右边空格
 Sub z3()
 Dim sr
 
 sr = " A B BC "
 Debug.Print Trim(sr)
 Debug.Print LTrim(sr)
 Debug.Print RTrim(sr)
 End Sub
 
''ASC 返回一个 Integer，代表字符串中首字母的字符代码,ANSI 字符集
''CHr 返回 String，其中包含有与指定的字符代码相关的字符
Sub z4()
  Debug.Print Asc("Z")
  Debug.Print Chr(90)
End Sub

'''space 和 string生成重复的字符

 Sub z5()
 
    Debug.Print "A" & Space(10) & "B"
    Debug.Print "C" & String(10, "a") & "D"
    
 End Sub

''第36集

Option Explicit

''like "对比的字符串"
''Option Compare Text
  '' 字符串1 like 字符串2
 Sub L1()
   Debug.Print "ABC" Like "ABc"
 End Sub

''通配符?
  ''判断BA是不是长度为2，且第二个字符为A
 Sub L2()
   Debug.Print "BA" Like "?A"
 End Sub

''通配符*
     ''判断字符串中是否包括cel
 Sub L3()
   Debug.Print "Excel精英培训" Like "*cel*"
 End Sub

''判断含通配符的字符串

  ''把通配符放在[]内，就代表本身字符的对比

 Sub l4()
   ''Debug.Print "QAB" Like "Q?B"
   Debug.Print "QaB" Like "Q?B"
   Debug.Print "Q?B" Like "Q[?]B"
   ''Debug.Print ""
 End Sub
 

''判断是指定位数数字
  ''判断数字是否为两个整数构成的
 Sub l9()
    Debug.Print 5 Like "##"
 End Sub

''判断在某个区间的字符
 
  Sub L10()
   ''[最小-最大最小2-最小3]
    ''Debug.Print "q" Like "[A-Za-z]"  '' 判断q是不是字母
   '' Debug.Print "H" Like "[A-GM-Z]"  '' 判断H是不是在A-G，M-Z区间
    Debug.Print 8 Like "[!2-9]"
  End Sub

''判断非在某个区间的字符
   Sub L11()
   
     Debug.Print "A" Like "[!C-Z]"
     
   End Sub
   
''判断在列出的字符里

   Sub L12()
   
      Debug.Print "M" Like "[!ABCDEUE]"
      
   End Sub
    
''判断A~C开头，F~G结尾
  
   Sub L13()
     
     Debug.Print "AEREM" Like "[A-C]*[L-P]"
     Debug.Print "A334M" Like "[A-C]###[L-P]"
     
   End Sub
 

 Option Explicit

Sub 求和()
Dim x, y, k, z
For x = 2 To 11
  
  For y = 2 To 12
    If Cells(y, 1) Like Cells(x, "F") Then
    z = Cells(x, "F")
       k = k + Cells(y, 2)
       Range("a" & y).Interior.ColorIndex = 3
    End If
  Next y
  Cells(x, "g") = k
  Cells(x, "f").Interior.ColorIndex = 3
  k = 0
  Stop
  Cells(x, "f").Interior.ColorIndex = xlNone
  Range("a2:a12").Interior.ColorIndex = xlNone
Next x
End Sub


''第37集

Option Explicit

''一 正则表达式

   ''正则表达式是处理字符串的外部工具，它可以根据设置的字符串对比规则，进行字符串的对比、替换等操作。
   
   ''正则表达式的作用：
     ''1、完成复杂的字符串判断
     ''2、在字符串判断时，可以最大限度的避开循环，从而达到提高运行效率的目的。
   
''二 使用方法
   
   ''1、引用法
   ''点击VBE编辑器菜单：工具 - 引用，选取: Microsoft VBScript Regular Expressions 5.5,引用后在程序开始进行如下声明
     ''Dim regex As New RegExp
     Sub t1()
       Dim reg As New RegExp
     End Sub
     
    ''2、直接他建法
''     代码引用 (后期绑定)
''     Dim regex As Object
''     Set regex = CreateObject("VBScript.RegExp") ''创建正则对象

     Sub t2()
       Dim reg As Object
       Set reg = CreateObject("VBScript.RegExp")
     End Sub

 ''三 常用属性
    
    ''1 Global属性:
       ''如果值为true,则搜索全部字符
       ''如果值为False,则搜索到第1个即停止
       ''1 例:
       Sub t3()
         Dim reg As New RegExp
         Dim sr
         sr = "ABCEA"
         With reg
           .Global = True
           .Pattern = "A"
           Debug.Print .Replace(sr, "")
         End With
       End Sub
       
    ''2 IgnoreCase 属性
       ''如果搜索是区分大小写的，为False（缺省值）True不分
    
    ''3 Pattern 属性
       '' 一个字符串，用来定义正则表达式。缺省值为空文本。
    ''4 Multiline 属性,字符串是不是使用了多行,如果是多行,$适用于每一行的最后一个
       
       Sub t4()
         Dim reg As New RegExp
         Dim sr
         sr = "AEA" & Chr(10) & "ABCA"
         With reg
           .Global = True
           .MultiLine = True
           .Pattern = "A$"
           ''.Pattern = "^A"
           Debug.Print .Replace(sr, "")
         End With
       End Sub
       
     ''5  Execute 方法
         ''返回一个 MatchCollection 对象，该对象包含每个成功匹配的 Match 对象,
         ''返回的信息包括:
           ''FirstIndex:开始位置
           ''Length; 长度
           ''Value:长度
       Sub t5()
         Dim reg As New RegExp
         Dim sr, matc
         sr = "A454BCEA5"
         With reg
           .Global = True
           .Pattern = "A\d+"
           Set matc = .Execute(sr)
         End With
         Stop
       End Sub
     
       Function ns(rg)
         Dim reg As New RegExp
         Dim sr, ma, s, m, x
         With reg
           .Global = True
           .Pattern = "\d*\.?\d*"
         Set ma = .Execute(rg)
           For Each m In ma
              s = s + Val(m)
           Next m
         End With
        ns = s
       '' Stop
       End Function
       
     ''6、Text方法
        ''返回一个布尔值，该值指示正则表达式是否与字符串成功匹配。其实就是判断两个字符串是否匹配成功
        Sub t7()
         Dim reg As New RegExp
         Dim sr
         sr = "BCR6EA"
         With reg
           .Global = True
           .Pattern = "\d+"
           If .test(sr) Then MsgBox "字符串中含有数字"
         End With
        End Sub
        
 

 
 Option Explicit

Function 提取中文(rg As String, k As Integer)

  Dim regx As New RegExp
  With regx
   .Global = True
   If k = 1 Then
   
    .Pattern = "\D"
    
   ElseIf k = 2 Then
   
    .Pattern = "\w"
    
   End If
   
   提取中文 = .Replace(rg, "")
  End With

End Function


''第38集

Option Explicit
 ''正则表达式的核心是设置对比的规则，也就是设置Pattern属性，而组成这些规则除了字符本身以外，是具有特定含义的符号。
 ''下面介绍的是正规表达式中常用符号的第一部分。
 
''\号

  ''1.放在不便书写的字符前面,如换行符(\r),回车符(\n),制表符(\t),\自身(\\)
  
  ''2.放在有特殊意义字符的前面,表示它自身,"\$","\^","\."
  
  ''3.放在可以匹配多个字符的前面
      
       ''\d 0~9的数字
       ''\w 任意一个字母或数字或下划线，也就是 A~Z,a~z,0~9,_ 中任意一个
       ''\s 包括空格、制表符、换页符等空白字符的其中任意一个
       
       ''以上改为大写时,为相反的意思,如\D 表示非数字类型
       
        Sub t1()
           Dim regx As New RegExp
           Dim sr
           sr = "AE45B646C"
           With regx
             .Global = True
             .Pattern = "\d" ''排除非数字
             Debug.Print .Replace(sr, "")
           End With
        End Sub
''.(点)

   ''可以匹配除换行符以外的所有字符

''+号
   ''+表示一个字符可以有任意多个重复的。
    
   Sub t11()
     Dim regx As New RegExp
     Dim sr
     sr = "A234CA7A"
     With regx
      .Global = True
      .Pattern = "A\d+"
      Debug.Print .Replace(sr, "")
     End With
     
   End Sub
''{}号
  ''可以设置重复次数
    ''1 {n} 重复n次
        Sub t16()
           Dim regx As New RegExp
           Dim sr
           sr = "A234CA7A67"
           With regx
            .Global = True
            .Pattern = "\d{5}" ''连续两个数字
            Debug.Print .Replace(sr, "")
           End With
           
         End Sub
   ''2  {m,n}最小重复m次,最多重复n次
     
        Sub t22()
           Dim regx As New RegExp
           Dim sr
           sr = "A234CA7A6789"
           With regx
            .Global = True
            .Pattern = "\d{4,5}" ''连续两个数字或连续三个数字
            Debug.Print .Replace(sr, "")
           End With
         End Sub
    ''3 {m,} 最少重复m次,相当于+
         Sub t23()
           Dim regx As New RegExp
           Dim sr
           sr = "A2348t6CA7A67"
           With regx
            .Global = True
            .Pattern = "\d{2,}" ''连续两个数字或连续三个数字
            Debug.Print .Replace(sr, "")
           End With
         End Sub
         
''* 可以出现0等任意次   相当于 {0,}，比如："\^*b"可以匹配 "b","^^^b"...

'' ?
  ''1 匹配表达式0次或者1次，相当于 {0,1}，比如："a[cd]?"可以匹配 "a","ac","ad"

        Sub t24()
           Dim regx As New RegExp
           Dim sr
           sr = "A23.48CA7A6..7"
           With regx
            .Global = True
            .Pattern = "\d+\.?\d+" ''最多连续1个
            Debug.Print .Replace(sr, "")
           End With
         End Sub
    ''2 利用+?的格式可以分段匹配
          
      Sub t87()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "<td><p>aa</p></td> <td><p>bb</p></td>"
        With regex
          .Global = True
          .Pattern = "<td>.*?</td>"
         Set mat = .Execute(sr)
          For Each m In mat
            Debug.Print m
          Next m
        End With
      End Sub
      
     Sub t88()
              
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = " aba  aca  ada "
        With regex
          .Global = True
          .Pattern = "\s.+?\s"
         Set mat = .Execute(sr)
          For Each m In mat
            Debug.Print m
          Next m
        End With

     End Sub
     

 

''第39集
Option Explicit

''^符号：限制的字符在最前面,如^\d表示以数字开头
 
    Sub T34()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "d234我345d43"
        With regex
          .Global = True
          .Pattern = "^\d*"
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub

''$符号：限制的字符在最后面，如 A$表示最后一个字符是A

   
    Sub T3433()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "R243r"
        With regex
          .Global = True
           .Pattern = "^\D.*\D$"
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub

''\b
  ''空格(包含开头和结尾)
  
        Sub t26()
           Dim regx As New RegExp
           Dim sr
           sr = "A12dA56 A4"
           With regx
            .Global = True
            .Pattern = "\bA\d+"
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
    
    Sub T272()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "ad bf cr de ee"
        With regex
          .Global = True
           .Pattern = ".+?\b"
            Set mat = .Execute(sr)
            For Each m In mat
              If m <> " " Then Debug.Print m
            Next m
        End With
      End Sub
''|
 ''可以设置两个条件,匹配左边或右边的
        Sub t27()
           Dim regx As New RegExp
           Dim sr
           sr = "A12DA56 A4B34D"
           With regx
            .Global = True
            .Pattern = "A\d+|B\d+"
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
''\un 匹配 n，其中 n 是以四位十六进制数表示的 Unicode 字符。
''汉字一的编码是4e00,最后一个代码是9fa5
       Sub t2722()
           Dim regx As New RegExp
           Dim sr
           sr = "A12d我A爱56你 A4"
           With regx
            .Global = True
            .Pattern = "[\u4e00-\u9fa5]"
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub

Option Explicit

''()
  ''可以让括号内作为一个整体产生重复
   
        Sub t29()
           Dim regx As New RegExp
           Dim sr
           sr = "A3A3QA3A37BDFE87A8"
           With regx
            .Global = True
            .Pattern = "((A3){2})" ''相当于A3A3
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
  ''取匹配结果的时候，括号中的表达式可以用 \数字引用

        Sub t30()
           Dim regx As New RegExp
           Dim sr
           sr = "A3A3QA3A37BDFE87A8"
           With regx
            .Global = True
            .Pattern = "((A3){2})Q\1"
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
          Sub t31()
           Dim regx As New RegExp
           Dim sr
           sr = "A3A3B4B4QB4B47BDFE87A8"
           With regx
            .Global = True
            .Pattern = "((A3){2})((B4){2})Q\4"
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
        
''用(?=字符)可以先进行预测查找，到一个匹配项后，将在匹配文本之前开始搜索下一个匹配项。 不会保存匹配项以备将来之用。
  
  ''例：截取某个字符之前的数据
      Sub t343()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "100元8000元57元"
        With regex
          .Global = True
           .Pattern = "\d+(?=元)" ''查找任意多数字后的元，查找到后从元以前开始查找（因为元前的数字已被使用，
                                  ''所以只能从元开始查找）匹配 ()后面的，因为后面没有设置，所以只显示前面的数字，元不再显示
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub
   ''例：验证密码，条件是4-8位，必须包含一个数字
      Sub t355()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "A8ayaa"
        With regex
          .Global = True
           .Pattern = "^(?=.*\d).{4,8}$"
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub
      
''用(?!字符)可以先进行负预测查找，到一个匹配项后，将在匹配文本之前开始搜索下一个匹配项。 不会保存匹配项以备将来之用。
     Sub t356()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "中国建筑集团公司"
        With regex
          .Global = True
           .Pattern = "^(?!中国).*"
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub
 
''()与|一起使用可以表示or

      Sub t344()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "100元800块7元"
        With regex
          .Global = True
           .Pattern = "\d+(元|块)"
           ''.Pattern = "\d+(?=元|块)"
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub


Option Explicit

''[]
 ''使用方括号 [ ] 包含一系列字符，能够匹配其中任意一个字符。用 [^ ] 不包含一系列字符，
 ''则能够匹配其中字符之外的任意一个字符。同样的道理，虽然可以匹配其中任意一个，但是只能是一个，不是多个
 
  ''1 和括号内的其中一个匹配
        Sub t29()
           Dim regx As New RegExp
           Dim sr
           sr = "ABDC"
           With regx
            .Global = True
            .Pattern = "[BC]"
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
        
   ''2 非括号内的字符
        
        Sub T35()
           Dim regx As New RegExp
           Dim sr
           sr = "ABCDBDC"
           With regx
            .Global = True
            .Pattern = "[^BC]"
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
   ''3 在一个区间
        Sub t38()
           Dim regx As New RegExp
           Dim sr
           sr = "ABCDGWDFUFE"
           With regx
            .Global = True
            .Pattern = "[a-h]"
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
        Sub t40()
           Dim regx As New RegExp
           Dim sr
           sr = "124325436789"
           With regx
            .Global = True
            .Pattern = "[1-47-9]"
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub

''第40集

Option Explicit

''()
  ''可以让括号内作为一个整体
   
        Sub t29()
           Dim regx As New RegExp
           Dim sr
           sr = "A3A3QA3A37BDFEA387A8"
           With regx
            .Global = True
            .Pattern = "(A3){2}" ''相当于A3A3
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
  ''取匹配结果的时候，括号中的表达式可以用 \数字引用

        Sub t30()
           Dim regx As New RegExp
           Dim sr
           sr = "A3A3QA3A37BDFEA387A8"
           With regx
            .Global = True
            .Pattern = "((A3){2})Q\1" ''A3A3QA3A3
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
          Sub t31()
           Dim regx As New RegExp
           Dim sr
           sr = "A3A3B4B4QB4B47BDFE87A8"
           With regx
            .Global = True
            .Pattern = "((A3){2})((B4){2})Q\4"
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
        
''用(?=字符)可以先进行预测查找，到一个匹配项后，将在匹配文本之前开始搜索下一个匹配项。 不会保存匹配项以备将来之用。
  
  ''例：截取某个字符之前的数据
      Sub t343()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "100元8000元57元"
        With regex
          .Global = True
           .Pattern = "\d+(?=元)." ''查找任意多数字后的元，查找到后从元以前开始查找,查找和\d匹配的。
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub
   ''例：验证密码，条件是4-8位，必须包含一个数字
      Sub t355()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "A8ayaa"
        With regex
          .Global = True
           .Pattern = "^(?=.*\d).{4,8}$"
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub
      
''用(?!字符)可以先进行负预测查找，到一个匹配项后，将在匹配文本之前开始搜索下一个匹配项。 不会保存匹配项以备将来之用。
     Sub t356()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "建筑集团公司"
        With regex
          .Global = True
           .Pattern = "^(?!中国).*"
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub
 
''()与|一起使用可以表示or

      Sub t344()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "100元800块7元"
        With regex
          .Global = True
          '' .Pattern = "\d+(元|块)"
           
           .Pattern = "\d+元|\d+块"
           
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub

''第41集

Option Explicit

Sub 按钮1_单击()
     Dim regx As New RegExp
     Dim sr, x, mat, m
  For x = 2 To Range("a65536").End(xlUp).Row
     sr = Cells(x, 1)
     With regx
      .Global = True
      .Pattern = Cells(x, 2)
      If Cells(x, 5) = 1 Then
       Cells(x, 3) = .Replace(sr, "")
      Else
      If .test(sr) = False Then
       Cells(x, 3) = "没有匹配的"
       Else
       Cells(x, 3) = .Execute(sr)(0)
       End If
      End If
     End With
    Next x
End Sub

''第42集
Option Explicit
''1 数据类型综述
   ''在VBA中的数据类型有整数、文本、对象等类型。这些不同的类型有着特定的作用，在进行运算时也会占用
   ''不同大小的内存，所以我们在编写程序时为了提高运行效率，一般都要定义数据的类型。
   
''2 数据类型对程序运行的影响
     ''byte                       占用1个字节
     ''integer,boolean            占用2个字节
     ''long,single                占用4个字节
     ''Double,Currency,date       占用8个字节
     ''object                     占用4个字节
     '''string(不定长)             占用10+字符长度个字节
     '''string(定长)               占用字符串长度个字节
     ''Variant(任意数字类型)      占用16个字节
     ''Variant(字符串)            占用24+字符串长度个字节
   Sub sss1()
      Dim x As Long
      Dim t
      ''Dim k1 As Byte     ''用时0.03125s
      Dim k
      ''Dim k1 As Integer ''用时0.15625s
      Dim k1 As String   ''用时0.203125s
      k = 1
      t = Timer
      For x = 1 To 1000000
        k1 = k
      Next x
      Debug.Print Timer - t
    End Sub



''1 检查是否为空
   Sub s1()
     Debug.Print Range("a1") = "" ''判断真空,无法判断假空
     Debug.Print Len(Range("a1")) = 0 ''判断真空，无法判断假空
     Debug.Print VBA.IsEmpty(Range("a1")) ''假空时返回FALSE
     Debug.Print VBA.TypeName(Range("a1").Value) ''返回Empty表示为空
   End Sub
   
   Sub 速度测试()
     Dim t
     Dim x As Long
     t = Timer
     For x = 1 To 100000
       ''If Range("a1") = "" Then ''用时0.81
      '' If Len(Range("a1")) = 0 Then ''0.84
      '' If VBA.IsEmpty(Range("a1")) Then ''速度 0.79
       ''If VBA.TypeName(Range("a1").Value) = Empty Then ''0.84
       End If
     Next x
   Debug.Print Timer - t
   End Sub

''2 检查是否为数字
   Sub s2()
    Debug.Print VBA.IsNumeric(Range("a1"))
    Debug.Print Application.WorksheetFunction.IsNumber(Range("A1"))
    Debug.Print VBA.TypeName(Range("A1").Value)
   '' Debug.Print Range("a1").Value Like "#" ''判断一位整数
   '' Debug.Print Range("a1") Like "*#*" ''判断是否包含数字
   End Sub
      Sub 速度测试2()
     Dim t
     Dim x As Long
     t = Timer
     For x = 1 To 100000
       ''If VBA.IsNumeric(Range("a1")) Then ''用时0 0.79
       ''If Application.WorksheetFunction.IsNumber(Range("A1")) Then ''0.9218
       ''If VBA.TypeName(Range("A1").Value) = "Double" Then ''速度 0.84
       End If
     Next x
   Debug.Print Timer - t
   End Sub

''3 检查是否为文本
   Sub t3()
     Debug.Print Application.IsText(Range("a1"))
     Debug.Print "B" Like "[A-Za-z]" ''判断是否为字母
     Debug.Print Len(Range("a1"))
     Debug.Print Range("a1") Like "*[一－龥]*" ''判断字符串中是否包含汉字
   End Sub

''4 判断结果是否为错误值
  Sub s4()
    Debug.Print VBA.IsError(Range("a1"))
    Debug.Print TypeName(Range("a1").Value)
  End Sub
  
''5 判断是否为数组
   Sub s5()
     Dim arr
     arr = Range("A1:A2")
     Erase arr
     Debug.Print VBA.IsArray(arr)
   End Sub
''6 判断是否为日期
   Sub s6()
      Debug.Print VBA.IsDate(Range("a2"))
   End Sub
   
Option Explicit

''一、类型转换函数：CBool, CByte, CCur, CDate, CDbl, CDec, CInt, CLng, CSng, CStr, CVar

''上述函数是把表达式转换成相对应的数字类型，比如clng转换成长整型,cstr转换成文本型

Sub ss1()
 Dim s As Integer
 s = 2334
 MsgBox 截取(CStr(s)) ''因为自定义函数参数要求是文本类型，而s是数值类型，所以需要用cstr转换成文本类型
End Sub

Function 截取(x As String)
  截取 = Left(x, 2)
End Function

Sub ss2()
 Debug.Print 1 + True ''CInt(1 = 1)
End Sub

''二、Format函数
 
  ''format函数用法等同于工作表中的text函数，可以格式化显示数字或文本
 
 Sub ss3()
  Dim n, n1
  n = 234.3372
  n1 = 41105
  Debug.Print Format(n, "0.00")
  Debug.Print Format(n, "0")
  Debug.Print Format(n, "\价格\:0.00")
  Debug.Print Format(n1, "yyyy-mm-dd")
 End Sub

''第43集

Option Explicit
Dim k
Sub ttt1()
Application.OnTime TimeValue("15:46:00"), "A"
End Sub
Sub a()
  MsgBox "test"
End Sub
Sub ttt2()
Application.OnTime Now + TimeValue("00:00:02"), "A"
End Sub

Sub 时间显示()
  Dim x
  If k = 1 Then
    k = 0
   End
  End If
  Range("a1") = Format(Now, "h:mm:ss")
  Application.OnTime Now + TimeValue("00:00:01"), "时间显示"
  x = DoEvents
End Sub

Sub 结束时间显示()
 k = 1
End Sub

Option Explicit

''1 返回当前日期、时间（指本机系统设置的日期和时间）
  Sub t1()
    Debug.Print Date ''返回当前日期
    Debug.Print Time ''返回当前时间
    Debug.Print Now  ''返回当前日期+时间
  End Sub
  
''2 格式化显示日期
   Sub t2()
     Debug.Print Format(Now, "yyyy-mm-dd")
     Debug.Print Format(Now, "yyyy年mm月dd日")
     Debug.Print Format(Now, "yyyy年mm月dd日 h:mm:ss")
     Debug.Print Format(Now, "d-mmm-yy") ''英文月份
     Debug.Print Format(Now, "d-mmmm-yy") ''英文月份
     Debug.Print Format(Now, "aaaa") ''中文星期
     Debug.Print Format(Now, "ddd") ''英文星期前三个字母
     Debug.Print Format(Now, "dddd") ''英文星期完整显示
   End Sub
''3 根据年月日返回日期
   Sub t3()
     Debug.Print VBA.DateSerial(2011, 10, 1)
   End Sub
''4 根据小时分种返回时间
   Sub t4()
     Debug.Print VBA.TimeSerial(1, 2, 1)
   End Sub

''5 返回年月日小时分秒

  Sub t5()
  Dim d
    d = "2011-10-28 01:10:03"
    Debug.Print Year(d) & "年"
    Debug.Print Month(d) & "月"
    Debug.Print Day(d) & "日"
    Debug.Print Hour(d) & "时"
    Debug.Print VBA.Minute(d) & "分"
    Debug.Print Second(d) & "秒"
  End Sub


Option Explicit

''1 计算两个日期相隔天数,月数,年数,小时,分种,秒
  
   Sub tt1()
   Dim d1, d2 As Date
    d1 = #11/21/2011#
    d2 = #12/1/2011#
    Debug.Print "相隔" & (d2 - d1) & "天"
    Debug.Print "相隔" & DateDiff("d", d1, d2) & "天"
    Debug.Print "相隔" & DateDiff("m", d1, d2) & "月"
    Debug.Print "相隔" & DateDiff("yyyy", d1, d2) & "年"
    Debug.Print "相隔" & DateDiff("q", d1, d2) & "季"
    Debug.Print "相隔" & DateDiff("w", d1, d2) & "周"
    Debug.Print "相隔" & DateDiff("h", d1, d2) & "小时"
    Debug.Print "相隔" & DateDiff("n", d1, d2) & "分种"
    Debug.Print "相隔" & DateDiff("s", d1, d2) & "秒"
   End Sub
   
    Sub tt2() ''计算两时间的差
      Dim t, x
      t = Timer
      For x = 1 To 10000000
      Next x
      Debug.Print Timer - t
    End Sub

''2 日期时间加减
   Sub tt3()
    Dim d1, d2 As Date
    d1 = "2001-10-1 00:00:00"
    Debug.Print VBA.DateAdd("d", 10, d1) ''加上10天
    Debug.Print VBA.DateAdd("m", 10, d1) ''加上10个月
    Debug.Print VBA.DateAdd("yyyy", 10, d1) ''加上10年
    Debug.Print VBA.DateAdd("yyyy", -10, d1) ''减少10年
    Debug.Print VBA.DateAdd("h", 10, d1) ''加上10小时后的时间
    Debug.Print VBA.DateAdd("n", 10, d1) ''加上10分种后的时间
    Debug.Print VBA.DateAdd("s", 10, d1) ''加上10秒后的时间
   End Sub


''第44集

Option Explicit

''1 添加图形
     ''添加shape对象的代码都可以用录制宏的方法得到,大家想了解添加的方法,就去录制一个宏吧.
     ''如下面是添加各种图形的宏
Sub Macro1()
''
'' Macro1 Macro
'' 宏由 Lenovo User 录制，时间: 2011-12-17
''
    ActiveSheet.Pictures.Insert ("D:\My Documents\My Pictures\兰色幻想头像.jpg") ''插入图片
    ActiveSheet.Shapes.AddLine(391.5, 214.5, 513.75, 273#).Select ''添加直线
    ActiveSheet.Shapes.AddShape msoShapeRectangle, 468#, 148.5, 94.5, 39.75 ''添加矩形
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 407.25, 308.25, _
        73.5, 90#).Select ''添加文本框
    ActiveSheet.Buttons.Add(534.75, 241.5, 96.75, 41.25).Select ''添加窗体中的控件
    ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False _
        , DisplayAsIcon:=False, Left:=603.75, Top:=160.5, Width:=83.25, Height _
        :=20.25).Select ''添加控件工具箱控件
End Sub

''2 控制图形的位置
      ''同样可以利用录制宏可以录制出代码。
Sub Macro2()
''
'' Macro2 Macro
'' 宏由 Lenovo User 录制，时间: 2011-12-17
    ActiveSheet.Shapes("Picture 6").Select
    Selection.ShapeRange.IncrementLeft -27#  ''水平向左移动
    Selection.ShapeRange.IncrementTop -51#   ''垂直向上移动
    
End Sub



''Option Explicit
'''shapes 对象，
    ''该对象代表工作表或图形工作表上的所有图形,它是sheets和chart的子对象（属性）

Sub t2()
  On Error Resume Next
  Dim ms As Shape
  k = 1
  For Each ms In Sheet1.Shapes
    k = k + 1
    Cells(k, 1) = ms.Name
    Cells(k, 2) = ms.Type
    Cells(k, 3) = ms.BottomRightCell.Address
    Cells(k, 4) = ms.TopLeftCell.Address
    Cells(k, 5) = ms.Hyperlink.Address
    Cells(k, 6) = ms.Visible
    Cells(k, 7) = ms.OnAction
    Cells(k, 8) = ms.Top
    Cells(k, 9) = ms.Width
    Cells(k, 10) = ms.Height
    Cells(k, 11) = ms.Left
  Next ms
End Sub

Sub test()
Sheet1.Shapes(1).Visible = True
End Sub


Option Explicit

Sub 批量插入复选框()
 Dim RG As Range
   Dim S As Shape
  For Each S In ActiveSheet.Shapes
   If InStr(S.Name, "Ch") > 0 Then
      S.Delete
   End If
 Next S
 For Each RG In Range("B2:B15")
    ActiveSheet.CheckBoxes.Add(RG.Left, RG.Top + 5, RG.Width - 20, RG.Height - 20).Select
    With Selection
        .Characters.Text = "是"
        .Value = xlOff
        .LinkedCell = RG.Address
    End With
 Next RG
End Sub



Option Explicit
Sub 连线()
  Dim rg1 As Range, rg2 As Range, rg3 As Range, rg4 As Range
''删除已有线条
 Dim S As Shape
 Dim RG As Range
 For Each S In ActiveSheet.Shapes
   If S.Type = 9 Then
     S.Delete
   End If
 Next S
 ''连线
  Set rg1 = Range("b7")
  Set rg2 = Range("b8")
  Set rg4 = Range("c10")
  Set rg3 = Range("c12")
    ActiveSheet.Shapes.AddLine(rg1.Left, rg1.Top + rg1.Height / 2, rg3.Left, rg3.Top + rg3.Height / 2).Select
    Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
    ActiveSheet.Shapes.AddLine(rg2.Left, rg2.Top + rg2.Height / 2, rg4.Left, rg4.Top + rg4.Height / 2).Select
    Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
End Sub



Option Explicit


Sub 文本框7_单击()
  Application.Caller.Name
End Sub

Option Explicit

Sub 图片导入()
''删除已有图片
 Dim S As Shape
 Dim RG As Range
 For Each S In ActiveSheet.Shapes
   If S.Type <> 8 Then
     S.Delete
   End If
 Next S
''导入图片
   
  For Each RG In Range("b2:b5")
   '' Range("B2").Select
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, RG.Left, RG.Top, RG.Width, RG.Height).Select
    Selection.ShapeRange.Fill.UserPicture "E:\80集视频\VBA视频80集\VBA80集第44集\" & RG.Offset(0, -1) & ".jpg"
  Next RG
End Sub




''第45集
Option Explicit

Sub 随机挑选演示程序1()
  Dim arr
  Dim x As Integer, num As Integer, k As Integer
  Range("c1:c10") = ""
  Range("a1:a10") = Application.Transpose(Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J"))
  For x = 1 To 10
     num = (Rnd() * (10 - 1) + 1) \ 1
     Range("a1:a10").Interior.ColorIndex = xlNone
     Range("a" & num).Interior.ColorIndex = 6
     Range("c" & x) = Range("a" & num)
  Next x
End Sub
Sub 移形换位演示程序()
  Dim arr
  Dim x As Integer, num As Integer, k As Integer, sr As String
  Range("c1:c10") = ""
  Range("a1:a10") = Application.Transpose(Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J"))
  For x = 1 To 10
     num = (Rnd() * ((10 - x + 1) - 1) + 1) \ 1
     Range("a1:a" & (10 - x + 1)).Interior.ColorIndex = xlNone
     Range("a" & num).Interior.ColorIndex = 6
     Range("c" & x) = Range("a" & num)
     ''下面开始换位
      sr = Range("a" & num)
      Range("a" & num) = Range("a" & (10 - x + 1))
      Range("a" & (10 - x + 1)) = sr
      Range("a" & (10 - x + 1)).Interior.ColorIndex = 1
  Next x
End Sub


Option Explicit

Sub 随机抽取字典法()
 Dim d As Object
 Dim arr, num As Integer, x As Integer, arr1(1 To 20000, 1 To 1) As String, t
 t = Timer
 Set d = CreateObject("scripting.dictionary")
 arr = Range("a1:a20000")
 For x = 1 To 20000
100:
   num = Rnd() * (20000 - 1) + 1
   If d.exists(num) Then
     GoTo 100
   Else
     d(num) = ""
     arr1(x, 1) = arr(num, 1)
   End If
 Next x
    Range("c1:c20000") = ""
    Range("c1:c20000") = arr1
   [d65536].End(xlUp).Offset(1, 0) = Timer - t
End Sub



Option Explicit
''提速依据
   ''在换位时数字的换位速度要比文本型要快。所以借力数值型数组达到提速的目的
Sub 移形随机排序()
   Dim arr
   Dim arr1(1 To 20000, 1 To 1) As String, sr As String
   Dim x As Integer, num, t
   t = Timer
   arr = Range("a1:a20000")
   For x = 1 To UBound(arr)
      num = (Rnd() * ((20000 - x + 1) - 1) + 1) \ 1
      arr1(x, 1) = arr(num, 1)
      ''换位
      sr = arr(num, 1)
      arr(num, 1) = arr(20000 - x + 1, 1)
      arr(20000 - x + 1, 1) = sr
   Next x
   Range("c1:c20000") = ""
   Range("c1:c20000") = arr1
   [d65536].End(xlUp).Offset(1, 0) = Timer - t
End Sub


Sub 移形随机排序升级()
   Dim arr
   Dim arr1(1 To 20000, 1 To 1) As String, sr As Integer
   Dim x As Integer, num, t, y
   Dim arr2(1 To 20000)
   t = Timer
   arr = Range("a1:a20000")
   For y = 1 To 20000
     arr2(y) = y
   Next y
   For x = 1 To UBound(arr)
      num = (Rnd() * ((20000 - x + 1) - 1) + 1) \ 1
      arr1(x, 1) = arr(arr2(num), 1)
      ''换位
      sr = arr2(num)
      arr2(num) = arr2(20000 - x + 1)
      arr2(20000 - x + 1) = num
   Next x
   Range("c1:c20000") = ""
   Range("c1:c20000") = arr1
    [F65536].End(xlUp).Offset(1, 0) = Timer - t
End Sub

''第46集

Option Explicit
Sub jin(n)
 If n < 4 Then
   jin n + 1
   jin n + 1
 End If
End Sub

Sub 调用jin()
  jin 1
End Sub


Option Explicit
Dim k As Long
''递归基础

''1 什么是递归？
  ''递归就是自已调用自已。
 ''2,用递归有什么好处？
   ''简化代码，让程序更简捷。特别是在循环层数不定的情况下可以大大简单代码。
 ''3,递归有什么坏处？
    ''因为递归在使用时会产生大量储存临时信息的“栈”（按先进先出储存信息），所以运行效果比较低，所以一般不建议使用递归设计程序
''2 例:  计算4的阶乘 (4 * 3 * 2 * 1 = 24)
   
   Sub 一般方法()
     Dim k, x
     k = 1
     For x = 4 To 1 Step -1
        k = k * x
     Next x
     MsgBox k
   End Sub
   
   Sub 递归1()
      MsgBox s(5)
   End Sub
''函数法
   Function s(n As Integer) As Integer
     If n = 1 Then
        s = 1
     Else
       s = n * s(n - 1)
     End If
   End Function
  Sub 递归2()
    k = 1
    s2 4
    MsgBox k
  End Sub
'''sub过程法
   Sub s2(n As Integer)
    '' Dim m
     If n > 0 Then
      k = k * n
     ''m = n
      s2 n - 1
     End If
   End Sub
   
''3 例：计算1+2+3+.5
 Sub 递归3()
   k = 0
   add5 1
   ''MsgBox k
 End Sub
 
  Sub add5(n As Integer)
   If n < 5 Then
     k = k + n
     add5 n + 1
   End If
  End Sub
 

Option Explicit
Dim arr1(1 To 100, 1 To 1) ''把分组后的结果放在arr1中
Dim k As Integer ''作为arr1填充时的行数
Sub 组合()
  Dim arr
  k = 0
  Erase arr1
  arr = Range("a2:a" & Range("a65536").End(xlUp).Row)
  zuhe arr, 1, "", 0
  Range("c2").Resize(100) = ""
  Range("c2").Resize(k) = arr1
End Sub

Sub zuhe(arr, x, sr, y)
''arr 把源数组导入子过程
''x 递归的索引号
'''sr 连接的字符串
''y 连接的次数
If y = [b2] Then
 k = k + 1
 arr1(k, 1) = sr
 Exit Sub
End If
If x < UBound(arr) + 1 Then
  zuhe arr, x + 1, sr & arr(x, 1), y + 1
  zuhe arr, x + 1, sr, y
End If
 
End Sub

Option Explicit
Dim arr1(1 To 100, 1 To 1) ''把分组后的结果放在arr1中
Dim k As Integer ''作为arr1填充时的行数
Sub 组合()
  Dim arr
  k = 0
  Erase arr1
  arr = Range("a2:a" & Range("a65536").End(xlUp).Row)
  zuhe arr, 1, "", 0
  Range("c2").Resize(100) = ""
  Range("c2").Resize(k) = arr1
End Sub

Sub zuhe(arr, x, sr, y)
''arr 把源数组导入子过程
''x 递归的索引号
'''sr 连接的字符串
''y 连接的次数
If y = [b2] Then
 k = k + 1
 arr1(k, 1) = sr
 Exit Sub
End If
If x < UBound(arr) + 1 Then
  zuhe arr, x + 1, sr & arr(x, 1), y + 1
  zuhe arr, x + 1, sr, y
End If
 
End Sub


Option Explicit
Dim arr1(1 To 10000, 1 To 1)  As String  ''公式表达式放在arr1中
Dim k As Integer ''作为arr1填充时的行数
Dim g As Integer, h As Integer
Dim arr
Dim k1
Sub 组合()
  k = 0
  Dim t
  t = Timer
  Erase arr1
  arr = Range("a2:a" & Range("a65536").End(xlUp).Row)
  g = [b2]
  h = [c2]
  zuhe 1, 0, "", 0
  Range("d2").Resize(k) = arr1
  [e1] = k1
  MsgBox "找到 " & k & " 个解! 花费" & Format(Timer - t, "0.00") & "秒"
End Sub

Sub zuhe(x%, z%, sr$, gg As Byte)
    If z + arr(x, 1) = h And gg = g - 1 Then
        k = k + 1
        arr1(k, 1) = sr & arr(x, 1) & "=" & h
        Exit Sub
    End If
    If x < UBound(arr) And z < h Then
      If z + arr(x, 1) < h Then
       zuhe x + 1, z + arr(x, 1), sr & arr(x, 1) & "+", gg + 1
      End If
      zuhe x + 1, z, sr, gg
    End If
End Sub

Sub 循环法()
  Dim x As Integer
  Dim y As Integer
  Dim z As Integer
  Dim t
  Dim arr(1 To 1000, 1 To 1) As String
  Dim q As Long, q1 As Long
  t = Timer
   For x = 1 To 97
     For y = x + 1 To 98
       For z = y + 1 To 99
       q1 = q1 + 1
       If x + y + z = 54 Then
           q = q + 1
        arr1(q, 1) = x & "+" & y & "+" & z & "=54"
      End If
   Next z, y, x
  Range("e2").Resize(10000) = ""
  Range("e2").Resize(q) = arr1
  MsgBox Timer - t
End Sub


Option Explicit
Dim arr1(1 To 10000, 1 To 1)  As String  ''公式表达式放在arr1中
Dim k As Integer ''作为arr1填充时的行数
Dim g As Integer, h As Integer
Dim arr
Dim k1
Sub 组合()
  k = 0
  Dim t
  t = Timer
  Erase arr1
  arr = Range("a2:a" & Range("a65536").End(xlUp).Row)
  g = [b2]
  h = [c2]
  zuhe 1, 0, ""
  Range("d2").Resize(k) = arr1
  [e1] = k1
  MsgBox "找到 " & k & " 个解! 花费" & Format(Timer - t, "0.00") & "秒"
End Sub

Sub zuhe(x%, z%, sr$)
    If z + arr(x, 1) = h Then
        k = k + 1
        arr1(k, 1) = sr & arr(x, 1) & "=" & h
        Exit Sub
    End If
    If x < UBound(arr) And z < h Then
      If z + arr(x, 1) < h Then
       zuhe x + 1, z + arr(x, 1), sr & arr(x, 1) & "+"
      End If
      zuhe x + 1, z, sr
    End If
End Sub


''第48集


Option Explicit

'1 判断文件夹是否存在
   'dir函数的第二个参数是vbdirectory时可以返回路径下的指定文件和文件夹，如果结果为""，则表示不存在。
  Sub w1()
    If Dir(ThisWorkbook.path & "\2011年报表2", vbDirectory) = "" Then
       MsgBox "不存在"
    Else
       MsgBox "存在"
    End If
  End Sub
  
'2 新建文件夹
   'Mikdir语句可以创建一个文件夹
    Sub w2()
      MkDir ThisWorkbook.path & "\Test"
    End Sub
   
'3 删除文件夹
   
   'RmDir语句可以删除一个文件夹，如果想要使用 RmDir 来删除一个含有文件的目录或文件夹，则会发生错误。
   '在试图删除目录或文件夹之前，先使用 Kill 语句来删除所有文件。
   
    Sub w3()
    Kill ThisWorkbook.path & "\test\*"
      RmDir ThisWorkbook.path & "\test"
    End Sub
'4 文件夹重命名
    Sub w4()
      Name ThisWorkbook.path & "\test" As ThisWorkbook.path & "\test2"
    End Sub
     
'5 文件夹移动
     '同样使用name方法，可以达到移动的效果，而且连文件夹的文件一起移动
    
    Sub w5()
      Name ThisWorkbook.path & "\test2" As ThisWorkbook.path & "\2011年报表\test100"
    End Sub
    
'6 文件夹复制
        Sub CopyFile_fso()
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFolder ThisWorkbook.path & "\测试新建文件夹", ThisWorkbook.path & "\2011年报表\"
        Set fso = Nothing
        End Sub
'7 打开文件夹
   '使用shell函数桌面管理程序打开文件夹
    Sub w7()
      Shell "explorer.exe " & ThisWorkbook.path & "\2011年报表", 1
    End Sub


Option Explicit

'遍历指定文件夹中的文件

 Sub 遍历文件()
  Dim Filename As String, mypath As String, k As Integer
  mypath = ThisWorkbook.path & "\2011年报表\1月\A公司\"
  Range("A1:A10") = ""
  Filename = Dir(mypath & "*.xls")
  Do
    k = k + 1
    Cells(k, 1) = Filename
    Filename = Dir
  Loop Until Filename = ""
 End Sub
 Sub 遍历子文件()
  Dim Filename As String, mypath As String, k As Integer
  mypath = ThisWorkbook.path & "\2011年报表\"
  Range("A1:A10") = ""
  Filename = Dir(mypath, vbDirectory)
  Do
    If Not Filename Like "*.*" Then
      k = k + 1
      Cells(k, 1) = Filename
    End If
    Filename = Dir
  Loop Until Filename = ""
 End Sub


 

   