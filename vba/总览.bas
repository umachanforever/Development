Attribute VB_Name = "ģ��10"
''��һ�� ʲôVBA

Sub test()
 Range("a1") = 100
End Sub


Sub ����100()
''
'' ����100 Macro
'' ���� Lenovo User ¼�ƣ�ʱ��: 2011-4-22
''

''
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "100"
    Range("B4").Select
End Sub
Sub ɾ��A1������()
''
'' ɾ��A1������ Macro
'' ���� Lenovo User ¼�ƣ�ʱ��: 2011-4-22
''

''
    Range("A1").Select
    Selection.ClearContents
End Sub

Sub ����()
  Range("b2") = ""
End Sub

���ڶ��� VBA�����󷽷�������

''VBA����

  ''VBA�еĶ�����ʵ�������ǲ����ľ��з��������Ե�excel��֧�ֵĶ���

''Excel�еļ������ö����ʾ����

 ''1��������
 
      '' Workbooks �����������ϣ����еĹ�����,Workbooks(N)����ʾ�Ѵ򿪵ĵ�N��������
      '' Workbooks ("����������")
      '' ActiveWorkbook ���ڲ����Ĺ�����
      '' ThisWorkBook ''�������ڵĹ�����
      
  ''2��������
    '' '''sheets("����������")
      '''sheet1 ��ʾ��һ������Ĺ�����,Sheet2��ʾ�ڶ�������Ĺ�����....
      '''sheets(n) ��ʾ������˳�򣬵�n��������
      ''ActiveSheet ��ʾ�������������ڹ�����
      ''worksheet Ҳ��ʾ��������������ͼ�������깤����ȡ�

  ''3����Ԫ��
       ''cells ���е�Ԫ��
       ''Range ("��Ԫ���ַ")
       ''Cells(����,����)
       ''Activecell ����ѡ�л�༭�ĵ�Ԫ��
       '''selection ����ѡ�л�ѡȡ�ĵ�Ԫ���Ԫ������
''һ��VBA����

    ''VBA���Ծ���VBA���������е��ص�
    ''��ʾĳ����������Եķ�����
        
        ''����.����=����ֵ
        
    Sub ttt()
      Range("a1").Value = 100
    End Sub

    Sub ttt1()
      Sheets(1).Name = "�����������"
    End Sub

    Sub ttt2()
    
       Sheets("Sheet2").Range("a1").Value = "abcd"
    
    End Sub
    
    
    Sub ttt3()
     
      Range("A2").Interior.ColorIndex = 3
      
    End Sub


''�� ��VBA����

   ''VBA������������VBA�����ϵĶ���
     
     ''��ʾ��ĳ������������VBA�Ķ����ϣ�����������ĸ�ʽ��

        
  Sub ttt4()
  
      ţ��.�� ��ĳ̶�:=�߳���
     
      Range("A1").Copy Range("A2")
  End Sub
   
  Sub ttt5()
  
    Sheet1.Move before:=Sheets("Sheet3")
    
  End Sub
        
''VBA�еĴ���Ļ����ṹ����ɲ���




''VBA���
''һ����������
  ''���к�������һ������

Sub test()  ''��ʼ���
  
  Range("a1") = 100

End Sub   ''�������


''���������������
   
   ''���к���Է���һ��ֵ
   
Function shcount()

  shcount = Sheets.Count
  
End Function


''�����ڳ�����Ӧ�õ����

  Sub test2()
    
    Call test
    
  End Sub

 Sub test3()
 
   For x = 1 To 100   ''for next ѭ�����
      Cells(x, 1) = x
   Next x
 
 End Sub
 

''������ �ж����

Sub �ж�1() ''�������ж�
  If Range("a1").Value > 0 Then
     Range("b1") = "����"
  Else
     Range("b1") = "������0"
  End If
End Sub

Sub �ж�2() ''�������ж�
  If Range("a1").Value > 0 Then
     Range("b1") = "����"
  ElseIf Range("a1") = 0 Then
     Range("b1") = "����0"
  ElseIf Range("B1") <= 0 Then
     Range("b1") = "����"
  End If
End Sub

Sub �������ж�2()
 If Range("a1") <> "" And Range("a2") <> "" Then
   Range("a3") = Range("a1") * Range("a2")
 End If
End Sub

 Sub �ж�4()
  Range("a3") = IIf(Range("a1") <= 0, "��������", "����")
End Sub



Sub �ж�1() ''�������ж�
  Select Case Range("a1").Value
  Case Is > 0
     Range("b1") = "����"
  Case Else
     Range("b1") = "������0"
  End Select
End Sub

Sub �ж�2() ''�������ж�
  Select Case Range("a1").Value
  Case Is > 0
     Range("b1") = "����"
  Case Is = 0
     Range("b1") = "0"
  Case Else
     Range("b1") = "����"
  End Select
End Sub

Sub �ж�3()
 If Range("a3") < "G" Then
   MsgBox "A-G"
 End If
End Sub



Sub if�����ж�()
If Range("a2") <= 1000 Then
  Range("b2") = 0.01
ElseIf Range("a2") <= 3000 Then
  Range("b2") = 0.03
ElseIf Range("a2") > 3000 Then
  Range("b2") = 0.05
End If
End Sub

Sub select�����ж�()
 Select Case Range("a2").Value
 Case 0 To 1000
   Range("b2") = 0.01
 Case 1001 To 3000
   Range("b2") = 0.03
 Case Is > 3000
   Range("b2") = 0.05
 End Select
End Sub



''���Ľ� ѭ�����

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
      Cells(x, 2) = "�ϵ�"
      Exit Do
   End If
 Loop Until x = 14
End Sub
  


''���彲 ����

Dim m As Integer
''����
''һ��ʲô�Ǳ�����
''��ν���������ǿɱ�������ͺ������ڴ�����ʱ��ŵ�һ��С���ӣ����С���ӷŵ�ʲô���岻�̶���
Sub t1()
  Dim X As Integer ''x����һ������
  For X = 1 To 10
    Cells(X, 1) = X
  Next X
End Sub

''����С��������Է�ʲô��
   ''1 ������
     ''��t1
     
    ''2 ���ı�
    Sub t2()
     Dim st As String
     Dim X As Integer
     For X = 1 To 10
      st = st & "Excel��Ӣ��ѵ"
     Next X
    End Sub
    
     ''3 �Ŷ���
     
      Sub t3()
        Dim rg As Range
        Set rg = Range("a1")
        rg = 100
      End Sub
      
      
      ''4 ������
       Sub t4()
       
        Dim arr(1 To 10) As Integer, X As Integer
        For X = 1 To 10
          arr(X) = X
        Next X
        
       End Sub
 

''�������������ͺ�����
 
   ''1 ����������
      
       ''��������ļ�
       
   ''2 ΪʲôҪ��������
   
    
   ''3 ��������
      
      ''dim public
   
    
''�ġ������Ĵ������
   
   ''1 ���̼�����:���̽���������ֵ�ͷ�
       
       ''��t1
   
   ''2 ģ�鼶����:������ֵֻ�ڱ�ģ���б��֣��������ر�ʱ��ʱ�ͷ�
       ''��5
         Sub t6()
            m = 1
         End Sub
         Sub t5()
          MsgBox m
          m = 7
         End Sub
   ''3 ȫ�ּ�����: �����е�ģ���ж����Ե��ã�ֵ�ᱣ�浽EXCEL�ر�ʱ�Żᱻ�ͷš�
       
       '' public ����
       
         Sub t7()
           MsgBox qq
         End Sub
''�� �������ͷ�
    
     ''һ������£����̼������ڹ������н�����ͻ��Զ����ڴ����ͷţ���ֻ��һЩ���ⲿ���õĶ����������Ҫʹ��set ����=nothing�����ͷš�



Public qq As Integer

Sub DD()
  qq = 12
End Sub


Option Explicit


�������� ��ʽ�뺯��

Option Explicit

''һ���ڵ�Ԫ�������빫ʽ

''1����VBA�ڵ�Ԫ����������ͨ��ʽ

     Sub t1()
       Range("d2") = "=b2*c2"
     End Sub
     
     Sub t2()
      Dim x As Integer
      For x = 2 To 6
       Cells(x, 4) = "=b" & x & "*c" & x
      Next x
     End Sub

''2����VBA�ڵ�Ԫ����������ŵĹ�ʽ
     Sub t3()
     
     Range("c16") = "=SUMIF(A2:A6,""b"",B2:B6)" ''���������žͰѵ����żӱ�
     
     End Sub
      
''3����VBA�ڵ�Ԫ�����������鹫ʽ

    Sub t4()
      Range("c9").FormulaArray = "=SUM(B2:B6*C2:C6)"
    End Sub
    
''�������õ�Ԫ��ʽ����ֵ

     Sub t5()
         Range("d16") = Evaluate("=SUMIF(A2:A6,""b"",B2:B6)")
         Range("d9") = Evaluate("=SUM(B2:B6*C2:C6)")
     End Sub
  
''�������ù�������
    
     Sub t6()
        
        Range("d8") = Application.WorksheeFunction.CountIf(Range("A1:A10"), "B")
        
     End Sub

''�ġ�����VBA����

     Sub t7()
     
      Range("C20") = VBA.InStr(Range("a20"), "E")

     End Sub
     
   

''�塢��д�Զ��庯��

      Function wn()
         wn = Application.Caller.Parent.Name
      End Function
     

''���߽� VBE�༭��

''VBA���߼���VBE�༭��

''һ��VBE�Ĵ���
 ''1�����̴���
 
    ''A ��ʾ���������������
    ''B ����
    ''C ģ��
    ''D ��ģ��

 ''range("a1")=10
    
    ''��Ӧ���̴��ڵĶ����ģ�壬��ʾ���������һЩ������
     
 ''3�����봰��
    ''A ע�����ֵ�����
    ''B ��������������
    ''C ����ǿ��ת�е�����
    ''D �������к͵���
         ''�������
         ''���öϵ�
    ''E �����б��͹����б��
 ''4����������
 
 ''�������ڿ��԰����й����е�ֵ������ʾ��������Ҫ���ڳ���ĵ���

Sub d()
 Dim x As Integer, st As String
 For x = 1 To 10
    st = st & Cells(x, 1)
    Debug.Print "��" & x & "�����н��:" & st
 Next x
End Sub

 ''5�����ش���

   ''�ڱ��ش����п�����ʾ�����ж�ʱ������Ϣ������ֵ��������Ϣ�ȡ�

 Sub d1()
 Dim x As Integer, k As Integer
 For x = 1 To 10
   k = k + Cells(x, 1)
  '' If k > 26 Then
  '' Stop
  '' End If
 Next x
 End Sub

''�ڰ˽� ��֧��end���

''VBA���߼���VBE�༭��

''һ��VBE�Ĵ���
 ''1�����̴���
 
    ''A ��ʾ���������������
    ''B ����
    ''C ģ��
    ''D ��ģ��

 ''range("a1")=10
    
    ''��Ӧ���̴��ڵĶ����ģ�壬��ʾ���������һЩ������
     
 ''3�����봰��
    ''A ע�����ֵ�����
    ''B ��������������
    ''C ����ǿ��ת�е�����
    ''D �������к͵���
         ''�������
         ''���öϵ�
    ''E �����б��͹����б��
 ''4����������
 
 ''�������ڿ��԰����й����е�ֵ������ʾ��������Ҫ���ڳ���ĵ���

Sub d()
 Dim x As Integer, st As String
 For x = 1 To 10
    st = st & Cells(x, 1)
    Debug.Print "��" & x & "�����н��:" & st
 Next x
End Sub

 ''5�����ش���

   ''�ڱ��ش����п�����ʾ�����ж�ʱ������Ϣ������ֵ��������Ϣ�ȡ�

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

''Goto���,��ת��ָ���ĵط�

Sub t1()
  Dim x As Integer
  Dim sr
100:
  sr = Application.InputBox("����������", "������ʾ")
  If Len(sr) = 0 Or Len(sr) = 5 Then GoTo 100
  
End Sub

''gosub..return ,����ȥ,��������

Sub t2()
  Dim x As Integer
  For x = 1 To 10
     If Cells(x, 1) Mod 2 = 0 Then GoSub 100
  Next x
Exit Sub
100:
   Cells(x, 1) = "ż��"
   Return
End Sub

''on error resume next ''��������,��������ִ����һ��

 Sub t3()
  On Error Resume Next
  Dim x As Integer
  For x = 1 To 10
    Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
  Next x
 End Sub
 
''on error goto  ''����ʱ����ָ��������
 
  Sub t4()
  On Error GoTo 100
  Dim x As Integer
  For x = 1 To 10
    Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
  Next x
   Exit Sub
100:
   MsgBox "�ڵ�" & x & "�г�����"
  End Sub
 
 ''on error goto 0 ''ȡ��������ת
 
  Sub t5()
  On Error Resume Next
  Dim x As Integer
  For x = 1 To 10
    If x > 5 Then On Error GoTo 0
    Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
  Next x
   Exit Sub

  End Sub

''��9�� excel�ļ�����

Option Explicit

''excel�ļ��͹�����

  ''excel�ļ�����excel��������excel�ļ�����Ҫexcel�̵�֧��
  
   ''Workbooks  ���������ϣ���ָexcel�ļ�������
   
   ''Workbooks("A.xls")������ΪA��excel������
     Sub t1()
        Workbooks("A.xls").Sheets(1).Range("a1") = 100
     End Sub
   
   ''workbooks(2)������˳�򣬵ڶ����򿪵Ĺ�������
      Sub t2()
        Workbooks(2).Sheets(2).Range("a1") = 200
     End Sub
   ''ActiveWorkbook �����򿪶��excel������ʱ�������ڲ������Ǹ�����ActiveWorkbook�����������
   
   ''Thisworkbook��VBA�������ڵĹ�������������򿪶��ٸ������������۵�ǰ���ĸ��������ǻ��,thisworkbook����ָ�����ڵĹ�������

''����������

    ''Windows("A.xls"),A�������Ĵ��ڣ�ʹ��windows�������ù��������ڵ�״̬�����Ƿ����صȡ�
     Sub t3()
        Windows("A.xls").Visible = False
     End Sub
     Sub t4()
        Windows(2).Visible = True
     End Sub
    

   


Option Explicit

''Goto���,��ת��ָ���ĵط�

Sub t1()
  Dim x As Integer
  Dim sr
100:
  sr = Application.InputBox("����������", "������ʾ")
  If Len(sr) = 0 Or Len(sr) = 5 Then GoTo 100
  
End Sub

''gosub..return ,����ȥ,��������

Sub t2()
  Dim x As Integer
  For x = 1 To 10
     If Cells(x, 1) Mod 2 = 0 Then GoSub 100
  Next x
Exit Sub
100:
   Cells(x, 1) = "ż��"
   Return
End Sub

''on error resume next ''��������,��������ִ����һ��

 Sub t3()
  On Error Resume Next
  Dim x As Integer
  For x = 1 To 10
    Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
  Next x
 End Sub
 
''on error goto  ''����ʱ����ָ��������
 
  Sub t4()
  On Error GoTo 100
  Dim x As Integer
  For x = 1 To 10
    Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
  Next x
   Exit Sub
100:
   MsgBox "�ڵ�" & x & "�г�����"
  End Sub
 
 ''on error goto 0 ''ȡ��������ת
 
  Sub t5()
  On Error Resume Next
  Dim x As Integer
  For x = 1 To 10
    If x > 5 Then On Error GoTo 0
    Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
  Next x
   Exit Sub

  End Sub


''��10��  excel���������

Option Explicit

''excel������ķ���


  ''excel�������������࣬һ��������ƽ���õĹ�����(worksheet)����һ����ͼ�����ȡ��������ͳ����sheets
  
   '''sheets  �������ϣ���ָexcel���ֹ�����
   
   '''sheets("A")������ΪA��excel������
     Sub t1()
        Sheets("A").Range("a1") = 100
     End Sub
   
   ''workbooks(2)������˳�򣬵ڶ����򿪵Ĺ�������
      Sub t2()
        Sheets(2).Range("a1") = 200
     End Sub
   ''ActiveSheet �����򿪶��excel������ʱ�������ڲ������Ǹ�����ActiveSheet
   
   
Option Explicit

''1 �ж�A�������ļ��Ƿ����
    Sub s1()
     Dim X As Integer
      For X = 1 To Sheets.Count
        If Sheets(X).Name = "A" Then
          MsgBox "A���������"
          Exit Sub
        End If
      Next
      MsgBox "A����������"
    End Sub
   
''2 excel������Ĳ���

  Sub s2()
     Dim sh As Worksheet
     Set sh = Sheets.Add
       sh.Name = "ģ��"
       sh.Range("a1") = 100
  End Sub

''3 excel���������غ�ȡ������
  
 Sub s3()
    Sheets(2).Visible = True
 End Sub

''4 excel��������ƶ�

   Sub s4()
     Sheets("Sheet2").Move before:=Sheets("sheet1") '''sheet2�ƶ���sheet1ǰ��
     Sheets("Sheet1").Move after:=Sheets(Sheets.Count) '''sheet1�ƶ������й�����������
   End Sub
  
''6 excel������ĸ���
   Sub s5() ''�ڱ���������
      Dim sh As Worksheet
      Sheets("ģ��").Copy before:=Sheets(1)
       Set sh = ActiveSheet
          sh.Name = "1��"
          sh.Range("a1") = "����"
   End Sub
   
   Sub s6() ''���Ϊ�¹�����
      Dim wb As Workbook
       Sheets("ģ��").Copy
       Set wb = ActiveWorkbook
          wb.SaveAs ThisWorkbook.Path & "/1��.xls"
          wb.Sheets(1).Range("b1") = "����"
          wb.Close True
   End Sub
''7 ����������
   Sub s7()
      Sheets("sheet2").Protect "123"
   End Sub
   Sub s8() ''�жϹ������Ƿ�����˱�������
      If Sheets("sheet2").ProtectContents = True Then
        MsgBox "������������"
      Else
        MsgBox "������û����ӱ���"
      End If
   End Sub
   
 ''8 ������ɾ��
     Sub s9()
       Application.DisplayAlerts = False
         Sheets("ģ��").Delete
       Application.DisplayAlerts = True
     End Sub
''9 �������ѡȡ
     Sub s10()
       Sheets("sheet2").Select
     End Sub

  
''ϰ��

Option Explicit

Sub �ձ����ʽ����()
      Dim sh As Worksheet
      Dim co As String
      Sheets("�ձ���ģ��").Visible = True
      Sheets("�ձ���ģ��").Copy after:=Sheets(Sheets.Count)
     co = Sheets.Count - 2
      If co > 31 Then co = 1
    Sheets("�ձ���ģ�� (2)").Name = co & "�ձ���"
    Sheets("�ձ���ģ��").Visible = False
End Sub

Sub �ձ����ʽ����1()
''��������
Dim i As Integer

Dim ws As Worksheet
Dim a As String

Set ws = Sheets("�ձ���ģ��")

ws.Visible = -1

i = Val(Sheets(Sheets.Count).Name)
a = Sheets(Sheets.Count).Name
ws.Copy after:=Sheets(Sheets.Count)

If i Then

ActiveSheet.Name = i + 1 & "�ձ���"

Else

ActiveSheet.Name = "1�ձ���"

End If

ws.Visible = 0

Sheets(1).Select

End Sub


Sub ��汨��()
      On Error Resume Next
      Application.DisplayAlerts = False
      Dim wb As Workbook
      Dim x As Integer
      Dim i As String
      x = 1
      Do While x < Sheets.Count
        i = CStr(x) & "�ձ���"
       Sheets(i).Select
       Sheets(i).Copy
       Set wb = ActiveWorkbook
       wb.SaveAs ThisWorkbook.Path & "/" & Sheets(i).Name & ".xls"
    x = x + 1
     wb.Close True
    Loop
    Application.DisplayAlerts = True
End Sub


Sub ��汨2()
''��������
Dim i As Integer

Dim sh As Worksheet

For i = 1 To Sheets.Count

If Sheets(i).Name Like "*�ձ���" Then

Sheets(i).Copy

Set sh = ActiveSheet

sh.SaveAs ThisWorkbook.Path & "\" & sh.Name ''& ".xls"

ActiveWorkbook.Close True

End If

Next

End Sub


''��ʮһ�� ��Ԫ���ѡȡ

Option Explicit


''1 ��ʾһ����Ԫ��(a1)
 Sub s()
   Range("a1").Select
   Cells(1, 1).Select
   Range("A" & 1).Select
   Cells(1, "A").Select
   Cells(1).Select
   [a1].Select
 End Sub


''2 ��ʾ���ڵ�Ԫ������
   
   
   Sub d() ''ѡȡ��Ԫ��a1:c5
''     Range("a1:c5").Select
''     Range("A1", "C5").Select
''     Range(Cells(1, 1), Cells(5, 3)).Select
     ''Range("a1:a10").Offset(0, 1).Select
      Range("a1").Resize(5, 3).Select
   End Sub
   
''3 ��ʾ�����ڵĵ�Ԫ������
   
    Sub d1()
    
      Range("a1,c1:f4,a7").Select
      
      ''Union(Range("a1"), Range("c1:f4"), Range("a7")).Select
      
    End Sub
    
    Sub dd() ''unionʾ��
      Dim rg As Range, x As Integer
      For x = 2 To 10 Step 2
        If x = 2 Then Set rg = Cells(x, 1)
        
        Set rg = Union(rg, Cells(x, 1))
      Next x
      rg.Select
    End Sub
    
''4 ��ʾ��
  
    Sub h()
    
      ''Rows(1).Select
      ''Rows("3:7").Select
      ''Range("1:2,4:5").Select
       Range("c4:f5").EntireRow.Select
       
    End Sub
    
''5 ��ʾ��
    
   Sub L()
    
      '' Columns(1).Select
      '' Columns("A:B").Select
      '' Range("A:B,D:E").Select
      Range("c4:f5").EntireColumn.Select ''ѡȡc4:f5���ڵ���
       
   End Sub

''6 ���������µĵ�Ԫ���ʾ����

    Sub cc()
    
      Range("b2").Range("a1") = 100
      
    End Sub
    
''7 ��ʾ����ѡȡ�ĵ�Ԫ������

   Sub d2()
     Selection.Value = 100
   End Sub

''ϰ��

''Option Explicit

Sub ��ѡ�������()
Dim A As Range '', B
For Each A In Selection
    If IsNumeric(A) And A > 0 Then
      A = "����"
    End If
Next
End Sub

Sub ѡȡ()
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


''��12�� ���ⵥԪ��λ

Option Explicit


''1 ��ʹ�õĵ�Ԫ������

  Sub d1()
  
    Sheets("sheet2").UsedRange.Select
    
    ''wb.Sheets(1).Range("a1:a10").Copy Range("i1")
    
  End Sub


''2 ĳ��Ԫ�����ڵĵ�Ԫ������

   Sub d2()
    
      Range("b8").CurrentRegion.Select
    
   End Sub
   
   
''3 ������Ԫ������ͬ������

    Sub d3()
     
    Intersect(Columns("b:c"), Rows("3:5")).Select
  
    End Sub
   
''4 ���ö�λ����ѡȡ���ⵥԪ��
  
    Sub d4()
  
       Range("A1:A6").SpecialCells(xlCellTypeBlanks).Select
       
    End Sub
    
''5 �˵㵥Ԫ��
 
   Sub d5()
   
     Range("a65536").End(xlUp).Offset(1, 0) = 1000
     
   End Sub
  
   Sub d6()
   
     Range(Range("b6"), Range("b6").End(xlToRight)).Select
     
   End Sub
    
''ʵ��
Option Explicit

Sub t()
 Dim x As Integer
  For x = 2 To 6
    If Cells(x, 2) > 0 Then
      Cells(x, "N") = "1��"
    Else
      Cells(x, "N") = Range("b" & x).End(xlToRight).Column - 1 & "��"
    End If
  Next x
  
End Sub


 ''ϰ��
 Option Explicit

''��Ŀ1:
''   B:D�и��е�Ԫ��,���Ϊ�ǿ�,���ڸ���A���������1,��A����ʾ,
''
''Ҫ��: ����ʹ��ѭ��
''��Ŀ2:
''    �򿪱�·���µ�A.Xls�ļ� , �����ļ��е����й��������ϸ���ݺϲ���������, ��������
''
''ע:     A.xls�ļ��й�������������ϸ������������������.�������������е�����������ͬ
Sub ��һ��()
''    Intersect(Columns(1), Range("B:D"). _  ''�ָ������÷���������һ��
''    SpecialCells(xlCellTypeConstants).EntireRow) = 1
    Intersect(Columns(1), Range("B:D").SpecialCells(xlCellTypeConstants).EntireRow) = 1
End Sub
Sub ��2��()
Dim i As Integer, wbk As Workbook
Set wbk = Workbooks.Open(ThisWorkbook.Path & "\A.xls")
With ThisWorkbook.Sheets("��2��")
For i = 1 To wbk.Sheets.Count
    If i = 1 Then
       wbk.Sheets(i).UsedRange.Copy .Range("A1") ''�ѱ��⿼�ǽ�ȥ��
    Else
       wbk.Sheets(i).UsedRange.Offset(1, 0).Copy .Range("A" & .[A65536].End(xlUp).Row + 1) ''ֱ���ڵ�Ԫ���������������
    End If
Next i
wbk.Close True
End With
End Sub
Sub se2()
    Dim wb As Workbook, i As Integer
    Set wb = Workbooks.Open(ThisWorkbook.Path & "\A.xls")
    For i = 1 To wb.Sheets.Count
        wb.Sheets(i).UsedRange.Offset(1, 0).Copy ThisWorkbook.Sheets("��2��").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0) ''�����ƶ�һ�񣬲�Ȼ������
    Next i
    wb.Close
End Sub
Sub copy�÷�()
    ''Worksheets(1).Range("A1:D4").Copy Destination:=Worksheets(2).Range("E1")  ''Destination:=����ʡ��
    Worksheets(1).Range("A1:D4").Copy Worksheets(2).Range("E1")
End Sub
Sub fuzhi()
''      Dim sh As Worksheet
''      Workbooks("��12����ϰ��.xls").Sheets("��2��").Copy after:=Workbooks("��12����ϰ��0.xls").Sheets(1)

    Workbooks(1).Sheets(2).UsedRange.Offset(0, 0).Select
End Sub



''��13�� Option Explicit

''1 ��Ԫ���ֵ

   Sub x1()
    Range("b7") = Range("I3").Value
    Range("b8") = Range("c2").Text
    Range("b9") = "''" & Range("I3").Formula
   End Sub

 ''2 ��Ԫ��ĵ�ַ
   
    Sub x2()
     With Range("b2").CurrentRegion
       [b12] = .Address
       [c12] = .Address(0, 0)
       [d12] = .Address(1, 0)
       [e12] = .Address(0, 1)
       [f12] = .Address(1, 1)
     End With
    End Sub
 
 ''3 ��Ԫ���������Ϣ
    Sub x3()
      With Range("b2").CurrentRegion
        [b13] = .Row
        [b14] = .Rows.Count
        [b15] = .Column
        [b16] = .Columns.Count
        [b17] = .Range("a1").Address
      End With
    End Sub
     
 ''4����Ԫ��ĸ�ʽ��Ϣ
    Sub x4()
      With Range("b2")
        [b19] = .Font.Size
        [b20] = .Font.ColorIndex
        [b21] = .Interior.ColorIndex
        [b22] = .Borders.LineStyle
      End With
    End Sub
       
  ''5����Ԫ����ע��Ϣ
     Sub x5()
        [B24] = Range("I2").Comment.Text
     End Sub

  ''6 ��Ԫ���λ����Ϣ
     Sub x6()
        With Range("B2")
          [b26] = .Top
          [b27] = .Left
          [b28] = .Height
          [b29] = .Width
        End With
     End Sub
  ''7 ��Ԫ����ϼ���Ϣ
    Sub x7()
      With Range("b3")
        [b31] = .Parent.Name
        [b32] = .Parent.Parent.Name
      End With
    End Sub
   ''8 �����ж�
      Sub x8()
       With Range("i3")
        [b34] = .HasFormula
        [b35] = .Hyperlinks.Count
       End With
      End Sub
    ''9 ��Ԫ���������ͣ�����
      
    
''ϰ��
Sub ��һ��()
      Dim x, y As Integer
      With Range("c4").CurrentRegion
      x = .Rows.Count
      y = .Columns.Count
      [a1] = .Cells(x, y).Address(0, 0)
      End With
End Sub
Sub �ڶ���()
    Range("F3").Comment.Shape.Left = Range("E1").Left
End Sub


''��14�� ��Ԫ���ʽ


''һ���ж���ֵ�ĸ�ʽ
  ''1 �ж��Ƿ�Ϊ�յ�Ԫ��
    Sub d1()
       [b1] = ""
       ''If Range("a1") = "" Then
       ''If Len([a1]) = 0 Then
       If VBA.IsEmpty([a1]) Then
          [b1] = "��ֵ"
        End If
    End Sub
  ''2 �ж��Ƿ�Ϊ����
    Sub d2()
      [b2] = ""
      ''If VBA.IsNumeric([a2]) And [a2] <> "" Then
      ''If Application.WorksheetFunction.IsNumber([a2]) Then
        [b2] = "����"
      End If
    End Sub
  ''3 �ж��Ƿ�Ϊ�ı�
    Sub d3()
      [b3] = ""
      ''If Application.WorksheetFunction.IsText([A3]) Then
       If VBA.TypeName([a3].Value) = "String" Then
         [b3] = "�ı�"
      End If
    End Sub
  ''4 �ж��Ƿ�Ϊ����
     Sub d4()
        [b4] = ""
        If [a4] > "z" Then
          [b4] = "����"
        End If
     End Sub
  ''5 �жϴ���ֵ
  Sub d10()
      [b5] = ""
      ''If VBA.IsError([a5]) Then
      If Application.WorksheetFunction.IsError([a5]) Then
         [b5] = "����ֵ"
      End If
  End Sub
    Sub d11()
      [b6] = ""
      If VBA.IsDate([a6]) Then
         [b6] = "����"
      End If
  End Sub

''�������õ�Ԫ���Զ����ʽ
   Sub d30()
        Range("d1:d8").NumberFormatLocal = "0.00"
   End Sub


''������ָ����ʽ�ӵ�Ԫ�񷵻���ֵ
   
   ''Format�����﷨(�͹�������Text�÷�����һ��)
   
    ''Format(��ֵ,�Զ����ʽ����)
    

    
    

	Option Explicit
''Excel�е���ɫ
   
    ''Excel�е���ɫ���������ַ�ʽ��ȡ��һ����EXCEL������ɫ����һ��������QBCOLOR��������
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
    Dim �� As Integer, �� As Integer, �� As Integer
    �� = 255
    �� = 123
    �� = 100
    Range("g1").Interior.Color = RGB(��, ��, ��)
  End Sub

  
''��Ԫ��ϲ�

  Sub h1()
    
    Range("g1:h3").Merge
    
  End Sub
  
  ''�ϲ�����ķ�����Ϣ
  Sub h2()
   
   Range("e1") = Range("b3").MergeArea.Address ''���ص�Ԫ�����ڵĺϲ���Ԫ������
   
  End Sub
  
  ''�ж��Ƿ񺬺ϲ���Ԫ��
  Sub h3()
   ''MsgBox Range("b2").MergeCells
    ''MsgBox Range("A1:D7").MergeCells
    Range("e2") = IsNull(Range("a1:d7").MergeCells)
    Range("e3") = IsNull(Range("a9:d72").MergeCells)
  End Sub
  
 ''�ۺ�ʾ��
 
   ''�ϲ�H����ͬ��Ԫ��
   
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


	
''��15��

Option Explicit

Sub c1()
  Rows(4).Insert
End Sub

Sub c2() ''�����в����ƹ�ʽ
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
      Cells(x + 1, "c") = Cells(x, "c") & " С��"
      Cells(x + 1, "h") = "=sum(h" & m1 & ":h" & m2 & ")"
      Cells(x + 1, "h").Resize(1, 4).FillRight
      Cells(x + 1, "i") = ""
      x = x + 1
      m1 = m2 + 2
    End If
  Next x
End Sub

Sub dd() ''ɾ��С����
 Columns(1).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub

''
Option Explicit

 
''1 ��Ԫ������
 
    Sub t1()
      Range("a1") = "a" & "b"
      Range("b1") = "a" & Chr(10) & "b" ''���д�����
    End Sub
    
''2 ��Ԫ���ƺͼ���
    
      Sub t2()
        Range("a1:a10").Copy Range("e1") ''A1��A10�����ݸ��Ƶ�C1
      End Sub
    
      Sub t3()
        Range("a1:a10").Copy
        ActiveSheet.Paste Range("d1") ''ճ����D1
      End Sub
      
      Sub t4()
        Range("a1:a10").Copy
        Range("e1").PasteSpecial (xlPasteValues) ''ֻճ��Ϊ��ֵ
      End Sub
      Sub t5()
        Range("a1:a10").Cut
        ActiveSheet.Paste Range("f1") ''ճ����f1
      End Sub

      Sub t6()
        Range("c1:c10").Copy
        Range("a1:a10").PasteSpecial Operation:=xlAdd ''ѡ��ճ��-��
      End Sub
      
      Sub T7()
          Range("G1:G10") = Range("A1:A10").Value
      End Sub
''3 ��乫ʽ
    Sub T8()
      Range("b1") = "=a1*10"
      Range("b1:b10").FillDown ''������乫ʽ
    End Sub
    
''��16��

Option Explicit

''1 ʹ��ѭ������ (�ڵ�Ԫ���в���Ч��̫��)

''2 ���ù�������
  
    Sub c1() ''�ж��Ƿ����,��������������
      Dim hao As Integer
      Dim icount As Integer
      icount = Application.WorksheetFunction.CountIf(Sheets("�����ϸ��").[b:b], [g3])
      If icount > 0 Then
       MsgBox "����ⵥ�����Ѿ����ڣ��벻Ҫ�ظ�¼��"
       MsgBox Application.WorksheetFunction.Match([g3], Sheets("�����ϸ��").[b:b], 0)
      End If
    End Sub
    
    
''3 ʹ��Find����

    Sub c2()
      Dim r As Integer, r1 As Integer
      Dim icount As Integer
      icount = Application.WorksheetFunction.CountIf(Sheets("�����ϸ��").[b:b], [g3])
      If icount > 0 Then
       r = Sheets("�����ϸ��").[b:b].Find(Range("G3"), Lookat:=xlWhole).Row ''���Һ����һ�γ��ֵ�λ��
       r1 = Sheets("�����ϸ��").[b:b].Find([g3], , , , , xlPrevious).Row
       MsgBox r & ":" & r1
      End If
    End Sub
 

   Sub c3() ''��������һ�зǿ��е�����
    
      MsgBox Sheets("�����ϸ��").Cells.Find("*", , , , , xlPrevious).Row
    
   End Sub

   
   
   
   Option Explicit
Sub ����()
  Dim c As Integer   ''�����ڿ����еĸ���
  Dim r As Integer   ''��ⵥ����������
  Dim cr As Integer  ''�����ϸ���е�һ�����е�����
With Sheets("�����ϸ��")
    c = Application.CountIf(.[b:b], Range("g3"))
    If c > 0 Then
       MsgBox "�õ��ݺ����Ѿ����ڣ�,�벻Ҫ�ظ�¼��"
       Exit Sub
    Else
       r = Application.CountIf(Range("b6:b10"), "<>")
       cr = .[b65536].End(xlUp).Row + 1
       .Cells(cr, 1).Resize(r, 1) = Range("e3")
       .Cells(cr, 2).Resize(r, 1) = Range("g3")
       .Cells(cr, 3).Resize(r, 1) = Range("c3")
       .Cells(cr, 4).Resize(r, 6) = Cells(6, 2).Resize(r, 6).Value
       MsgBox "���������"
    End If
 End With
End Sub

Sub ����()
  Dim c As Integer   ''�����ڿ����еĸ���
  Dim r As Integer   ''��ⵥ����������
  
With Sheets("�����ϸ��")
    c = Application.CountIf(.[b:b], Range("g3"))
    If c = 0 Then
       MsgBox "�õ��ݺ��벻���ڣ�"
       Exit Sub
    Else
        r = .[b:b].Find(Range("g3"), , , , , xlNext).Row
        Range("c3") = .Cells(r, 3)
        Range("e3") = .Cells(r, 1)
        Cells(6, 2).Resize(c, 5) = .Cells(r, 4).Resize(c, 5).Value
       MsgBox "��ѯ�����"
    End If
 End With
End Sub

Sub ɾ��()
 Dim c As Integer   ''�����ڿ����еĸ���
  Dim r As Integer   ''��ⵥ����������
  
With Sheets("�����ϸ��")
    c = Application.CountIf(.[b:b], Range("g3"))
    If c = 0 Then
       MsgBox "�õ��ݺ��벻���ڣ�"
       Exit Sub
    Else
        r = .[b:b].Find(Range("g3"), , , , , xlNext).Row
        .Range(r & ":" & c + r - 1).Delete
       MsgBox "ɾ�������"
    End If
 End With
End Sub
Sub �޸�()
  Call ɾ��
  Call ����
End Sub


''��17��


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
 MsgBox "��ʽ��ֵ�����˸ı�"
End Sub

Private Sub Worksheet_Deactivate()
  MsgBox "ллʹ��sheet3"
End Sub

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
 MsgBox Target.Address
End Sub

Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
  
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub


''��18��

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
  MsgBox "����������ֹ�����¹�����"
  Application.DisplayAlerts = False
   Sh.Delete
  Application.DisplayAlerts = True
End Sub

Private Sub Workbook_BeforePrint(Cancel As Boolean)
 MsgBox "��excel�ļ���ֹ��ӡ�������ӡ�������Ա��ϵ"
 Cancel = True
End Sub


Private Sub Workbook_Activate()
 
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
 MsgBox "�������水ť��"
End Sub

����19��

Public WithEvents app As Excel.Application

Private Sub app_NewWorkbook(ByVal Wb As Workbook)
 
End Sub

Private Sub app_SheetActivate(ByVal Sh As Object)

End Sub

Private Sub app_WorkbookNewSheet(ByVal Wb As Workbook, ByVal Sh As Object)

End Sub

Private Sub app_WorkbookOpen(ByVal Wb As Workbook)
'' a = Application.InputBox("�������excel�������", "��ȫ��ʾ")
'' If a <> 123 Then
''   Wb.Close False
''End If
End Sub

Private Sub Workbook_Open()
 Set app = Excel.Application
End Sub

����20��

''****************************************************************************************************
''*                              VBA����̳�                                                         *
''*                                       --------excel��Ӣ��ѵ��:��ɫ����                           *
''****************************************************************************************************


Sub v4() ''����ʱ��0.01��
 Dim t
 t = Timer
 For x = 1 To 100000
   m = m + 1000 ''��ӵ����ڴ��е�ֵ
 Next x
 MsgBox Timer - t
End Sub

Sub v5() ''����ʱ��0.5��
 Dim t
 t = Timer
 For x = 1 To 100000
   m = m + Cells(1, 1) ''���õ�Ԫ���е�ֵ
 Next x
  MsgBox Timer - t
End Sub


''1��ʲô��VBA�����أ�
    
    ''VBA������Ǵ���һ�����ݵ����ݿռ�?�������Ϳ�������,�������ı�,�����Ƕ���,Ҳ������VBA����.
    
''2 VBA���������̬
     '' VBA�������Ա�����ʽ��ŵ�һ���ռ�,��Ҳ�������У�Ҳ��������ά�ռ䡣
           
    ''1) ��������
          ''array(1,2)
          ''array(array(1,2,4),array("a","b","c"))
    ''2) ��̬����
          ''x(4) ''��5��λ�ã���Ŵ�0~4
          ''arr(1 to 10) ''��10��λ�ã����1~10
          ''arr(1 to 10,1 to 2) ''10��2�еĿռ䣬�ܹ�20��λ�ã����Ƕ�ά����
          ''arr(1 to 10,1 to 2,1 to 3) ''��ά���飬��10*2*3=60��λ�á�������ά����
    ''3)��̬����
          ''arr() ''��֪���ж����ж�����


Option Explicit

''��VBA������д������
   
   ''1�������(��)д��Ͷ�ȡ
   
     Sub t1() ''д��һά����
     Dim x As Integer
     Dim arr(1 To 10)
   arr(2) = 190
   arr(10) = 5
     End Sub
  
    Sub t2() ''���ά����д�����ݺͶ�ȡ
     Dim x As Integer, y As Integer
     Dim arr(1 To 5, 1 To 4)
     For x = 1 To 5
       For y = 1 To 4
         arr(x, y) = Cells(x, y)
       Next y
     Next x
    MsgBox arr(3, 1)
    End Sub
    
   ''2����̬����
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
       
   ''3������д��
    
      Sub t4() ''�ɳ������鵼��
      Dim arr
      arr = Array(1, 2, 3, "a")
      Stop
      End Sub
    
     Sub t5() ''�ɵ�Ԫ��������
       Dim arr
       arr = Range("a1:d5")
       Stop
     End Sub
	 
	 
''��21��

Option Explicit

''VBA����
  ''1�����ڴ��ж�ȡ
      
       ''���ڴ��ж�ȡ�����ڼ������㣬ֱ��������ĸ�ʽ
        
        ''�������(5)
        ''�������(3,2)
     ''��:
        Sub d1()
         Dim arr, arr1()
         Dim x As Integer, k As Integer, m As Integer
         arr = Range("a1:a10") ''�ѵ�Ԫ���������ڴ�������
         m = Application.CountIf(Range("a1:a10"), ">10") ''�������10�ĸ���
         ReDim arr1(1 To m)
         For x = 1 To 10
           If arr(x, 1) > 10 Then
              k = k + 1
              arr1(k) = arr(x, 1)
           End If
         Next x
        End Sub
        
        
  ''2����ȡ���뵥Ԫ����
          
      Sub d2() ''��ά������뵥Ԫ��
        Dim arr, arr1(1 To 5, 1 To 1)
        Dim x As Integer
        arr = Range("b2:c6")
        For x = 1 To 5
          arr1(x, 1) = arr(x, 1) * arr(x, 2)
        Next x
        Range("d2").Resize(10) = arr1
      End Sub
      
      Sub d3() ''һά������뵥Ԫ��
        Dim arr, arr1(1 To 5)
        Dim x As Integer
        arr = Range("b2:c6")
        For x = 1 To 5
          arr1(x) = arr(x, 1) * arr(x, 2)
        Next x
        ''Range("a13").Resize(1, 5) = arr1
        Range("d2").Resize(5) = Application.Transpose(arr1)
      End Sub
       
      Sub d4() ''���鲿�ִ���
        Dim arr, arr1(1 To 10000, 1 To 1)
        Dim x As Integer
        arr = Range("b2:c6")
        For x = 1 To 5
          arr1(x, 1) = arr(x, 1) * arr(x, 2)
        Next x
        Range("d2").Resize(5) = arr1
      End Sub

''��22��

Option Explicit

''1������Ĵ�С
''�������ñ������ģ���ô��λ��һ������Ĵ�С��

 ''Lbound(����) ���Ի�ȡ�������С�±�(���)
 ''Ubound(����) ���Ի�ȡ���������ϱ�(���)
 ''Ubound(����,1) ���Ի��������з���(��1ά)����ϱ�
 ''Ubound(����,2) ���Ի��������з���(��2ά)������ϱ�

Sub d6()
 Dim arr
 Dim k, m
 arr = Range("a2:d5")
 For x = 1 To UBound(arr, 1)
   
 Next x
End Sub


''2����̬����Ķ�̬����
   
     ''���һ�������޷��򲻷��������ܵĴ�С������һЩ����������ֲ������п�λ����ʱ���Ǿ���Ҫ�ö�̬�ĵ��뷽��
  ''
     ''ReDim Preserve arr() ��������һ����̬��С�����飬���ҿ��Ա���ԭ������ֵ�����൱�ڳ���С�ˣ����Ը��������󣬵�����ֻ��
        ''����δάʵ�ֶ�̬�������һά��������δά��ֻ��һά
    
     ����1��sheet1������
     
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
   
''3 �������
     ''�������ʹ��earse���
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
     

''��23��

Option Explicit

''
''1 �������ֵ
    Sub s()
    Dim arr1()
    
    arr1 = Array(1, 12, 4, 5, 19)
    
    MsgBox "1, 12, 4, 5, 19���ֵ" & Application.Max(arr1)
    MsgBox "1, 12, 4, 5, 19��Сֵ:" & Application.Min(arr1)
    MsgBox "1, 12, 4, 5, 19�ڶ���ֵ��" & Application.Large(arr1, 2)
    MsgBox "1, 12, 4, 5, 19�ڶ�Сֵ��" & Application.Small(arr1, 2)
    
    End Sub
    
 ''2�����

    ''��application.Sum (����)
    
''3 ͳ�Ƹ���

  ''counta��count��������ͳ��VBA��������ָ�����������������ݵĸ���
  
     Sub s1()
     
     Dim arr1, arr2(0 To 10), x
     arr1 = Array("a", "3", "", 4, 6)
     For x = 0 To 4
       arr2(x) = arr1(x)
     Next x
     
     MsgBox "����1�����ָ�����" & Application.Count(arr2)
     
     MsgBox "����2���������ֵ�ĸ���" & Application.CountA(arr2)
     
     End Sub
 
 ''3 �����������
     
     Sub s2()
      Dim arr
      On Error Resume Next
      arr = Array("a", "c", "b", "f", "d")
      MsgBox Application.Match("f", arr, 0)
     If Err.Number = 13 Then
        MsgBox "���Ҳ���"
      End If
     End Sub

Option Explicit

'' 1��split����
     ''���ָ������ַ�����ȡ��VBA����,��������һά���飬��Ŵ�0��ʼ
 
     '''split(�ַ���,�ָ���)
   
    Sub t1()
      Dim sr, arr
      sr = "A-BC-FGR-H"
      arr = VBA.Split(sr, "-")
      MsgBox Join(arr, ",")
    End Sub
     
'' 2��Filter������
     ''������ɸѡ����������ֵ���һ���µ�����

     ''Filter(����,ɸѡ����,��/��)
     
     ''ע������ǣ�true���򷵻ذ��������飬������򷵻طǰ���������
    Sub t2()
     Dim arr, arr1, arr2
     arr = Application.Transpose(Range("A2:A10"))
     arr1 = VBA.Filter(arr, "W", True)
     arr2 = VBA.Filter(arr, "W", False)
     Range("B2").Resize(UBound(arr1) + 1) = Application.Transpose(arr1)
     Range("C2").Resize(UBound(arr2) + 1) = Application.Transpose(arr2)
    End Sub
    
''3��index������
    ''���øù����������԰Ѷ�ά�����ĳһ�л�ĳһ�н�ȡ����������һ���µ����顣
     '' Application.Index(��ά����,0,����)) ���ض�ά����
     '' Application.Index(��ά����,����,0)) ����һά����
    Sub t3()
     Dim arr, arr1, arr2
      arr = Range("a2:d6")
      arr1 = Application.Index(arr, , 1)
      arr2 = Application.Index(arr, 4, 0)
      Stop
    End Sub

''4��vlookup����
      ''Vlookup�����ĵ�һ������������VBA���飬���ص�Ҳ��һ��VBA����
    Sub t4()
    Dim arr, arr1
      arr = Range("a2:d6")
      arr1 = Application.VLookup(Array("B", "C"), arr, 4, 0)
    End Sub
''5 Sumif������Countif����
     ''Countif��sumif�����ĵڶ�������������ʹ�����飬����Ҳ���Է���һ��VBA���飬�磺
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



 

 
''��24��

Option Explicit

Sub ��Ԫ��ѭ��()
 Dim x As Integer
 Dim t
 �����ɫ
 t = Timer
 For x = 2 To Range("a65536").End(xlUp).Row
   If Range("d" & x) > 500 Then
     Range(Cells(x, 1), Cells(x, 4)).Interior.ColorIndex = 3
   End If
 Next x
 MsgBox Timer - t
End Sub

Sub �����ɫ()
 Range("a:d").Interior.ColorIndex = xlNone
End Sub

Sub ���鷽��()
 Dim arr, t
 Dim x As Integer
 Dim sr As String, sr1 As String
 �����ɫ
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
Sub ���鷽��2()
Dim arr, t
 Dim x As Integer, x1 As Integer
 Dim sr As String, sr1 As String
 �����ɫ
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
Sub ���鷽��3()
Dim arr, t
 Dim x As Integer, x1 As Integer
 Dim sr As String, sr1 As String
 �����ɫ
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
''����Ҳ�������ø�ʽ��
   ''����������������⣬��Ȼû����ɫ������ȸ�ʽ�����Ǳ�����range������Ա�ʾ��������������ĵ�Ԫ������
   ''���������ص㣬���Ǿ���Ҫ���鹹�쵥Ԫ���ַ����Ȼ�������Ե�Ԫ����и�ʽ���á�
   ''ע�⣬��Ԫ���ַ������>255�����������Ԫ��������࣬���ǻ���Ҫ�ִη������õ�Ԫ���ʽ
   
Sub �����ɫ()
 Range("a2:d2,a7:d7,a10:d10").Interior.ColorIndex = 3
End Sub




����25��
Option Explicit

Sub ð������()
 Dim arr, temp, x, y, t, k
 t = Timer
 arr = Range("a1:a10")
 For x = 1 To UBound(arr) - 1
   For y = x + 1 To UBound(arr) ''ֻ�͵�ǰ��������������бȽ�
     If arr(x, 1) > arr(y, 1) Then ''���������������ĳһ������
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


Sub ð��������ʾ()
 Dim arr, temp, x, y, t, k
 For x = 1 To 9
                         Range("a" & x).Interior.ColorIndex = 3
   For y = x + 1 To 10  ''ֻ�͵�ǰ��������������бȽ�
                         Range("a" & y).Interior.ColorIndex = 4
     If Cells(x, 1) > Cells(y, 1) Then ''���������������ĳһ������
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


Sub ѡ������()
  Dim arr, temp, x, y, t, iMax, k, k1, k2
  t = Timer
  arr = Range("a1:a10")
  For x = UBound(arr) To 1 + 1 Step -1
     iMax = 1 ''��������
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

Sub ѡ������Ԫ����ʾ()
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
    
''��26��

Option Explicit


Sub ��������()
Dim arr, temp, x, y, t, iMax, k, k1, k2
  t = Timer
  arr = Range("a1:a10")
  For x = 2 To UBound(arr)
  
     temp = arr(x, 1) ''�ǵ�Ҫ�����ֵ
     
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

Sub ��������Ԫ����ʾ()
On Error Resume Next
  Dim arr, temp, x, y, t, iMax, k
  For x = 2 To 10
  
     temp = Cells(x, 1) ''�ǵ�Ҫ�����ֵ
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

    

    ''��ֻ��һ��ֵ��������

    If (iUBound - iLBound) Then

        For iOuter = iLBound To iUBound

            If lngArray(iOuter) > lngArray(iMax) Then iMax = iOuter

        Next iOuter

        

        iTemp = lngArray(iMax)

        lngArray(iMax) = lngArray(iUBound)

        lngArray(iUBound) = iTemp

    

        ''��ʼ��������

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

        

        ''����ֵ

        iTemp = lngArray(iLeftCur)

        lngArray(iLeftCur) = lngArray(iRightCur)

        lngArray(iRightCur) = iTemp

    Loop

    

    ''�ݹ��������

    lngArray(iLeftEnd) = lngArray(iRightCur)

    lngArray(iRightCur) = iPivot

    

    InnerQuickSort lngArray, iLeftEnd, iRightCur - 1

    InnerQuickSort lngArray, iRightCur + 1, iRightEnd

End Sub






Sub ϣ������()
  Dim arr
  Dim �ܴ�С, ���, x, y, temp, t
  t = Timer
  arr = Range("a1:a30")
  �ܴ�С = UBound(arr) - LBound(arr) + 1
  ��� = 1
  If �ܴ�С > 13 Then
     Do While ��� < �ܴ�С
       ��� = ��� * 3 + 1
     Loop
     ��� = ��� \ 9
  End If
''  Stop
  Do While ���
     For x = LBound(arr) + ��� To UBound(arr)
      temp = arr(x, 1)
      For y = x - ��� To LBound(arr) Step -���
         If arr(y, 1) <= temp Then Exit For
         arr(y + ���, 1) = arr(y, 1)
        '' k1 = k1 + 1
      Next y
      arr(y + ���, 1) = temp
     Next x
    ��� = ��� \ 3
   Loop
  '' MsgBox k1
   ''Range("e3").Resize(5000) = ""
    Range("d1").Resize(UBound(arr)) = arr
   ''Range("e2") = Timer - t
End Sub
Sub ����˳��()
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
Sub ϣ������Ԫ����ʾ()
  Dim arr
  Dim �ܴ�С, ���, x, y, temp, t
  t = Timer
  arr = Range("a1:a" & Range("a65536").End(xlUp).Row)
  �ܴ�С = UBound(arr) - LBound(arr) + 1
  ��� = 1
  If �ܴ�С > 13 Then
     Do While ��� < �ܴ�С
       ��� = ��� * 3 + 1
     Loop
     ��� = ��� \ 9
  End If
''  Stop
  Do While ���
     For x = LBound(arr) + ��� To UBound(arr)
      temp = Cells(x, 1)
      Range("a" & x).Interior.ColorIndex = 3
      For y = x - ��� To LBound(arr) Step -���
          Range("a" & y).Interior.ColorIndex = 6
         If Cells(y, 1) <= temp Then Exit For
         Cells(y + ���, 1) = Cells(y, 1)
        '' k1 = k1 + 1
      Next y
      Cells(y + ���, 1) = temp
      Range("a1:a30").Interior.ColorIndex = xlNone
     Next x
    ��� = ��� \ 3
   Loop
  '' MsgBox k1
   ''Range("e3").Resize(5000) = ""
   '' Range("d1").Resize(UBound(arr)) = arr
   ''Range("e2") = Timer - t
End Sub


Option Explicit

Sub ����������֮ð�ݷ�()
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
Sheets("��6").Activate
Application.ScreenUpdating = True

End Sub
Sub ����������֮ѡ��()
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
Sheets("��6").Activate
Application.ScreenUpdating = True
End Sub
Sub ����������֮���뷨()
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
Sheets("��6").Activate
Application.ScreenUpdating = True
End Sub

Sub ����˳��()
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
Sheets("��6").Activate
Application.ScreenUpdating = True
End Sub

Sub ϣ������()

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
Sheets("��6").Activate
Application.ScreenUpdating = True

End Sub


Option Explicit

Sub ����������֮ϣ������()
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
Sheets("��6").Activate
Application.ScreenUpdating = True

End Sub

''��27��

Option Explicit

''1 ʲô��VBA�ֵ䣿
   ''�ֵ䣨dictionary����һ���������ݵ�С�ֿ⡣�������С�
      ''��һ�н�key , ���������ظ���Ԫ�ء�
      ''�ڶ�����item,ÿһ��key��Ӧһ��item����������Ϊ�ظ�
            ''Key   item
             ''A     10
             ''B     20
             ''C     30
             ''Z     10

''2 ��Ȼ�����飬Ϊʲô��Ҫѧ�ֵ䣿
   ''ԭ��:���٣����������
      ''1) A��ֻ��װ����ظ���Ԫ�أ���������ص���Ժܷ������ȡ���ظ���ֵ
      ''2) ÿһ��key��Ӧһ��Ψһ��item��ֻҪָ��key��ֵ���Ϳ������Ϸ������Ӧ��item�������ֵ����ʵ�ֿ��ٵĲ���

''3 �ֵ���ʲô���ޣ�
    ''�ֵ�ֻ�����У����Ҫ������е����ݣ�����Ҫͨ���ַ�������ϺͲ����ʵ�֡�
    ''�ֵ���û�ķ�һ��ʱ�䣬����������������ֵ�����ƾ��޷����ֳ�����
    
''4 �ֵ��������δ����ֵ䣿
    
    ''�ֵ�����scrrun.dll���ӿ��ṩ�ģ�Ҫ�����ֵ������ַ���
      ''��һ�ַ�����ֱ�Ӵ�����
        '''set d = CreateObject("scripting.dictionary")
      ''�ڶ��ַ��������÷�
        ''����-����-���-�ҵ�scrrun.dll-ȷ��

		
Option Explicit
 
 ''1 װ������
    Sub t1()
      Dim d As New Dictionary
      Dim x As Integer
      For x = 2 To 4
       d.Add Cells(x, 1).Value, Cells(x, 2).Value
      Next x
      MsgBox d.Keys(1)
      '''stop
    End Sub
 ''2 ��ȡ����
    Sub t2()
      Dim d
      Dim arr
      Dim x As Integer
      Set d = CreateObject("scripting.dictionary")
      For x = 2 To 4
       d.Add Cells(x, 1).Value, Cells(x, 2).Value
      Next x
      ''MsgBox d("����")
      ''MsgBox d.Keys(2)
      Range("d1").Resize(d.Count) = Application.Transpose(d.Keys)
      Range("e1").Resize(d.Count) = Application.Transpose(d.Items)
      arr = d.Items
    End Sub

  ''3 �޸�����
    Sub t3()
      Dim d As New Dictionary
      Dim x As Integer
      For x = 2 To 4
       d.Add Cells(x, 1).Value, Cells(x, 2).Value
      Next x
      d("����") = 78
      MsgBox d("����")
      d("����") = 100
      MsgBox d("����")
    End Sub

  ''4 ɾ������
    Sub t4()
      Dim d As New Dictionary
      Dim x As Integer
      For x = 2 To 4
        d(Cells(x, 1).Value) = Cells(x, 2).Value
      Next x
       d.Remove "����"
     '' MsgBox d.Exists("����")
      d.RemoveAll
      MsgBox d.Count
    End Sub
 
''���ִ�Сд
    Sub t5()
      Dim d As New Dictionary
      Dim x
      For x = 1 To 5
        d(Cells(x, 1).Value) = ""
      Next x
      Stop
    End Sub

''��28��


	Option Explicit

Sub ���˫�����()
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
 MsgBox d("����")
End Sub


Option Explicit

Sub ����()
 Dim d As New Dictionary
 Dim arr, x
 arr = Range("a2:b10")
 For x = 1 To UBound(arr)
   d(arr(x, 1)) = d(arr(x, 1)) + arr(x, 2) ''key��Ӧ��item��ֵ��ԭ���Ļ����ϼ��µ�
 Next x
 Range("d2").Resize(d.Count) = Application.Transpose(d.Keys)
 Range("e2").Resize(d.Count) = Application.Transpose(d.Items)
End Sub

Option Explicit

Sub ��ȡ���ظ��Ĳ�Ʒ()
 Dim d As New Dictionary
 Dim arr, x
 arr = Range("a2:a12")
 For x = 1 To UBound(arr)
      d(arr(x, 1)) = ""
 Next x
 Range("c2").Resize(d.Count) = Application.Transpose(d.Keys)
End Sub

''��29��

Option Explicit

Sub ���巨֮���л���()
 Dim ����(1 To 10000, 1 To 3)
 Dim ����
 Dim arr, x, k
 Dim d As New Dictionary
 arr = Range("a2:c" & Range("a65536").End(xlUp).Row)
 For x = 1 To UBound(arr)
   If d.Exists(arr(x, 1)) Then
      ���� = d(arr(x, 1))
      ����(����, 2) = ����(����, 2) + arr(x, 2)
      ����(����, 3) = ����(����, 3) + arr(x, 3)
   Else
      k = k + 1
      d(arr(x, 1)) = k
      ����(k, 1) = arr(x, 1)
      ����(k, 2) = arr(x, 2)
      ����(k, 3) = arr(x, 3)
   End If
 Next x
 Range("f2").Resize(k, 3) = ����
End Sub


Option Explicit

Sub ���巨֮���������л���()
 Dim ����(1 To 10000, 1 To 4)
 Dim ����
 Dim arr, x As Integer, sr As String, k As Integer
 Dim d As New Dictionary
 arr = Range("a2:d" & Range("a65536").End(xlUp).Row)
 For x = 1 To UBound(arr)
    sr = arr(x, 1) & "-" & arr(x, 2)
    If d.Exists(sr) Then
      ���� = d(sr)
      ����(����, 3) = ����(����, 3) + arr(x, 3)
      ����(����, 4) = ����(����, 4) + arr(x, 4)
    Else
      k = k + 1
      d(sr) = k
      ����(k, 1) = arr(x, 1)
      ����(k, 2) = arr(x, 2)
      ����(k, 3) = arr(x, 3)
      ����(k, 4) = arr(x, 4)
    End If
 Next x
   Range("g2").Resize(k, 4) = ����
End Sub


Option Explicit

Sub ���巨֮����͸�ӱ�ʽ����()
 Dim d As New Dictionary
 Dim ����(1 To 10000, 1 To 7)
 Dim ����, ����
 Dim arr, x, k
 
 arr = Range("a2:c" & Range("a65536").End(xlUp).Row)
 
 For x = 1 To UBound(arr)
   ���� = (InStr("1��2��3��4��5��6��", arr(x, 2)) + 1) / 2 + 1
   If d.Exists(arr(x, 1)) Then
      ���� = d(arr(x, 1))
      
      ����(����, ����) = ����(����, ����) + arr(x, 3)
   Else
      k = k + 1
      d(arr(x, 1)) = k
      ����(k, 1) = arr(x, 1)
      ����(k, ����) = arr(x, 3)
   End If
 Next x
 
 Range("f2").Resize(k, 7) = ����

End Sub


''��30��

Option Explicit

''1 ȡ�ù������ܸ������Զ��庯��

Function shcount()

 shcount = Sheets.Count
 
End Function
Sub dd()
 MsgBox getv(Range("a7"))
End Sub


''2 ȡ�õ�Ԫ����ʾֵ���Զ��庯��
 
  Function getv(rg As Range)
  
    getv = rg.Text
    
  End Function
 
''3 ��ȡ�ַ����ĺ���
 
 Function jiequ(sr As String, fh As String, wz As Integer)
    
    Dim Arr
    Arr = Split(sr, fh)
    jiequ = Arr(wz - 1)
    
 End Function
  
''4 ��ȡ���ظ�ֵ�ĸ���

  Function ���ظ�����(rg As Range)
   Dim d, Arr, ar
   Arr = rg
   Set d = CreateObject("scripting.dictionary")
   For Each ar In Arr
     d(ar) = ""
   Next ar
   ���ظ����� = d.Count
  End Function
 Sub test()
  
  MsgBox jiequ("A-BRT-C-EF", "-", 2)
  
 End Sub

 
 Option Explicit

''1 ʲô���Զ��庯����
  ''��VBA����VBA���������ǻ����Ե��ù��������������ܲ������ѱ�д�����أ����ԣ�����Ǳ����������Զ��庯��
  
''2 ��ô��д�Զ��庯����
 
   ''���ǿ��԰�����Ľṹ��д�Զ��庯��
  
    '' Function ��������(����1,����2....)
         
         ''����
         ''��������=���ص�ֵ������
         
    '' End Function
    
    
Option Explicit

''1 ��ô���Զ��庯�������й�������ʹ�ã�
  
   ''�� �Ѻ����Զ��庯�����ļ����Ϊ�ӽغ꣬Ȼ��ͨ������-�ӽغ�-����ҵ�����ļ�-ȷ����
   
''2 ��ô���Զ��庯�����˵��

    ''����-��-���������Զ��庯��������-ѡ��--��˵������д���������������
    
''3����ô���Զ��庯������

    Sub ����()
     Application.MacroOptions "���ظ�����", Category:=4
    End Sub
     
   ''ע:
         ''0 ��ȫ��
         ''1 ����
         ''2 ���ں�ʱ��
         ''3 ��ѧ������
         ''4 ͳ��
         ''5 ���Һ�����
         ''6 ���ݿ�
         ''7 �ı�
         ''8 �߼�
         ''9 ��Ϣ
     



''��32��

Option Explicit

''һ��ʲôMsgBox����
   ''�����Ե���һ�����ڣ���ʾ���趨�����ݡ����Ҵ������п�������ѡ��İ�ť�������ͬ�İ�ť�᷵�ز�ͬ����ֵ��
 ''��msgbox��Ϣ���ڿ�������һ������Ի��Ļ��ᣬ�Ը��߳�����һ��Ӧ����ô��
  
    Sub test1()
      MsgBox "��Һã�����msgbox����"
    End Sub

''���������﷨
   
   ''Msgbox (��������ʾ������,��ť��ͼʾ���,���ڱ���,��صİ����ļ�,�����ļ������ĵı��)
   
  
 Option Explicit

''��ť����
   ''��Ϣ�����ɰ�ť��ʾ,ͼ����ʾ,ȱʡ��ť���������⹦�����,��Щ���ܶ������������,�������ֻ��Ҫ��"+"��
   
  Sub test8()
    MsgBox "test", vbYesNoCancel + vbExclamation + vbDefaultButton2 + vbMsgBoxHelpButton ''��ʾȷ����ȡ����ť����ʾѯ��ͼ��
  End Sub
  Sub test9()
    MsgBox "mytest", vbExclamation + vbYesNo ''��ʾΣ��ͼ����Ƿ�ť
  End Sub
  Sub test10()
    MsgBox "���Դ���ṹ", vbYesNoCancel + vbMsgBoxHelpButton + vbCritical + vbDefaultButton3, "�����ĸ���ť�Ĵ���"
  End Sub
 Sub dd()
   MsgBox "dd", vbYesNo + vbExclamation + vbMsgBoxHelpButton
 End Sub


 Option Explicit

''1��������ʾ������
    
     ''1) ������ʾ:ֻ��Ҫ����һ����������һ���ַ����������ַ����ı��ʽ����
        
        ''��:
        Sub test2()
          MsgBox "���,��ӭ���ʹ��"
          MsgBox "���!,��ӭ��ʹ��" & ThisWorkbook.Name
        End Sub
      
      ''2) ������ʾ��
            ''chr(10) �������ɻ��з�
            ''chr(13) �������ɻس���
            ''vbcrlf ���з��ͻس���
            ''vbCr ��ͬ��chr(10)
            ''vblf ��ͬ��chr(13)
         ''����
         Sub test3()
           MsgBox "�Ұ�" & Chr(10) & "Excel��Ӣ��ѵ"
          '' MsgBox "�Ұ���" & Chr(13) & "Excel"
          '' MsgBox "����" & vbCrLf & "����ˮ��"

         End Sub
     
        ''3) �����ʾ
          ''chr(9) �Ʊ��
          Sub test4()
             MsgBox "����" & Chr(9) & "ְҵ" & Chr(10) & "����" & Chr(9) & "����ʦ" _
                     & Chr(10) & "����ΰ" & Chr(9) & "��ʦ"
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
         
          ''�ÿո������
            '' space(n) ���Բ���N���ո�
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
 
 ''2  �������ʾ����
    Sub test7()
      MsgBox "�˶Թ�ϵ������", , "ϵͳ��ʾ"
    End Sub
   

 Option Explicit

''Ҫ�����Ϣ����,����Ҫ�����ǵ������İ�ť���ܷ���һ��ֵ,���߳������ǵ����ĸ���ť.

 Sub test11()
  Dim k
  k = MsgBox("���Է���ֵ", vbYesNoCancel)
  MsgBox "�����˰�ť:" & Choose(k, "ȷ��", "ȡ��", "��ֹ", "����", "����", "��", "��")
 End Sub

 ''Ӧ��ʾ��
   Sub test12()
     If MsgBox("��ȷ��Ҫɾ����15����?", vbQuestion + vbYesNo, "ɾ����ʾ") = vbYes Then
       Rows(15).Delete
       MsgBox "ɾ���ɹ�"
     Else
       MsgBox "��ȡ����ɾ��"
     End If
   End Sub


Option Explicit
 
  ''Ҫ��Ӱ���,��Ҫ����msgbox �����ĵ��ĺ͵��������
    ''���ĸ������ǰ����ļ���·��,�����ļ�Ҫ����C:\WINDOWS\Help·����
    ''����������Ͱ����ļ������й�,��Ϊ��׼���Ĵ򿪰����ļ������õ������ı��,���û��������Ϊ0
  Sub test13()
  Dim x
  x = MsgBox("������Ӱ�����Ч��", vbOKCancel + vbMsgBoxHelpButton, "���԰���!", "D:/a.chm", 0) ''"C:\WINDOWS\Help\excel.chm", 0)
  End Sub



Option Explicit

''1 �Զ���ʱ�ر���Ϣ��,������������Ϣ�����

 Sub AA()
    Dim WshShell As Object
    Set WshShell = CreateObject("Wscript.Shell")
    WshShell.Popup "1���رգ�", 1, "��ʾ��", 16
End Sub
 
 
''��33��


Option Explicit

''1.inpubox����
  
  ''�﷨:
    ''inputbox(�������ʾ����,�������,Ĭ��ֵ,ˮƽλ��,��ֱλ��,�����ļ�,�����ļ�ID
    
''2.Application�����Inputbox����:��ʾһ�������û�����ĶԻ��򡣷��ش˶Ի������������Ϣ
  
   ''�﷨:
     ''Application.InputBox(�Ի�����ʾ����,��������,�ı�����Ĭ��ֵ,x����,y����,�����ļ�,�����ļ�������ID,�ı�������������)
  
  ''���һ��������ֵ˵��:
 ''     ֵ    ����
      ''0     ��ʽ
      ''1     ����
      ''2     �ı� (�ַ���)
      ''4     �߼�ֵ (True �� False)
      ''8     ��Ԫ�����ã���Ϊһ�� Range ����
      ''16    ����ֵ���� #N/A
      ''64    ��ֵ����

 ''ʲôʱ���÷���,ʲôʱ���ú���
       ''������Ĳ������Կ���inputbox�����ͷ����Ĳ�֮ͬ���Ƿ����Ⱥ������˺��漸���β���,���ֻ�Ǽ򵥵�����,�����÷���,
    ''�����Ҫ��Ӱ�����������������,����Application�����Inputbox����.

	
	
Option Explicit
  ''���һ��������ֵ˵��:
 ''     ֵ    ����
      ''0     ��ʽ
      ''1     ����
      ''2     �ı� (�ַ���)
      ''4     �߼�ֵ (True �� False)
      ''8     ��Ԫ�����ã���Ϊһ�� Range ����
      ''16    ����ֵ���� #N/A
      ''64    ��ֵ����
      
'' 1.���õ�Ԫ��
     ''inputbox��������������ֵΪ8��ʱ��,���������ѡ��Ԫ��ĵ�ַ.ʹ�ñ�����ʹ��SET�����Ķ������,�򷵻ص���һ����Ԫ�����,
  ''���򷴻ص������Ԫ�������ֵ,��VBA����.
   Sub text5()
     Dim rg As Range
     Set rg = Application.InputBox("��ѡ��Ԫ������", "ѡȡ��ʾ", , , , , , 8)
     MsgBox rg.Parent.Name & "!" & rg.Address
   End Sub
  
     Sub text6()
     Dim rg
      rg = Application.InputBox("��ѡ��Ԫ������", "ѡȡ��ʾ", , , , , , 8)
     MsgBox rg(2, 1)
   End Sub

 ''2 ��ʽ����
    ''�����һ����������Ϊ0ʱ,�������빫ʽ,���ص�Ҳ��һ����ʽ�ַ���,�����ʽ�к���Ԫ������,�����Զ�ת����rc���ø�ʽ(�Ե�ǰ���Ԫ��Ϊ����)
    
    Sub test7()
      Dim r
      r = Application.InputBox("�����빫ʽ", "������ʾ", , , , , , 0)
      MsgBox r
    End Sub

 ''3 �������뷵�ص���ֵ��ʽ
  Sub test8()
      Dim r
      r = Application.InputBox("�����빫ʽ", "������ʾ", , , , , , 1) ''��������������ʾ��Ч������
      MsgBox r
  End Sub
  Sub test9()
      Dim r
      r = Application.InputBox("�����빫ʽ", "������ʾ", , , , , , 2) ''���������ַ�,��Ȼ,����������Ҳ���ַ�
      MsgBox TypeName(r)
  End Sub
 ''4.��ֵ����
    ''����ѡȡ��Ԫ�������ֵ��Ϊ����,Ҳ���������Դ��д����ŵ�һά���ά����
  Sub test10()
      Dim r
      r = Application.InputBox("�����빫ʽ", "������ʾ", , , , , , 64) ''���������ַ�,��Ȼ,����������Ҳ���ַ�
      MsgBox r(2, 1)
  End Sub
 

 Option Explicit
''1 ��������ݷ��ظ�һ������
Sub test1()
  Dim sr
  sr = InputBox("�������", "����", 100)
    MsgBox sr
  sr = Application.InputBox("�������", "����", 100)
    MsgBox sr
End Sub

''2 ���������ֱ�ӵ�ȷ������ʲô
 
Sub test2()
  Dim sr
  sr = InputBox("�������", "����")
    MsgBox sr
  sr = Application.InputBox("�������", "����")
    MsgBox sr
End Sub

        ''�������Է��ֵ��������κ�����ֱ�ӵ�ȷ�����᷵�ؿ�,�������ǾͿ����ÿ����ж��Ƿ�����������

 Sub test3()
  Dim sr
  sr = InputBox("�������", "����")
  If sr = "" Then
    MsgBox "��û������͵���ȷ��"
  End If
    
  sr = Application.InputBox("�������", "����")
      If sr = "" Then
    MsgBox "��û������͵���ȷ��"
      ElseIf sr = "False" Then
      
  End If
End Sub

''3 ���ֱ�ӵ���"�˳�"��ť����ʲôֵ����
  
Sub test4()
  Dim sr
  sr = InputBox("�������", "����")
    MsgBox sr ''���ؿ�
  sr = Application.InputBox("�������", "����")
    MsgBox sr ''����False
End Sub

         ''������2,3���Կ���,�����Ҫ�ж��Ƿ����������ݺ��Ƿ������˳�,��Inpubox����ʱ�жϷ���ֵ�Ƿ�Ϊ�վͿ�����,
   '' �����Inputbox����,����Ҫ���������ж�.

   
   
 ''��34�� 
 Option Explicit

''һ FileDialog ������
 ''�ṩ�ļ��Ի��򣬹����� Microsoft Office Ӧ�ó����б�׼�ġ��򿪡��͡����桱�Ի������ơ�
 ''������Щ�Ի��򣬽���������û����Լ���ָ�����������Ӧ��ʹ�õ��ļ����ļ��С�

''
''���򿪡��Ի������û�ѡ��һ����������������Ӧ�ó�����ʹ�� Execute �����򿪵��ļ���
''�����Ϊ���Ի������û�ѡ��һ������ʹ�� Execute �������浱ǰ�ļ����ļ���
''���ļ�ѡȡ�����Ի������û�ѡ��һ�������ļ����û�ѡ����ļ�·�������� FileDialogSelectedItems ���ϡ�
''���ļ���ѡȡ�����Ի������û�ѡ��һ��·�����û�ѡ����ļ�·�������� FileDialogSelectedItems ���ϡ�

''�� ���Ժͷ���
  
   ''1 AllowMultiSelect ��������û����ļ��Ի�����ѡ�����ļ����򷵻� True��Boolean ���ͣ��ɶ�д
   ''2 SelectedItems ѡȡ�Ķ���ļ�����
   ''3 InitialFileName ����:���ó�ʼ·�����ļ�����
   ''4 InitialView ���� :�������ó�ʼ�ļ�����ʾ����
   ''5 show �����ж��û��Ƿ�����ȡ����ť,������ȡ���᷵��0,���򷵻�-1
   
    ''ѡ�񲢷���һ���ļ�����·��
      Sub f1()
        Dim f
        Dim dig As Object
        Set dig = Application.FileDialog(msoFileDialogOpen)
        With Application.FileDialog(msoFileDialogOpen)
           .AllowMultiSelect = True
           .Filters.Add "Excel�ļ�", "*.xls", 1
           .InitialFileName = ThisWorkbook.FullName ''"d:\"
           .InitialView = msoFileDialogViewDetails
           .Title = "�Ի������"
           .Show
           MsgBox .Show
          For Each f In .SelectedItems
            MsgBox f
          Next f
        End With
        Set dig = Nothing
      End Sub
   ''ѡ�񲢷����ļ���
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
     .Filters = "Excel���,*.xls"
     .InitialFileName = "����.xls"
     .FilterIndex = 1
     .Title = "����"
  End With
End Sub


Option Explicit

'' һ�� ���������﷨

   ''GetOpenFilename�൱��Excel�򿪴��ڣ�ͨ���ô���ѡ��Ҫ�򿪵��ļ��������Է���ѡ����ļ�����·�����ļ�����
     ''ע:�˷����������������ļ�?

  ''Application.GetOpenFilename(�ļ�����ɸѡ����,������ʾ�ڼ������͵��ļ�,����,�Ƿ�����ѡ�����ļ���)
  
     
  
 ''����ʾ��
   
      ''1 ������ֻ��excel�ļ�
      
        ''���ô�ĳ���ļ�����������Ĺ���
          
           ''"�ļ�����˵������,*.�ļ����ͺ��"
      Sub t1()
        Dim f
        f = Application.GetOpenFilename("Excel�ļ�,*.xls")
        MsgBox f
      End Sub
         
      ''2���򿪶����ļ�����(word��excel)
        
       ''�򿪶����ļ����ͣ�ֻ��Ҫ��","����������µ��ļ�����˵�����ļ����͡�
       
      Sub t2()
        Dim f
        f = Application.GetOpenFilename("Excel2003�ļ�,*.xls,Word�ļ�,*.doc")
        MsgBox f
      End Sub
  
      ''3 �򿪶����ļ�����,Ĭ����ʾword�ļ�
      
      Sub t3()
        Dim f
        f = Application.GetOpenFilename("Excel2003�ļ�,*.xls,Word�ļ�,*.doc,�ı��ļ�,*.txt", 2)
        MsgBox f
      End Sub
        
       ''4 ���öԻ�������
       
       Sub t4()
          Dim f
           f = Application.GetOpenFilename("Excel2003�ļ�,*.xls,Word�ļ�,*.doc,�ı��ļ�,*.txt", 2, "ѡ��Ҫ���ܵ��ļ�")
           MsgBox f
      End Sub
   
       ''5 ѡ�����ļ�,����������ʽ����
      Sub t5()
        Dim f
        ChDrive "E"
        ChDir Application.Path
        ''ChDir ".."
        f = Application.GetOpenFilename("Excel2003�ļ�,*.xls,Word�ļ�,*.doc,�ı��ļ�,*.txt", 1, MultiSelect:=True)
        MsgBox f(1)
      End Sub
 



     ''GetSaveAsFilename�﷨:
         
         '' GetSaveAsFilename(Ĭ����ʾ���ļ���,ɸѡ����,���ɸѡ����ʱ��ʾ�ڼ���,����)
         ''ע:�ô���Ҳ����ʵ���Եı������.ֻ��Ϊ�����ļ�����һ��;��

       Sub t1()
        Dim f
        f = Application.GetSaveAsFilename("ʾ��.xls", "excel���,*.xls", , "����ʾ��")
        MsgBox f
      End Sub

	  Option Explicit
  
   ''chdrive �̷� ���Ըı�Ĭ��������
   ''chdir  ·��  ���Ըı�Ĭ��·��
   
      Sub t6()
        Dim f
         ChDrive "E"
         ChDir ThisWorkbook.Path
        ''ChDir ".."
        f = Application.GetOpenFilename("Excel2003�ļ�,*.xls,Word�ļ�,*.doc,�ı��ļ�,*.txt", 1, MultiSelect:=True)
       '' MsgBox f(1)
      End Sub

	  
''��35��

Option Explicit
''�ַ�����ȡ

''left,right,mid,Len
Sub z1()
  Dim sr
  sr = "Excel��Ӣ��ѵ��"
  Debug.Print Left(sr, 5)
  Debug.Print Right(sr, 5)
  Debug.Print Mid(sr, 3, 5)
  Debug.Print Left(sr, Len(sr) - 1)
End Sub

'''split
 
Sub z2()
  Dim sr, arr
  sr = "Excel�ľ���Ӣ����ѵ��"
  arr = Split(sr, "��")
  Debug.Print UBound(arr)
End Sub


''val

 Sub z3()
  Dim sr
  sr = "89.90��Ԫ"
  Debug.Print Val(sr)
 End Sub

''�ַ������
 ''&
 Sub a4()
  Debug.Print "a" & "b"
 End Sub
 ''join
  
 Sub a5()
  Dim sr, arr
  sr = "Excel-��Ӣ-��ѵ��"
  arr = Split(sr, "-")
  Debug.Print Join(arr, "+")
End Sub


Option Explicit

''instr ��ǰ����

Sub c1()
  Dim sr
  sr = "Excel��Ӣ��ѵ"
  Debug.Print InStr(sr, "��Ӣ") > 0
End Sub

''InStrRev �Ӻ���ǰ

Sub c2()
  Dim sr
  sr = "Excel��Ӣ��ѵ��ѵ��̳"
  Debug.Print InStr(sr, "��")
End Sub
''Replace�滻

Sub c5()
 Dim sr
  sr = "Excel��Ӣ��ѵ��"
  sr = Replace(sr, "��ѵ��", "��̳")
  Debug.Print sr
End Sub

''mid����滻

Sub c6()
 Dim sr
  sr = "Excel��Ӣ��ѵ��"
  Mid(sr, 8, 2) = "��̳"
  Debug.Print sr
End Sub


Option Explicit

''LCase ת����Сд

Sub z1()
  Debug.Print LCase("ABC")
End Sub

''UCcae ת���ɴ�д

Sub z2()

  Debug.Print UCase("Abc")
  
End Sub

'''strConv ����

''���� ֵ ˵��
''vbUpperCase 1 ���ַ�������ת�ɴ�д��
''vbLowerCase 2 ���ַ�������ת��Сд��
''vbProperCase 3 ���ַ�����ÿ���ֵĿ�ͷ��ĸת�ɴ�д
Sub ת��()

  Debug.Print VBA.StrConv("wHo ARE you?", vbProperCase)
  
End Sub

Sub ת��2()
 Dim i As Long
Dim x() As Byte
x = StrConv("ABCDEFG", vbFromUnicode)    '' ת���ַ�����
Debug.Print Application.Min(x)
For i = 0 To UBound(x)
    Debug.Print x(i)
Next

End Sub


''TRimɾ�����˿ո�
''Ltrim ɾ����߿ո�
''Rtrim ɾ���ұ߿ո�
 Sub z3()
 Dim sr
 
 sr = " A B BC "
 Debug.Print Trim(sr)
 Debug.Print LTrim(sr)
 Debug.Print RTrim(sr)
 End Sub
 
''ASC ����һ�� Integer�������ַ���������ĸ���ַ�����,ANSI �ַ���
''CHr ���� String�����а�������ָ�����ַ�������ص��ַ�
Sub z4()
  Debug.Print Asc("Z")
  Debug.Print Chr(90)
End Sub

'''space �� string�����ظ����ַ�

 Sub z5()
 
    Debug.Print "A" & Space(10) & "B"
    Debug.Print "C" & String(10, "a") & "D"
    
 End Sub

''��36��

Option Explicit

''like "�Աȵ��ַ���"
''Option Compare Text
  '' �ַ���1 like �ַ���2
 Sub L1()
   Debug.Print "ABC" Like "ABc"
 End Sub

''ͨ���?
  ''�ж�BA�ǲ��ǳ���Ϊ2���ҵڶ����ַ�ΪA
 Sub L2()
   Debug.Print "BA" Like "?A"
 End Sub

''ͨ���*
     ''�ж��ַ������Ƿ����cel
 Sub L3()
   Debug.Print "Excel��Ӣ��ѵ" Like "*cel*"
 End Sub

''�жϺ�ͨ������ַ���

  ''��ͨ�������[]�ڣ��ʹ������ַ��ĶԱ�

 Sub l4()
   ''Debug.Print "QAB" Like "Q?B"
   Debug.Print "QaB" Like "Q?B"
   Debug.Print "Q?B" Like "Q[?]B"
   ''Debug.Print ""
 End Sub
 

''�ж���ָ��λ������
  ''�ж������Ƿ�Ϊ�����������ɵ�
 Sub l9()
    Debug.Print 5 Like "##"
 End Sub

''�ж���ĳ��������ַ�
 
  Sub L10()
   ''[��С-�����С2-��С3]
    ''Debug.Print "q" Like "[A-Za-z]"  '' �ж�q�ǲ�����ĸ
   '' Debug.Print "H" Like "[A-GM-Z]"  '' �ж�H�ǲ�����A-G��M-Z����
    Debug.Print 8 Like "[!2-9]"
  End Sub

''�жϷ���ĳ��������ַ�
   Sub L11()
   
     Debug.Print "A" Like "[!C-Z]"
     
   End Sub
   
''�ж����г����ַ���

   Sub L12()
   
      Debug.Print "M" Like "[!ABCDEUE]"
      
   End Sub
    
''�ж�A~C��ͷ��F~G��β
  
   Sub L13()
     
     Debug.Print "AEREM" Like "[A-C]*[L-P]"
     Debug.Print "A334M" Like "[A-C]###[L-P]"
     
   End Sub
 

 Option Explicit

Sub ���()
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


''��37��

Option Explicit

''һ ������ʽ

   ''������ʽ�Ǵ����ַ������ⲿ���ߣ������Ը������õ��ַ����Աȹ��򣬽����ַ����ĶԱȡ��滻�Ȳ�����
   
   ''������ʽ�����ã�
     ''1����ɸ��ӵ��ַ����ж�
     ''2�����ַ����ж�ʱ����������޶ȵıܿ�ѭ�����Ӷ��ﵽ�������Ч�ʵ�Ŀ�ġ�
   
''�� ʹ�÷���
   
   ''1�����÷�
   ''���VBE�༭���˵������� - ���ã�ѡȡ: Microsoft VBScript Regular Expressions 5.5,���ú��ڳ���ʼ������������
     ''Dim regex As New RegExp
     Sub t1()
       Dim reg As New RegExp
     End Sub
     
    ''2��ֱ��������
''     �������� (���ڰ�)
''     Dim regex As Object
''     Set regex = CreateObject("VBScript.RegExp") ''�����������

     Sub t2()
       Dim reg As Object
       Set reg = CreateObject("VBScript.RegExp")
     End Sub

 ''�� ��������
    
    ''1 Global����:
       ''���ֵΪtrue,������ȫ���ַ�
       ''���ֵΪFalse,����������1����ֹͣ
       ''1 ��:
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
       
    ''2 IgnoreCase ����
       ''������������ִ�Сд�ģ�ΪFalse��ȱʡֵ��True����
    
    ''3 Pattern ����
       '' һ���ַ�������������������ʽ��ȱʡֵΪ���ı���
    ''4 Multiline ����,�ַ����ǲ���ʹ���˶���,����Ƕ���,$������ÿһ�е����һ��
       
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
       
     ''5  Execute ����
         ''����һ�� MatchCollection ���󣬸ö������ÿ���ɹ�ƥ��� Match ����,
         ''���ص���Ϣ����:
           ''FirstIndex:��ʼλ��
           ''Length; ����
           ''Value:����
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
       
     ''6��Text����
        ''����һ������ֵ����ֵָʾ������ʽ�Ƿ����ַ����ɹ�ƥ�䡣��ʵ�����ж������ַ����Ƿ�ƥ��ɹ�
        Sub t7()
         Dim reg As New RegExp
         Dim sr
         sr = "BCR6EA"
         With reg
           .Global = True
           .Pattern = "\d+"
           If .test(sr) Then MsgBox "�ַ����к�������"
         End With
        End Sub
        
 

 
 Option Explicit

Function ��ȡ����(rg As String, k As Integer)

  Dim regx As New RegExp
  With regx
   .Global = True
   If k = 1 Then
   
    .Pattern = "\D"
    
   ElseIf k = 2 Then
   
    .Pattern = "\w"
    
   End If
   
   ��ȡ���� = .Replace(rg, "")
  End With

End Function


''��38��

Option Explicit
 ''������ʽ�ĺ��������öԱȵĹ���Ҳ��������Pattern���ԣ��������Щ��������ַ��������⣬�Ǿ����ض�����ķ��š�
 ''������ܵ���������ʽ�г��÷��ŵĵ�һ���֡�
 
''\��

  ''1.���ڲ�����д���ַ�ǰ��,�绻�з�(\r),�س���(\n),�Ʊ��(\t),\����(\\)
  
  ''2.���������������ַ���ǰ��,��ʾ������,"\$","\^","\."
  
  ''3.���ڿ���ƥ�����ַ���ǰ��
      
       ''\d 0~9������
       ''\w ����һ����ĸ�����ֻ��»��ߣ�Ҳ���� A~Z,a~z,0~9,_ ������һ��
       ''\s �����ո��Ʊ������ҳ���ȿհ��ַ�����������һ��
       
       ''���ϸ�Ϊ��дʱ,Ϊ�෴����˼,��\D ��ʾ����������
       
        Sub t1()
           Dim regx As New RegExp
           Dim sr
           sr = "AE45B646C"
           With regx
             .Global = True
             .Pattern = "\d" ''�ų�������
             Debug.Print .Replace(sr, "")
           End With
        End Sub
''.(��)

   ''����ƥ������з�����������ַ�

''+��
   ''+��ʾһ���ַ��������������ظ��ġ�
    
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
''{}��
  ''���������ظ�����
    ''1 {n} �ظ�n��
        Sub t16()
           Dim regx As New RegExp
           Dim sr
           sr = "A234CA7A67"
           With regx
            .Global = True
            .Pattern = "\d{5}" ''������������
            Debug.Print .Replace(sr, "")
           End With
           
         End Sub
   ''2  {m,n}��С�ظ�m��,����ظ�n��
     
        Sub t22()
           Dim regx As New RegExp
           Dim sr
           sr = "A234CA7A6789"
           With regx
            .Global = True
            .Pattern = "\d{4,5}" ''�����������ֻ�������������
            Debug.Print .Replace(sr, "")
           End With
         End Sub
    ''3 {m,} �����ظ�m��,�൱��+
         Sub t23()
           Dim regx As New RegExp
           Dim sr
           sr = "A2348t6CA7A67"
           With regx
            .Global = True
            .Pattern = "\d{2,}" ''�����������ֻ�������������
            Debug.Print .Replace(sr, "")
           End With
         End Sub
         
''* ���Գ���0�������   �൱�� {0,}�����磺"\^*b"����ƥ�� "b","^^^b"...

'' ?
  ''1 ƥ����ʽ0�λ���1�Σ��൱�� {0,1}�����磺"a[cd]?"����ƥ�� "a","ac","ad"

        Sub t24()
           Dim regx As New RegExp
           Dim sr
           sr = "A23.48CA7A6..7"
           With regx
            .Global = True
            .Pattern = "\d+\.?\d+" ''�������1��
            Debug.Print .Replace(sr, "")
           End With
         End Sub
    ''2 ����+?�ĸ�ʽ���Էֶ�ƥ��
          
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
     

 

''��39��
Option Explicit

''^���ţ����Ƶ��ַ�����ǰ��,��^\d��ʾ�����ֿ�ͷ
 
    Sub T34()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "d234��345d43"
        With regex
          .Global = True
          .Pattern = "^\d*"
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub

''$���ţ����Ƶ��ַ�������棬�� A$��ʾ���һ���ַ���A

   
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
  ''�ո�(������ͷ�ͽ�β)
  
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
 ''����������������,ƥ����߻��ұߵ�
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
''\un ƥ�� n������ n ������λʮ����������ʾ�� Unicode �ַ���
''����һ�ı�����4e00,���һ��������9fa5
       Sub t2722()
           Dim regx As New RegExp
           Dim sr
           sr = "A12d��A��56�� A4"
           With regx
            .Global = True
            .Pattern = "[\u4e00-\u9fa5]"
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub

Option Explicit

''()
  ''��������������Ϊһ����������ظ�
   
        Sub t29()
           Dim regx As New RegExp
           Dim sr
           sr = "A3A3QA3A37BDFE87A8"
           With regx
            .Global = True
            .Pattern = "((A3){2})" ''�൱��A3A3
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
  ''ȡƥ������ʱ�������еı��ʽ������ \��������

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
        
''��(?=�ַ�)�����Ƚ���Ԥ����ң���һ��ƥ����󣬽���ƥ���ı�֮ǰ��ʼ������һ��ƥ��� ���ᱣ��ƥ�����Ա�����֮�á�
  
  ''������ȡĳ���ַ�֮ǰ������
      Sub t343()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "100Ԫ8000Ԫ57Ԫ"
        With regex
          .Global = True
           .Pattern = "\d+(?=Ԫ)" ''������������ֺ��Ԫ�����ҵ����Ԫ��ǰ��ʼ���ң���ΪԪǰ�������ѱ�ʹ�ã�
                                  ''����ֻ�ܴ�Ԫ��ʼ���ң�ƥ�� ()����ģ���Ϊ����û�����ã�����ֻ��ʾǰ������֣�Ԫ������ʾ
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub
   ''������֤���룬������4-8λ���������һ������
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
      
''��(?!�ַ�)�����Ƚ��и�Ԥ����ң���һ��ƥ����󣬽���ƥ���ı�֮ǰ��ʼ������һ��ƥ��� ���ᱣ��ƥ�����Ա�����֮�á�
     Sub t356()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "�й��������Ź�˾"
        With regex
          .Global = True
           .Pattern = "^(?!�й�).*"
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub
 
''()��|һ��ʹ�ÿ��Ա�ʾor

      Sub t344()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "100Ԫ800��7Ԫ"
        With regex
          .Global = True
           .Pattern = "\d+(Ԫ|��)"
           ''.Pattern = "\d+(?=Ԫ|��)"
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub


Option Explicit

''[]
 ''ʹ�÷����� [ ] ����һϵ���ַ����ܹ�ƥ����������һ���ַ����� [^ ] ������һϵ���ַ���
 ''���ܹ�ƥ�������ַ�֮�������һ���ַ���ͬ���ĵ�����Ȼ����ƥ����������һ��������ֻ����һ�������Ƕ��
 
  ''1 �������ڵ�����һ��ƥ��
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
        
   ''2 �������ڵ��ַ�
        
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
   ''3 ��һ������
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

''��40��

Option Explicit

''()
  ''��������������Ϊһ������
   
        Sub t29()
           Dim regx As New RegExp
           Dim sr
           sr = "A3A3QA3A37BDFEA387A8"
           With regx
            .Global = True
            .Pattern = "(A3){2}" ''�൱��A3A3
            Debug.Print .Replace(sr, "")
           End With
           
        End Sub
  ''ȡƥ������ʱ�������еı��ʽ������ \��������

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
        
''��(?=�ַ�)�����Ƚ���Ԥ����ң���һ��ƥ����󣬽���ƥ���ı�֮ǰ��ʼ������һ��ƥ��� ���ᱣ��ƥ�����Ա�����֮�á�
  
  ''������ȡĳ���ַ�֮ǰ������
      Sub t343()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "100Ԫ8000Ԫ57Ԫ"
        With regex
          .Global = True
           .Pattern = "\d+(?=Ԫ)." ''������������ֺ��Ԫ�����ҵ����Ԫ��ǰ��ʼ����,���Һ�\dƥ��ġ�
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub
   ''������֤���룬������4-8λ���������һ������
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
      
''��(?!�ַ�)�����Ƚ��и�Ԥ����ң���һ��ƥ����󣬽���ƥ���ı�֮ǰ��ʼ������һ��ƥ��� ���ᱣ��ƥ�����Ա�����֮�á�
     Sub t356()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "�������Ź�˾"
        With regex
          .Global = True
           .Pattern = "^(?!�й�).*"
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub
 
''()��|һ��ʹ�ÿ��Ա�ʾor

      Sub t344()
        Dim regex As New RegExp
        Dim sr, mat, m
        sr = "100Ԫ800��7Ԫ"
        With regex
          .Global = True
          '' .Pattern = "\d+(Ԫ|��)"
           
           .Pattern = "\d+Ԫ|\d+��"
           
            Set mat = .Execute(sr)
            For Each m In mat
              Debug.Print m
            Next m
        End With
      End Sub

''��41��

Option Explicit

Sub ��ť1_����()
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
       Cells(x, 3) = "û��ƥ���"
       Else
       Cells(x, 3) = .Execute(sr)(0)
       End If
      End If
     End With
    Next x
End Sub

''��42��
Option Explicit
''1 ������������
   ''��VBA�е������������������ı�����������͡���Щ��ͬ�����������ض������ã��ڽ�������ʱҲ��ռ��
   ''��ͬ��С���ڴ棬���������ڱ�д����ʱΪ���������Ч�ʣ�һ�㶼Ҫ�������ݵ����͡�
   
''2 �������ͶԳ������е�Ӱ��
     ''byte                       ռ��1���ֽ�
     ''integer,boolean            ռ��2���ֽ�
     ''long,single                ռ��4���ֽ�
     ''Double,Currency,date       ռ��8���ֽ�
     ''object                     ռ��4���ֽ�
     '''string(������)             ռ��10+�ַ����ȸ��ֽ�
     '''string(����)               ռ���ַ������ȸ��ֽ�
     ''Variant(������������)      ռ��16���ֽ�
     ''Variant(�ַ���)            ռ��24+�ַ������ȸ��ֽ�
   Sub sss1()
      Dim x As Long
      Dim t
      ''Dim k1 As Byte     ''��ʱ0.03125s
      Dim k
      ''Dim k1 As Integer ''��ʱ0.15625s
      Dim k1 As String   ''��ʱ0.203125s
      k = 1
      t = Timer
      For x = 1 To 1000000
        k1 = k
      Next x
      Debug.Print Timer - t
    End Sub



''1 ����Ƿ�Ϊ��
   Sub s1()
     Debug.Print Range("a1") = "" ''�ж����,�޷��жϼٿ�
     Debug.Print Len(Range("a1")) = 0 ''�ж���գ��޷��жϼٿ�
     Debug.Print VBA.IsEmpty(Range("a1")) ''�ٿ�ʱ����FALSE
     Debug.Print VBA.TypeName(Range("a1").Value) ''����Empty��ʾΪ��
   End Sub
   
   Sub �ٶȲ���()
     Dim t
     Dim x As Long
     t = Timer
     For x = 1 To 100000
       ''If Range("a1") = "" Then ''��ʱ0.81
      '' If Len(Range("a1")) = 0 Then ''0.84
      '' If VBA.IsEmpty(Range("a1")) Then ''�ٶ� 0.79
       ''If VBA.TypeName(Range("a1").Value) = Empty Then ''0.84
       End If
     Next x
   Debug.Print Timer - t
   End Sub

''2 ����Ƿ�Ϊ����
   Sub s2()
    Debug.Print VBA.IsNumeric(Range("a1"))
    Debug.Print Application.WorksheetFunction.IsNumber(Range("A1"))
    Debug.Print VBA.TypeName(Range("A1").Value)
   '' Debug.Print Range("a1").Value Like "#" ''�ж�һλ����
   '' Debug.Print Range("a1") Like "*#*" ''�ж��Ƿ��������
   End Sub
      Sub �ٶȲ���2()
     Dim t
     Dim x As Long
     t = Timer
     For x = 1 To 100000
       ''If VBA.IsNumeric(Range("a1")) Then ''��ʱ0 0.79
       ''If Application.WorksheetFunction.IsNumber(Range("A1")) Then ''0.9218
       ''If VBA.TypeName(Range("A1").Value) = "Double" Then ''�ٶ� 0.84
       End If
     Next x
   Debug.Print Timer - t
   End Sub

''3 ����Ƿ�Ϊ�ı�
   Sub t3()
     Debug.Print Application.IsText(Range("a1"))
     Debug.Print "B" Like "[A-Za-z]" ''�ж��Ƿ�Ϊ��ĸ
     Debug.Print Len(Range("a1"))
     Debug.Print Range("a1") Like "*[һ����]*" ''�ж��ַ������Ƿ��������
   End Sub

''4 �жϽ���Ƿ�Ϊ����ֵ
  Sub s4()
    Debug.Print VBA.IsError(Range("a1"))
    Debug.Print TypeName(Range("a1").Value)
  End Sub
  
''5 �ж��Ƿ�Ϊ����
   Sub s5()
     Dim arr
     arr = Range("A1:A2")
     Erase arr
     Debug.Print VBA.IsArray(arr)
   End Sub
''6 �ж��Ƿ�Ϊ����
   Sub s6()
      Debug.Print VBA.IsDate(Range("a2"))
   End Sub
   
Option Explicit

''һ������ת��������CBool, CByte, CCur, CDate, CDbl, CDec, CInt, CLng, CSng, CStr, CVar

''���������ǰѱ��ʽת�������Ӧ���������ͣ�����clngת���ɳ�����,cstrת�����ı���

Sub ss1()
 Dim s As Integer
 s = 2334
 MsgBox ��ȡ(CStr(s)) ''��Ϊ�Զ��庯������Ҫ�����ı����ͣ���s����ֵ���ͣ�������Ҫ��cstrת�����ı�����
End Sub

Function ��ȡ(x As String)
  ��ȡ = Left(x, 2)
End Function

Sub ss2()
 Debug.Print 1 + True ''CInt(1 = 1)
End Sub

''����Format����
 
  ''format�����÷���ͬ�ڹ������е�text���������Ը�ʽ����ʾ���ֻ��ı�
 
 Sub ss3()
  Dim n, n1
  n = 234.3372
  n1 = 41105
  Debug.Print Format(n, "0.00")
  Debug.Print Format(n, "0")
  Debug.Print Format(n, "\�۸�\:0.00")
  Debug.Print Format(n1, "yyyy-mm-dd")
 End Sub

''��43��

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

Sub ʱ����ʾ()
  Dim x
  If k = 1 Then
    k = 0
   End
  End If
  Range("a1") = Format(Now, "h:mm:ss")
  Application.OnTime Now + TimeValue("00:00:01"), "ʱ����ʾ"
  x = DoEvents
End Sub

Sub ����ʱ����ʾ()
 k = 1
End Sub

Option Explicit

''1 ���ص�ǰ���ڡ�ʱ�䣨ָ����ϵͳ���õ����ں�ʱ�䣩
  Sub t1()
    Debug.Print Date ''���ص�ǰ����
    Debug.Print Time ''���ص�ǰʱ��
    Debug.Print Now  ''���ص�ǰ����+ʱ��
  End Sub
  
''2 ��ʽ����ʾ����
   Sub t2()
     Debug.Print Format(Now, "yyyy-mm-dd")
     Debug.Print Format(Now, "yyyy��mm��dd��")
     Debug.Print Format(Now, "yyyy��mm��dd�� h:mm:ss")
     Debug.Print Format(Now, "d-mmm-yy") ''Ӣ���·�
     Debug.Print Format(Now, "d-mmmm-yy") ''Ӣ���·�
     Debug.Print Format(Now, "aaaa") ''��������
     Debug.Print Format(Now, "ddd") ''Ӣ������ǰ������ĸ
     Debug.Print Format(Now, "dddd") ''Ӣ������������ʾ
   End Sub
''3 ���������շ�������
   Sub t3()
     Debug.Print VBA.DateSerial(2011, 10, 1)
   End Sub
''4 ����Сʱ���ַ���ʱ��
   Sub t4()
     Debug.Print VBA.TimeSerial(1, 2, 1)
   End Sub

''5 ����������Сʱ����

  Sub t5()
  Dim d
    d = "2011-10-28 01:10:03"
    Debug.Print Year(d) & "��"
    Debug.Print Month(d) & "��"
    Debug.Print Day(d) & "��"
    Debug.Print Hour(d) & "ʱ"
    Debug.Print VBA.Minute(d) & "��"
    Debug.Print Second(d) & "��"
  End Sub


Option Explicit

''1 �������������������,����,����,Сʱ,����,��
  
   Sub tt1()
   Dim d1, d2 As Date
    d1 = #11/21/2011#
    d2 = #12/1/2011#
    Debug.Print "���" & (d2 - d1) & "��"
    Debug.Print "���" & DateDiff("d", d1, d2) & "��"
    Debug.Print "���" & DateDiff("m", d1, d2) & "��"
    Debug.Print "���" & DateDiff("yyyy", d1, d2) & "��"
    Debug.Print "���" & DateDiff("q", d1, d2) & "��"
    Debug.Print "���" & DateDiff("w", d1, d2) & "��"
    Debug.Print "���" & DateDiff("h", d1, d2) & "Сʱ"
    Debug.Print "���" & DateDiff("n", d1, d2) & "����"
    Debug.Print "���" & DateDiff("s", d1, d2) & "��"
   End Sub
   
    Sub tt2() ''������ʱ��Ĳ�
      Dim t, x
      t = Timer
      For x = 1 To 10000000
      Next x
      Debug.Print Timer - t
    End Sub

''2 ����ʱ��Ӽ�
   Sub tt3()
    Dim d1, d2 As Date
    d1 = "2001-10-1 00:00:00"
    Debug.Print VBA.DateAdd("d", 10, d1) ''����10��
    Debug.Print VBA.DateAdd("m", 10, d1) ''����10����
    Debug.Print VBA.DateAdd("yyyy", 10, d1) ''����10��
    Debug.Print VBA.DateAdd("yyyy", -10, d1) ''����10��
    Debug.Print VBA.DateAdd("h", 10, d1) ''����10Сʱ���ʱ��
    Debug.Print VBA.DateAdd("n", 10, d1) ''����10���ֺ��ʱ��
    Debug.Print VBA.DateAdd("s", 10, d1) ''����10����ʱ��
   End Sub


''��44��

Option Explicit

''1 ���ͼ��
     ''���shape����Ĵ��붼������¼�ƺ�ķ����õ�,������˽���ӵķ���,��ȥ¼��һ�����.
     ''����������Ӹ���ͼ�εĺ�
Sub Macro1()
''
'' Macro1 Macro
'' ���� Lenovo User ¼�ƣ�ʱ��: 2011-12-17
''
    ActiveSheet.Pictures.Insert ("D:\My Documents\My Pictures\��ɫ����ͷ��.jpg") ''����ͼƬ
    ActiveSheet.Shapes.AddLine(391.5, 214.5, 513.75, 273#).Select ''���ֱ��
    ActiveSheet.Shapes.AddShape msoShapeRectangle, 468#, 148.5, 94.5, 39.75 ''��Ӿ���
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 407.25, 308.25, _
        73.5, 90#).Select ''����ı���
    ActiveSheet.Buttons.Add(534.75, 241.5, 96.75, 41.25).Select ''��Ӵ����еĿؼ�
    ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False _
        , DisplayAsIcon:=False, Left:=603.75, Top:=160.5, Width:=83.25, Height _
        :=20.25).Select ''��ӿؼ�������ؼ�
End Sub

''2 ����ͼ�ε�λ��
      ''ͬ����������¼�ƺ����¼�Ƴ����롣
Sub Macro2()
''
'' Macro2 Macro
'' ���� Lenovo User ¼�ƣ�ʱ��: 2011-12-17
    ActiveSheet.Shapes("Picture 6").Select
    Selection.ShapeRange.IncrementLeft -27#  ''ˮƽ�����ƶ�
    Selection.ShapeRange.IncrementTop -51#   ''��ֱ�����ƶ�
    
End Sub



''Option Explicit
'''shapes ����
    ''�ö�����������ͼ�ι������ϵ�����ͼ��,����sheets��chart���Ӷ������ԣ�

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

Sub �������븴ѡ��()
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
        .Characters.Text = "��"
        .Value = xlOff
        .LinkedCell = RG.Address
    End With
 Next RG
End Sub



Option Explicit
Sub ����()
  Dim rg1 As Range, rg2 As Range, rg3 As Range, rg4 As Range
''ɾ����������
 Dim S As Shape
 Dim RG As Range
 For Each S In ActiveSheet.Shapes
   If S.Type = 9 Then
     S.Delete
   End If
 Next S
 ''����
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


Sub �ı���7_����()
  Application.Caller.Name
End Sub

Option Explicit

Sub ͼƬ����()
''ɾ������ͼƬ
 Dim S As Shape
 Dim RG As Range
 For Each S In ActiveSheet.Shapes
   If S.Type <> 8 Then
     S.Delete
   End If
 Next S
''����ͼƬ
   
  For Each RG In Range("b2:b5")
   '' Range("B2").Select
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, RG.Left, RG.Top, RG.Width, RG.Height).Select
    Selection.ShapeRange.Fill.UserPicture "E:\80����Ƶ\VBA��Ƶ80��\VBA80����44��\" & RG.Offset(0, -1) & ".jpg"
  Next RG
End Sub




''��45��
Option Explicit

Sub �����ѡ��ʾ����1()
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
Sub ���λ�λ��ʾ����()
  Dim arr
  Dim x As Integer, num As Integer, k As Integer, sr As String
  Range("c1:c10") = ""
  Range("a1:a10") = Application.Transpose(Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J"))
  For x = 1 To 10
     num = (Rnd() * ((10 - x + 1) - 1) + 1) \ 1
     Range("a1:a" & (10 - x + 1)).Interior.ColorIndex = xlNone
     Range("a" & num).Interior.ColorIndex = 6
     Range("c" & x) = Range("a" & num)
     ''���濪ʼ��λ
      sr = Range("a" & num)
      Range("a" & num) = Range("a" & (10 - x + 1))
      Range("a" & (10 - x + 1)) = sr
      Range("a" & (10 - x + 1)).Interior.ColorIndex = 1
  Next x
End Sub


Option Explicit

Sub �����ȡ�ֵ䷨()
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
''��������
   ''�ڻ�λʱ���ֵĻ�λ�ٶ�Ҫ���ı���Ҫ�졣���Խ�����ֵ������ﵽ���ٵ�Ŀ��
Sub �����������()
   Dim arr
   Dim arr1(1 To 20000, 1 To 1) As String, sr As String
   Dim x As Integer, num, t
   t = Timer
   arr = Range("a1:a20000")
   For x = 1 To UBound(arr)
      num = (Rnd() * ((20000 - x + 1) - 1) + 1) \ 1
      arr1(x, 1) = arr(num, 1)
      ''��λ
      sr = arr(num, 1)
      arr(num, 1) = arr(20000 - x + 1, 1)
      arr(20000 - x + 1, 1) = sr
   Next x
   Range("c1:c20000") = ""
   Range("c1:c20000") = arr1
   [d65536].End(xlUp).Offset(1, 0) = Timer - t
End Sub


Sub ���������������()
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
      ''��λ
      sr = arr2(num)
      arr2(num) = arr2(20000 - x + 1)
      arr2(20000 - x + 1) = num
   Next x
   Range("c1:c20000") = ""
   Range("c1:c20000") = arr1
    [F65536].End(xlUp).Offset(1, 0) = Timer - t
End Sub

''��46��

Option Explicit
Sub jin(n)
 If n < 4 Then
   jin n + 1
   jin n + 1
 End If
End Sub

Sub ����jin()
  jin 1
End Sub


Option Explicit
Dim k As Long
''�ݹ����

''1 ʲô�ǵݹ飿
  ''�ݹ�������ѵ������ѡ�
 ''2,�õݹ���ʲô�ô���
   ''�򻯴��룬�ó������ݡ��ر�����ѭ����������������¿��Դ��򵥴��롣
 ''3,�ݹ���ʲô������
    ''��Ϊ�ݹ���ʹ��ʱ���������������ʱ��Ϣ�ġ�ջ�������Ƚ��ȳ�������Ϣ������������Ч���Ƚϵͣ�����һ�㲻����ʹ�õݹ���Ƴ���
''2 ��:  ����4�Ľ׳� (4 * 3 * 2 * 1 = 24)
   
   Sub һ�㷽��()
     Dim k, x
     k = 1
     For x = 4 To 1 Step -1
        k = k * x
     Next x
     MsgBox k
   End Sub
   
   Sub �ݹ�1()
      MsgBox s(5)
   End Sub
''������
   Function s(n As Integer) As Integer
     If n = 1 Then
        s = 1
     Else
       s = n * s(n - 1)
     End If
   End Function
  Sub �ݹ�2()
    k = 1
    s2 4
    MsgBox k
  End Sub
'''sub���̷�
   Sub s2(n As Integer)
    '' Dim m
     If n > 0 Then
      k = k * n
     ''m = n
      s2 n - 1
     End If
   End Sub
   
''3 ��������1+2+3+.5
 Sub �ݹ�3()
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
Dim arr1(1 To 100, 1 To 1) ''�ѷ����Ľ������arr1��
Dim k As Integer ''��Ϊarr1���ʱ������
Sub ���()
  Dim arr
  k = 0
  Erase arr1
  arr = Range("a2:a" & Range("a65536").End(xlUp).Row)
  zuhe arr, 1, "", 0
  Range("c2").Resize(100) = ""
  Range("c2").Resize(k) = arr1
End Sub

Sub zuhe(arr, x, sr, y)
''arr ��Դ���鵼���ӹ���
''x �ݹ��������
'''sr ���ӵ��ַ���
''y ���ӵĴ���
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
Dim arr1(1 To 100, 1 To 1) ''�ѷ����Ľ������arr1��
Dim k As Integer ''��Ϊarr1���ʱ������
Sub ���()
  Dim arr
  k = 0
  Erase arr1
  arr = Range("a2:a" & Range("a65536").End(xlUp).Row)
  zuhe arr, 1, "", 0
  Range("c2").Resize(100) = ""
  Range("c2").Resize(k) = arr1
End Sub

Sub zuhe(arr, x, sr, y)
''arr ��Դ���鵼���ӹ���
''x �ݹ��������
'''sr ���ӵ��ַ���
''y ���ӵĴ���
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
Dim arr1(1 To 10000, 1 To 1)  As String  ''��ʽ���ʽ����arr1��
Dim k As Integer ''��Ϊarr1���ʱ������
Dim g As Integer, h As Integer
Dim arr
Dim k1
Sub ���()
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
  MsgBox "�ҵ� " & k & " ����! ����" & Format(Timer - t, "0.00") & "��"
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

Sub ѭ����()
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
Dim arr1(1 To 10000, 1 To 1)  As String  ''��ʽ���ʽ����arr1��
Dim k As Integer ''��Ϊarr1���ʱ������
Dim g As Integer, h As Integer
Dim arr
Dim k1
Sub ���()
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
  MsgBox "�ҵ� " & k & " ����! ����" & Format(Timer - t, "0.00") & "��"
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


''��48��


Option Explicit

'1 �ж��ļ����Ƿ����
   'dir�����ĵڶ���������vbdirectoryʱ���Է���·���µ�ָ���ļ����ļ��У�������Ϊ""�����ʾ�����ڡ�
  Sub w1()
    If Dir(ThisWorkbook.path & "\2011�걨��2", vbDirectory) = "" Then
       MsgBox "������"
    Else
       MsgBox "����"
    End If
  End Sub
  
'2 �½��ļ���
   'Mikdir�����Դ���һ���ļ���
    Sub w2()
      MkDir ThisWorkbook.path & "\Test"
    End Sub
   
'3 ɾ���ļ���
   
   'RmDir������ɾ��һ���ļ��У������Ҫʹ�� RmDir ��ɾ��һ�������ļ���Ŀ¼���ļ��У���ᷢ������
   '����ͼɾ��Ŀ¼���ļ���֮ǰ����ʹ�� Kill �����ɾ�������ļ���
   
    Sub w3()
    Kill ThisWorkbook.path & "\test\*"
      RmDir ThisWorkbook.path & "\test"
    End Sub
'4 �ļ���������
    Sub w4()
      Name ThisWorkbook.path & "\test" As ThisWorkbook.path & "\test2"
    End Sub
     
'5 �ļ����ƶ�
     'ͬ��ʹ��name���������Դﵽ�ƶ���Ч�����������ļ��е��ļ�һ���ƶ�
    
    Sub w5()
      Name ThisWorkbook.path & "\test2" As ThisWorkbook.path & "\2011�걨��\test100"
    End Sub
    
'6 �ļ��и���
        Sub CopyFile_fso()
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFolder ThisWorkbook.path & "\�����½��ļ���", ThisWorkbook.path & "\2011�걨��\"
        Set fso = Nothing
        End Sub
'7 ���ļ���
   'ʹ��shell����������������ļ���
    Sub w7()
      Shell "explorer.exe " & ThisWorkbook.path & "\2011�걨��", 1
    End Sub


Option Explicit

'����ָ���ļ����е��ļ�

 Sub �����ļ�()
  Dim Filename As String, mypath As String, k As Integer
  mypath = ThisWorkbook.path & "\2011�걨��\1��\A��˾\"
  Range("A1:A10") = ""
  Filename = Dir(mypath & "*.xls")
  Do
    k = k + 1
    Cells(k, 1) = Filename
    Filename = Dir
  Loop Until Filename = ""
 End Sub
 Sub �������ļ�()
  Dim Filename As String, mypath As String, k As Integer
  mypath = ThisWorkbook.path & "\2011�걨��\"
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


 

   