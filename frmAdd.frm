'Dim xlapp As New Excel.Application
'Option Explicit

Dim i, j As Long
Dim URL, Str, Succ, Bad, Sta As String                     'http://www.anycen.com/api/Student/GetStudentBeta/?state=1&addid=21406022038&addname=许智龙&note=000&key=e1152c54e099a972f1471a

Private Sub Form_Load()
'禁用:   DisableClose Me, True
'启用: DisableClose Me, False
DisableClose Me, True

i = 1
j = 1
Succ = "添加成功！"
Bad = "添加失败！"

Label1.Width = 0
Me.Caption = "正在计算长度..."

Timer1.Enabled = True

End Sub

Private Sub Label6_Click()
  Call frmLoad.Timer1_Timer
  Unload Me
End Sub

Private Sub Text2_Change()

Me.Caption = "正在批量添加学生信息..."

Dim xlApp As New Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
mypath = Text3.Text            '"D:\001.xls"
Set xlBook = xlApp.Workbooks.Open(mypath)
Set xlSheet = xlBook.Worksheets(1)

Do
DoEvents

x = xlSheet.Cells(i, 1)
Y = xlSheet.Cells(i, 2)

If x = "" Then
  Exit Do
End If

Text1.Text = "http://www.anycen.com/api/Student/GetStudentBeta/?state=1&addid=" & x & "&addname=" & Y & "&note=批量添加&key=e1152c54e099a972f1471a"

Str = Inet1.OpenURL(Text1.Text)

If Str = "{""status"":200,""message"":""ok"",""data"":{""information"":""success""}}" Then
  Sta = Succ
Else
  Sta = Bad
End If

List1.AddItem "学号：" & x & " - 姓名：" & Y & "   -   由管理员批量添加..................." & Sta
frmMain.List1.AddItem "学生：" & Y & " -   信息添加成功！"
i = i + 1
Label1.Width = ((i - 1) / j) * Picture1.Width

If Fix(((i - 1) / j) * 100) >= 50 Then
 Label3.ForeColor = &HFFFFFF
End If

Label3.Caption = Fix(((i - 1) / j) * 100) & "%"
Loop

xlBook.Close (True)
xlApp.Quit
Set xlApp = Nothing
Frame1.Visible = True
Command1.Enabled = True
DisableClose Me, False
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Dim xlApp As New Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
mypath = Text3.Text               '"D:\001.xls"
Set xlBook = xlApp.Workbooks.Open(mypath)
Set xlSheet = xlBook.Worksheets(1)

Do
DoEvents

x = xlSheet.Cells(j, 1)
Y = xlSheet.Cells(j, 2)

If x = "" Then
  Exit Do
End If

j = j + 1

Loop

xlBook.Close (True)
xlApp.Quit
Set xlApp = Nothing
Timer1.Enabled = False
j = j - 1
Label2.Caption = "本次共添加 " & j & " 名学生，具体时间视您的网络状况决定，请耐心等待。"
Text2.Text = j
End Sub

Private Sub Timer2_Timer()
  If List1.TopIndex < List1.ListCount - 1 Then
    List1.TopIndex = List1.TopIndex + 1
  Else
    'List1.TopIndex = List1.List(1)
  End If
End Sub
