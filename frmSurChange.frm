Dim PrevX&, PrevY&
Dim Hang, i, s As Long
Dim URL, getStr, Credit As String

Private Sub ChangeOk_Click()
  If Text1.Text = "请将要修改的学生学号填入，一行一个。" & vbCrLf Or Text1.Text = "" Or Text2.Text = "" Or Text2.Text = "请仔细填写" Or Text3.Text = "" Or Combo1.Text = "选择加减操作" Then
    MsgBox "请输入正确的内容！", 64, "提示"
  Else
    Label8.Caption = "学生学分信息批量修改 - 正在修改中..."
    Label6.Enabled = False
    ChangeOk.Enabled = False
  
    arr = Split(Text1.Text, vbCrLf)
    Hang = UBound(arr) + 1
  
    Dim Str() As String
    Str = Split(Text1.Text, vbCrLf)
    
    For i = 0 To UBound(Str)
      Text5.Text = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=0&id=" & Str(i) & "&key=" & frmLoad.Text15.Text
      getStr = Inet1.OpenURL(Text5.Text)
      s = s + 1
     
      If getStr = "{""status"":200,""message"":""ok"",""data"":}" Then
        MsgBox "学号：" & Str(i) & "存在错误，修改失败！", 64, "提示"
      Else
        Credit = Fun_GetStr(getStr, "credit"":""", """,")      'http://www.anycen.com/api/Student/GetStudentBeta/?state=3&id=21406022041&upcredit=0&upcredit=测试操作&upnote=备注&key=e1152c54e099a972f1471a
        If Combo1.Text = "增加学分" Then
          Credit = Val(Credit) + Val(Text3.Text)
        Else
          Credit = Val(Credit) - Val(Text3.Text)
          If Credit <= 0 Then
            MsgBox "已经为0"
            Credit = 0
          End If
        End If
          Text6.Text = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=3&id=" & Str(i) & "&upcredit=" & Credit & "&upoption=管理员批量修改&upnote=" & Text2.Text & "&key=" & frmLoad.Text15.Text
          d = Inet1.OpenURL(Text6.Text)
      End If
        Label11.Width = (s / Hang) * Me.Width
        'MsgBox Credit
    Next
    MsgBox "批量修改完成！", 64, "提示"
    Label8.Caption = "学生学分信息批量修改"
    Label6.Enabled = True
    ChangeOk.Enabled = True
  End If
End Sub

Private Sub ChangeOk_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  ChangeOk.BackColor = &HFFC0C0
End Sub

Private Sub Combo1_Click()
  Combo1.BackColor = &HC0FFFF
End Sub

Private Sub Command1_Click()
  Label11.Width = 0
  Timer1.Enabled = True
End Sub

Private Sub Form_Load()
  Combo1.AddItem "增加学分"
  Combo1.AddItem "减少学分"
  Label11.Width = 0
  s = 0
  Timer1.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label6.BackColor = &HFF8080
  ChangeOk.BackColor = &HFF8080
  Image1.Picture = Image2.Picture
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Image1.Picture = Image2.Picture
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  PrevX = x
  PrevY = Y
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 1 Then Move Left + x - PrevX, Top + Y - PrevY
  Image1.Picture = Image2.Picture
End Sub

Private Sub Label10_Click()
  Unload Me
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Image1.Picture = Image3.Picture
End Sub

Private Sub Label10_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Image1.Picture = Image4.Picture
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label6.BackColor = &HFFC0C0
End Sub

Private Sub Text1_Click()
  If Text1.Text = "请将要修改的学生学号填入，一行一个。" & vbCrLf Then
    Text1.BackColor = &HC0FFFF
    Text1.Text = ""
  End If
End Sub

Private Sub Text2_Click()
  If Text2.Text = "请仔细填写" Then
    Text2.ForeColor = &H80000008
    Text2.BackColor = &HC0FFFF
    Text2.Text = ""
  End If
End Sub

Private Sub Text3_Click()
  Text3.BackColor = &HC0FFFF
End Sub

Private Sub Timer1_Timer()
  If Label11.Width < Me.Width Then
    Label11.Width = Label11.Width + 50
  Else
    Timer1.Enabled = False
    MsgBox "0000"
  End If
End Sub
