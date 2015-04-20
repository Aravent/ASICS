Dim PrevX&, PrevY&
Dim Hang, i, s As Long
Dim URL, getStr, Credit, addStr As String

Private Sub ChangeOk_Click()
  If Text1.Text = "请仔细填写" Or Text1.Text = "" Or Text2.Text = "" Or Text2.Text = "请仔细填写" Or Text3.Text = "" Or Combo1.Text = "选择加减操作" Then
    MsgBox "请输入正确的内容！", 64, "提示"
  Else
    Label8.Caption = "学生学分信息修改 - 正在修改中..."
    Label6.Enabled = False
    ChangeOk.Enabled = False
    
    Do
    DoEvents
    
        If Combo1.Text = "增加学分" Then
          Credit = Val(Label19.Caption) + Val(Text3.Text)
        Else
          Credit = Val(Label19.Caption) - Val(Text3.Text)
          If Credit <= 0 Then
            MsgBox "已经为0"
            Credit = 0
          End If
        End If
        
        Label19.Caption = Credit
        
        If Val(Credit) <= Val(frmLoad.Text16.Text) Then
          Image5.Picture = Image7.Picture
        Else
          Image5.Picture = Image6.Picture
        End If
          
        Text6.Text = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=3&id=" & Label18.Caption & "&upcredit=" & Credit & "&upoption=管理员修改&upnote=" & Text2.Text & "&key=" & frmLoad.Text15.Text
          d = Inet1.OpenURL(Text6.Text)

    MsgBox "修改完成！", 64, "提示"
    Label8.Caption = "学生学分信息批量修改"
    Label6.Enabled = True
    ChangeOk.Enabled = True
    
    Exit Do
    Loop
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

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
  Combo1.AddItem "增加学分"
  Combo1.AddItem "减少学分"
  Label11.Width = 0
  s = 0
  Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label6.BackColor = &HFF8080
  Label9.BackColor = &HFF8080
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

Private Sub Label9_Click()
  Label9.Enabled = False
      If Len(Text1.Text) <> 11 Then
        MsgBox "请输入正确的学号", 64, "提示"
      Else
        'Text5.Text = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=0&id=" & Text1.Text & "&key=" & frmLoad.Text15.Text
        'MsgBox (UTF8EncodeURI(Text5.Text))
        'GetAPI = API.Conn(2, frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=0&id=" & Text1.Text & "&key=" & frmLoad.Text15.Text)
        addStr = Inet1.OpenURL(frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=0&id=" & Text1.Text & "&key=" & frmLoad.Text15.Text)
          'frmMain.Text7.Text = b
      
      If Len(addStr) = 37 Then
        MsgBox "没有找到学生信息，请核对后重试！", 64, "提示"
        Text1.Text = ""
      Else
          
          Text1.Text = Fun_GetStr(addStr, "id"":""", """,")      '提取信息
          Label18.Caption = Fun_GetStr(addStr, "id"":""", """,")
          Label17.Caption = Fun_GetStr(addStr, "name"":""", """,")
          Label19.Caption = Fun_GetStr(addStr, "credit"":""", """,")
         
      End If
      
      End If
  Label9.Enabled = True
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label9.BackColor = &HFFC0C0
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
  If Val(Label19.Caption) <= Val(frmLoad.Text16.Text) Then
    Image5.Picture = Image7.Picture
  Else
    Image5.Picture = Image6.Picture
  End If
End Sub
