Dim Str As String
Dim a, b, C As Integer

Private Sub Command1_Click()
frmMain.Show
End Sub

Private Sub Form_Load()
Dim st1, st2, st3, st4, st5, st6, st7, st8, st9, st10, Zo, Zo1, Zo2 As String
  Open App.Path & "\cofing.x" For Input As #1
  Input #1, Zo
  Input #1, st1
  Input #1, st2
  Input #1, st3
  Input #1, st4
  Input #1, Zo1
  Input #1, st5
  Input #1, st6
  Input #1, st7
  Input #1, st8
  Input #1, Zo2
  Input #1, st9
  Input #1, st10
  Text8 = Mid(st1, 8, Len(st1) - 7)
  Text9 = Mid(st2, 8, Len(st2) - 7)
  Text10 = Mid(st3, 8, Len(st3) - 7)
  Text11 = Mid(st4, 8, Len(st4) - 7)
  Text12 = Mid(st5, 8, Len(st5) - 7)
  Text13 = Mid(st6, 8, Len(st6) - 7)
  Text14 = Mid(st7, 8, Len(st7) - 7)
  Text15 = Mid(st8, 8, Len(st8) - 7)
  Text16 = Mid(st9, 8, Len(st9) - 7)
  Text17 = Mid(st10, 8, Len(st10) - 7)
  Close
  
  Me.Picture = LoadPicture(App.Path & "\images\loading0.bmp")
  Me.Width = 5025
  Me.Height = 4245
  Text2.Visible = False
  
  Timer1.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label1.ForeColor = &H80000010
End Sub

Private Sub Label1_Click()
  End
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label1.ForeColor = &H80&
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label1.ForeColor = &HFF&
End Sub

Private Sub Label2_Click()
  Me.Picture = LoadPicture(App.Path & "\images\loading1.bmp")
  Text2.Visible = True
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Me.Picture = LoadPicture(App.Path & "\images\loading1.bmp")
  Text2.Visible = True
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Me.Picture = LoadPicture(App.Path & "\images\loading0.bmp")
End Sub

Private Sub Label5_Click()
  Text1.Enabled = False
  Text2.Enabled = False
  Label6.Visible = True
  
  Str = Inet1.OpenURL("http://www.anycen.com/api/student/getstudentbeta/?state=5&username=" & Text1.Text & "&key=e1152c54e099a972f1471a")
  Text3.Text = MD5_32Bits(Text2.Text)
  On Error GoTo none
  a = InStr(2, Str, "}")
  b = InStr(a + 1, Str, "}")
  Text4.Text = Mid(Str, 2, a - 2)
  Text5.Text = Mid(Str, a + 2, b - a - 2)
  
  If Text4.Text = Text3.Text Then
    frmMain.Show
    Me.Hide
  Else
    MsgBox "密码错误,请核对后再输入！", 64, "提示"
    Text2.Text = ""
    Text1.Enabled = True
    Text2.Enabled = True
    Label6.Visible = False
  End If
  
  Exit Sub
none:
  MsgBox "用户不存在,请核对后再输入！", 64, "提示"
  Text1.Text = ""
  Text2.Text = ""
  Text1.Enabled = True
  Text2.Enabled = True
  Label6.Visible = False
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Me.Picture = LoadPicture(App.Path & "\images\loading0.bmp")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      Text1.Enabled = False
      Text2.Enabled = False
      Label6.Visible = True
  
      Str = Inet1.OpenURL("http://www.anycen.com/api/student/getstudentbeta/?state=5&username=" & Text1.Text & "&key=e1152c54e099a972f1471a")
      Text3.Text = MD5_32Bits(Text2.Text)
      On Error GoTo none
      a = InStr(2, Str, "}")
      b = InStr(a + 1, Str, "}")
      Text4.Text = Mid(Str, 2, a - 2)
      Text5.Text = Mid(Str, a + 2, b - a - 2)
  
      If Text4.Text = Text3.Text Then
        frmMain.Show
        Me.Hide
      Else
        MsgBox "密码错误,请核对后再输入！", 64, "提示"
        Text2.Text = ""
        Text1.Enabled = True
        Text2.Enabled = True
        Label6.Visible = False
      End If
  
      Exit Sub
none:
      MsgBox "用户不存在,请核对后再输入！", 64, "提示"
      Text1.Text = ""
      Text2.Text = ""
      Text1.Enabled = True
      Text2.Enabled = True
      Label6.Visible = False
  End If
End Sub

Private Sub Text6_Change()
  frmMain.Label22.Caption = Text6.Text
End Sub

Private Sub Text7_Change()
  frmMain.Label26.Caption = Text7.Text
End Sub

Public Sub Timer1_Timer()
Dim str1 As String
Dim a1, a2 As Integer
  Text18.Text = Inet2.OpenURL(Text10.Text)
  str1 = Inet3.OpenURL(Text12.Text & Text13.Text & "/" & Text14.Text & "/?state=4&key=" & Text15.Text)
  a1 = InStr(2, str1, "}")
  Text6.Text = Mid(str1, 2, a1 - 2)
  a2 = InStr(a1 + 1, str1, "}")
  Text7.Text = Mid(str1, a1 + 2, a2 - a1 - 2)
  Timer1.Enabled = False
End Sub
