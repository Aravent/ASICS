Option Explicit

Dim xCursor As Long, yCursor As Long
Dim GetAPI, URL As String
Dim i, j As Integer

Private Sub butgai_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  SearchCredit.BackColor = &HFF8080
End Sub

Private Sub ChangeOk_Click()
  Open App.Path & "\cofing.x" For Output As #1
  Print #1, "SOFT[]"
  Print #1, "UPDATE=" & frmLoad.Text8.Text
  Print #1, "VERSON=" & frmLoad.Text9.Text
  Print #1, "UPURLS=" & frmLoad.Text10.Text
  Print #1, "POWERS=" & frmLoad.Text11.Text
  Print #1, "API[]"
  Print #1, "SERURL=" & frmLoad.Text12.Text
  Print #1, "SERNAM=" & frmLoad.Text13.Text
  Print #1, "GETNAM=" & frmLoad.Text14.Text
  Print #1, "USEKEY=" & frmLoad.Text15.Text
  Print #1, "SETTING[]"
  Print #1, "ARMCID=" & Text8.Text
  Print #1, "USUCID="
  Close
  MsgBox "修改成功！", , "提示"
End Sub

Private Sub Command1_Click()
  Image13.Left = -Image13.Width + 1950
  Image12.Visible = True
  Image13.Visible = True
  Timer2.Enabled = True
End Sub

Public Sub Command2_Click()
  Picture6.Top = -500
  T1.Enabled = True
End Sub

Private Sub Form_Click()
  j = 1
  Timer1.Enabled = True
End Sub

Private Sub Form_Load()
  i = 0
  j = 0
  Me.Width = 14415
  Me.Height = 8025
  Text1.Width = 10
  Timer1.Enabled = False
  
  stro(0).Picture = Image2.Picture
  Home.Picture = index2.Picture
  
  stro(1).Picture = Image3.Picture
  Search.Picture = search3.Picture
  
  stro(2).Picture = Image3.Picture
  Add.Picture = Add3.Picture
  
  stro(3).Picture = Image3.Picture
  Seething.Picture = seething3.Picture
  
  Label14.Caption = "版本号:" & frmLoad.Text9.Text
  Label14.Left = Me.Width - Label14.Width - 150
  Label22.Caption = frmLoad.Text6.Text
  Label26.Caption = frmLoad.Text7.Text
  Text8.Text = frmLoad.Text16.Text
  Image13.Left = -Image13.Width + 1950
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Image1.Picture = Image4.Picture
End Sub

Private Sub Frame1_Click()
  j = 1
  Timer1.Enabled = True
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  SearchId.BackColor = &HFF8080
  SearchName.BackColor = &HFF8080
  SearchCredit.BackColor = &HFF8080
  Label66.BackColor = &HFF8080
  Label67.BackColor = &HFF8080
  Label68.BackColor = &HFF8080
End Sub

Private Sub Frame2_DblClick()
  Frame2.Visible = False
  butgai.Visible = False
  SearchCredit.Height = 315
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  SearchCredit.BackColor = &HFF8080
  Label45.BackColor = &HFF8080
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label58.BackColor = &HFF8080
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Image1.Picture = Image4.Picture
End Sub

Private Sub Image15_Click()

End Sub

Private Sub Label1_Click()
  i = 0
  j = 1
  Timer1.Enabled = True
  
  Frame1.Visible = False
  Frame3.Visible = False
  Frame5.Visible = False
  
  stro(0).Picture = Image2.Picture
  Home.Picture = index2.Picture
  stro(1).Picture = Image3.Picture
  Search.Picture = search3.Picture
  stro(2).Picture = Image3.Picture
  Add.Picture = Add3.Picture
  stro(3).Picture = Image3.Picture
  Seething.Picture = seething3.Picture
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  stro(0).Picture = Image2.Picture
  Home.Picture = index2.Picture
  If i = 0 Then
    stro(1).Picture = Image3.Picture
    Search.Picture = search3.Picture
    stro(2).Picture = Image3.Picture
    Add.Picture = Add3.Picture
    stro(3).Picture = Image3.Picture
    Seething.Picture = seething3.Picture
  ElseIf i = 1 Then
    stro(1).Picture = Image2.Picture
    Search.Picture = search2.Picture
    stro(2).Picture = Image3.Picture
    Add.Picture = Add3.Picture
    stro(3).Picture = Image3.Picture
    Seething.Picture = seething3.Picture
  ElseIf i = 2 Then
    stro(1).Picture = Image3.Picture
    Search.Picture = search3.Picture
    stro(2).Picture = Image2.Picture
    Add.Picture = Add2.Picture
    stro(3).Picture = Image3.Picture
    Seething.Picture = seething3.Picture
  ElseIf i = 3 Then
    stro(1).Picture = Image3.Picture
    Search.Picture = search3.Picture
    stro(2).Picture = Image3.Picture
    Add.Picture = Add3.Picture
    stro(3).Picture = Image2.Picture
    Seething.Picture = seething2.Picture
  End If
End Sub

Public Sub Label2_Click()
  i = 1
  j = 1
  Timer1.Enabled = True
  
  Frame1.Visible = True
  Frame3.Visible = False
  Frame5.Visible = False
  
  stro(1).Picture = Image2.Picture
  Search.Picture = search2.Picture
  stro(0).Picture = Image3.Picture
  Home.Picture = index3.Picture
  stro(2).Picture = Image3.Picture
  Add.Picture = Add3.Picture
  stro(3).Picture = Image3.Picture
  Seething.Picture = seething3.Picture
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  stro(1).Picture = Image2.Picture
  Search.Picture = search2.Picture
  If i = 0 Then
    stro(0).Picture = Image2.Picture
    Home.Picture = index2.Picture
    stro(2).Picture = Image3.Picture
    Add.Picture = Add3.Picture
    stro(3).Picture = Image3.Picture
    Seething.Picture = seething3.Picture
  ElseIf i = 1 Then
    stro(0).Picture = Image3.Picture
    Home.Picture = index3.Picture
    stro(2).Picture = Image3.Picture
    Add.Picture = Add3.Picture
    stro(3).Picture = Image3.Picture
    Seething.Picture = seething3.Picture
  ElseIf i = 2 Then
    stro(0).Picture = Image3.Picture
    Home.Picture = index3.Picture
    stro(2).Picture = Image2.Picture
    Add.Picture = Add2.Picture
    stro(3).Picture = Image3.Picture
    Seething.Picture = seething3.Picture
  ElseIf i = 3 Then
    stro(0).Picture = Image3.Picture
    Home.Picture = index3.Picture
    stro(2).Picture = Image3.Picture
    Add.Picture = Add3.Picture
    stro(3).Picture = Image2.Picture
    Seething.Picture = seething2.Picture
  End If
End Sub

Public Sub Label3_Click()
  i = 2
  j = 1
  Timer1.Enabled = True
  
  Frame3.Visible = True
  Frame1.Visible = False
  Frame5.Visible = False
  
  stro(2).Picture = Image2.Picture
  Add.Picture = Add2.Picture
  stro(0).Picture = Image3.Picture
  Home.Picture = index3.Picture
  stro(1).Picture = Image3.Picture
  Search.Picture = search3.Picture
  stro(3).Picture = Image3.Picture
  Seething.Picture = seething3.Picture
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  stro(2).Picture = Image2.Picture
  Add.Picture = Add2.Picture
  If i = 0 Then
    stro(0).Picture = Image2.Picture
    Home.Picture = index2.Picture
    stro(1).Picture = Image3.Picture
    Search.Picture = search3.Picture
    stro(3).Picture = Image3.Picture
    Seething.Picture = seething3.Picture
  ElseIf i = 1 Then
    stro(0).Picture = Image3.Picture
    Home.Picture = index3.Picture
    stro(1).Picture = Image2.Picture
    Search.Picture = search2.Picture
    stro(3).Picture = Image3.Picture
    Seething.Picture = seething3.Picture
  ElseIf i = 2 Then
    stro(0).Picture = Image3.Picture
    Home.Picture = index3.Picture
    stro(1).Picture = Image3.Picture
    Search.Picture = search3.Picture
    stro(3).Picture = Image3.Picture
    Seething.Picture = seething3.Picture
  ElseIf i = 3 Then
    stro(0).Picture = Image3.Picture
    Home.Picture = index3.Picture
    stro(1).Picture = Image3.Picture
    Search.Picture = search3.Picture
    stro(3).Picture = Image2.Picture
    Seething.Picture = seething2.Picture
  End If
End Sub

Private Sub Label4_Click()
  i = 3
  j = 1
  Timer1.Enabled = True
  
  Frame3.Visible = False
  Frame1.Visible = False
  Frame5.Visible = True
  
  stro(3).Picture = Image2.Picture
  Seething.Picture = seething2.Picture
  stro(0).Picture = Image3.Picture
  Home.Picture = index3.Picture
  stro(1).Picture = Image3.Picture
  Search.Picture = search3.Picture
  stro(2).Picture = Image3.Picture
  Add.Picture = Add3.Picture
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  stro(3).Picture = Image2.Picture
  Seething.Picture = seething2.Picture
  If i = 0 Then
    stro(0).Picture = Image2.Picture
    Home.Picture = index2.Picture
    stro(1).Picture = Image3.Picture
    Search.Picture = search3.Picture
    stro(2).Picture = Image3.Picture
    Add.Picture = Add3.Picture
  ElseIf i = 1 Then
    stro(0).Picture = Image3.Picture
    Home.Picture = index3.Picture
    stro(1).Picture = Image2.Picture
    Search.Picture = search2.Picture
    stro(2).Picture = Image3.Picture
    Add.Picture = Add3.Picture
  ElseIf i = 2 Then
    stro(0).Picture = Image3.Picture
    Home.Picture = index3.Picture
    stro(1).Picture = Image3.Picture
    Search.Picture = search3.Picture
    stro(2).Picture = Image2.Picture
    Add.Picture = Add2.Picture
  ElseIf i = 3 Then
    stro(0).Picture = Image3.Picture
    Home.Picture = index3.Picture
    stro(1).Picture = Image3.Picture
    Search.Picture = search3.Picture
    stro(2).Picture = Image3.Picture
    Add.Picture = Add3.Picture
  End If
End Sub

Private Sub Label41_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Frame2.Visible = False
  butgai.Visible = False
  SearchCredit.Height = 315
End Sub

Private Sub Label42_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Frame2.Visible = False
  butgai.Visible = False
  SearchCredit.Height = 315
End Sub

Private Sub Label43_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Frame2.Visible = False
  butgai.Visible = False
  SearchCredit.Height = 315
End Sub

Private Sub Label44_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Frame2.Visible = False
  butgai.Visible = False
  SearchCredit.Height = 315
End Sub

Private Sub Label47_Click()
  If IsNumeric(Credit1.Text) = True Then
    If IsNumeric(Credit2.Text) = True Then
      If Val(Credit1.Text) >= 0 Then
        If Val(Credit2.Text) >= 0 Then
          Image13.Left = -Image13.Width + 1950
          Image12.Visible = True
          Image13.Visible = True
          Timer2.Enabled = True
          URL = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=0&credit1=" & Credit1.Text & "&credit2=" & Credit2.Text & "&key=" & frmLoad.Text15.Text
          GetAPI = API.Conn(0, URL)
        Else
          MsgBox "输入格式错误，请重试！", 64, "提示"
        End If
      Else
        MsgBox "输入格式错误，请重试！", 64, "提示"
      End If
    Else
      MsgBox "输入格式错误，请重试！", 64, "提示"
    End If
  Else
    MsgBox "输入格式错误，请重试！", 64, "提示"
  End If
End Sub

Private Sub Label47_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label45.BackColor = &HFFC0C0
End Sub

Private Sub Label5_Click()
  Me.WindowState = 1
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Image1.Picture = Image6.Picture
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Image1.Picture = Image5.Picture
End Sub

Private Sub Label58_Click()
  If Text3.Text = "" Then
    Text3.BackColor = &HC0FFFF
  ElseIf Text4.Text = "" Then
    Text4.BackColor = &HC0FFFF
  Else
    If Text6.Text = "" Then
      Text6.Text = "初始学分"
    Else
      If Len(Text3.Text) <> 11 Then
        MsgBox "请输入正确的学号", 64, "提示"
      Else
        Image13.Left = -Image13.Width + 1950
        Image12.Visible = True
        Image13.Visible = True
        Timer2.Enabled = True
        URL = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=1&addid=" & Text3.Text & "&addname=" & Text4.Text & "&note=" & Text6.Text & "&key=" & frmLoad.Text15.Text
        GetAPI = API.Conn(1, URL)
        'Text7.Text = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=1&addid=" & Text3.Text & "&addname=" & Text4.Text & "&note=" & Text6.Text & "&key=" & frmLoad.Text15.Text
      End If
    End If
  End If
End Sub

Private Sub Label58_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label58.BackColor = &HFFC0C0
End Sub

Private Sub Label6_Click()
  End
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Image1.Picture = Image8.Picture
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Image1.Picture = Image7.Picture
End Sub

Private Sub Label62_Click()
  frmAddStuInfo.Show
End Sub

Private Sub Label66_Click()
frmSurChange.Show
End Sub

Private Sub Label66_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label66.BackColor = &HFFC0C0
End Sub

Private Sub Label67_Click()
  MsgBox "对不起，暂不支持此功能！", , "提示"
End Sub

Private Sub Label67_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label67.BackColor = &HFFC0C0
End Sub

Private Sub Label68_Click()
  MsgBox "对不起，暂不支持此功能！", , "提示"
End Sub

Private Sub Label68_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Label68.BackColor = &HFFC0C0
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Image1.Picture = Image4.Picture
End Sub

Private Sub Label8_Click()
  Text1.Visible = True
  Text1.Width = 0
  j = 0
  Timer1.Enabled = True
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Text1.Visible = True
  j = 0
  Timer1.Enabled = True
End Sub

Private Sub ListView1_Click()
  j = 1
  Timer1.Enabled = True
End Sub

Private Sub ListView1_DblClick()
  If ListView1.ListItems(ListView1.SelectedItem.Index).Text = "" Then
  Else
    frmSurEdit.Text1.Text = ListView1.ListItems(ListView1.SelectedItem.Index).Text
    frmSurEdit.Label18.Caption = ListView1.ListItems(ListView1.SelectedItem.Index).Text
    frmSurEdit.Label17.Caption = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1)
    frmSurEdit.Label19.Caption = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2)
    frmSurEdit.Show
  End If
End Sub

Private Sub Main_Move_Click()
  j = 1
  Timer1.Enabled = True
End Sub

Private Sub SearchCredit_Click()
  Frame2.Visible = True
  SearchCredit.Height = 355
  butgai.Visible = True
  SearchCredit.BackColor = &HFF8080
End Sub

Private Sub SearchCredit_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  SearchCredit.BackColor = &HFFC0C0
End Sub

Private Sub SearchId_Click()
      If Len(Text2.Text) <> 11 Then
        MsgBox "请输入正确的学号", 64, "提示"
      Else
        Image13.Left = -Image13.Width + 1950
        Image12.Visible = True
        Image13.Visible = True
        Timer2.Enabled = True
        URL = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=0&id=" & Text2.Text & "&key=" & frmLoad.Text15.Text
        GetAPI = API.Conn(0, URL)
      End If
End Sub

Private Sub SearchId_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  SearchId.BackColor = &HFFC0C0
End Sub

Private Sub SearchName_Click()
        Image13.Left = -Image13.Width + 1950
        Image12.Visible = True
        Image13.Visible = True
        Timer2.Enabled = True
        URL = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=0&name=" & Text2.Text & "&key=" & frmLoad.Text15.Text
        GetAPI = API.Conn(0, URL)
End Sub

Private Sub SearchName_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  SearchName.BackColor = &HFFC0C0
End Sub

Private Sub stro_Click(Index As Integer)
  j = 1
  Timer1.Enabled = True
End Sub

Private Sub stro_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  If i = 0 Then
    stro(0).Picture = Image2.Picture
    Home.Picture = index2.Picture
    stro(1).Picture = Image3.Picture
    Search.Picture = search3.Picture
    stro(2).Picture = Image3.Picture
    Add.Picture = Add3.Picture
    stro(3).Picture = Image3.Picture
    Seething.Picture = seething3.Picture
  ElseIf i = 1 Then
    stro(0).Picture = Image3.Picture
    Home.Picture = index3.Picture
    stro(1).Picture = Image2.Picture
    Search.Picture = search2.Picture
    stro(2).Picture = Image3.Picture
    Add.Picture = Add3.Picture
    stro(3).Picture = Image3.Picture
    Seething.Picture = seething3.Picture
  ElseIf i = 2 Then
    stro(0).Picture = Image3.Picture
    Home.Picture = index3.Picture
    stro(1).Picture = Image3.Picture
    Search.Picture = search3.Picture
    stro(2).Picture = Image2.Picture
    Add.Picture = Add2.Picture
    stro(3).Picture = Image3.Picture
    Seething.Picture = seething3.Picture
  ElseIf i = 3 Then
    stro(0).Picture = Image3.Picture
    Home.Picture = index3.Picture
    stro(1).Picture = Image3.Picture
    Search.Picture = search3.Picture
    stro(2).Picture = Image3.Picture
    Add.Picture = Add3.Picture
    stro(3).Picture = Image2.Picture
    Seething.Picture = seething2.Picture
  End If
End Sub

Private Sub T1_Timer()
  If Picture6.Top < 0 Then
    Picture6.Top = Picture6.Top + 50
  End If
  T2.Enabled = True
End Sub

Private Sub T2_Timer()
  T1.Enabled = False
  T3.Enabled = True
End Sub

Private Sub T3_Timer()
  T2.Enabled = False
  If Picture6.Top > -500 Then
    Picture6.Top = Picture6.Top - 50
  Else
    T3.Enabled = False
    Text3.Text = ""
    Text3.BackColor = &HFFFFFF
    Text4.Text = ""
    Text4.BackColor = &HFFFFFF
    Text6.Text = ""
    Text6.BackColor = &HFFFFFF
  End If
End Sub

Private Sub Text1_Click()
  If Text1.Text = "请输入搜索内容..." Then
    Text1.Text = ""
    Text1.ForeColor = &H0&
  End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)  '搜索框格式限定，检索方式判断
  If KeyAscii = 13 Then
    If IsNumeric(Text1.Text) = True Then
      If Len(Text1.Text) <> 11 Then
        MsgBox "请输入正确的学号", 64, "提示"
      Else
        Image13.Left = -Image13.Width + 1950
        Image12.Visible = True
        Image13.Visible = True
        Timer2.Enabled = True
        URL = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=0&id=" & Text1.Text & "&key=" & frmLoad.Text15.Text
        GetAPI = API.Conn(0, URL)
      End If
    Else
        Image13.Left = -Image13.Width + 1950
        Image12.Visible = True
        Image13.Visible = True
        Timer2.Enabled = True
        URL = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=0&name=" & Text1.Text & "&key=" & frmLoad.Text15.Text
        GetAPI = API.Conn(0, URL)
    End If
  End If
End Sub

Private Sub Text2_Click()
  If Text2.Text = "请输入搜索内容..." Then
    Text2.Text = ""
    Text2.ForeColor = &H0&
  End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)  '搜索框格式限定，检索方式判断
  If KeyAscii = 13 Then
    If IsNumeric(Text2.Text) = True Then
      If Len(Text2.Text) <> 11 Then
        MsgBox "请输入正确的学号", 64, "提示"
      Else
        Image13.Left = -Image13.Width + 1950
        Image12.Visible = True
        Image13.Visible = True
        Timer2.Enabled = True
        URL = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=0&id=" & Text2.Text & "&key=" & frmLoad.Text15.Text
        GetAPI = API.Conn(0, URL)
      End If
    Else
        Image13.Left = -Image13.Width + 1950
        Image12.Visible = True
        Image13.Visible = True
        Timer2.Enabled = True
        URL = frmLoad.Text12.Text & frmLoad.Text13.Text & "/" & frmLoad.Text14.Text & "/?state=0&name=" & Text2.Text & "&key=" & frmLoad.Text15.Text
        GetAPI = API.Conn(0, URL)
    End If
  End If
End Sub

Private Sub Timer1_Timer()
  If j = 0 Then
    If Text1.Width < 3500 Then
    Label8.Enabled = False
      Text1.Width = Text1.Width + 300
    Else
      Timer1.Enabled = False
      Label8.Enabled = True
    End If
  ElseIf j = 1 Then
      If Text1.Width > 200 Then
      Label8.Enabled = False
      Text1.Width = Text1.Width - 300
    Else
      Timer1.Enabled = False
      Text1.Visible = False
      Label8.Enabled = True
      If Text1.Text = "" Then
        Text1.Text = "请输入搜索内容..."
        Text1.ForeColor = &H808080
      End If
    End If
  End If
End Sub

Private Sub Main_Move_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button And vbLeftButton Then
    Me.Move Me.Left - xCursor + x, Me.Top - yCursor + Y
  End If
End Sub

Private Sub Main_Move_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button And vbLeftButton Then
    xCursor = x: yCursor = Y
  End If
End Sub

'Private Sub Bag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Butt_Cro.Picture = Butt_D.Picture
'End Sub


Private Sub Timer2_Timer()   '进度条代码
  If Image13.Left > Image12.Left Then
    Image12.Visible = False
    Image13.Visible = False
    MsgBox "连接服务器超时，请重试！", 64, "提示"
    Timer2.Enabled = False
  Else
    Image13.Left = Image13.Left + 20
End If
End Sub
