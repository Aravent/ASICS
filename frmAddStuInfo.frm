Dim i As Integer


Private Sub Command1_Click()
  If i = 1 Then
    Image1.Visible = False
    Image2.Visible = True
    Image2.Top = 0
    Image2.Left = 0
    Command2.Visible = True
    i = 2
  ElseIf i = 2 Then
    Image2.Visible = False
    Image3.Visible = True
    Image3.Top = 0
    Image3.Left = 0
    Command1.Caption = "完成"
    Command1.Enabled = False
    Text1.Visible = True
    Label1.Visible = True
    Label2.Visible = True
    i = 3
  ElseIf i = 3 Then
    Unload Me
  End If
End Sub

Private Sub Command2_Click()
  If i = 1 Then
    Command2.Visible = False
  ElseIf i = 2 Then
    Command2.Visible = False
    Image1.Visible = True
    Image2.Visible = False
    Image3.Visible = False
    Image1.Top = 0
    Image1.Left = 0
    i = 1
  ElseIf i = 3 Then
    Image1.Visible = False
    Image2.Visible = True
    Image3.Visible = False
    Image2.Top = 0
    Image2.Left = 0
    Command1.Caption = "下一步"
    Command1.Enabled = True
    Text1.Visible = False
    Label1.Visible = False
    Label2.Visible = False
    i = 2
  End If
End Sub

Private Sub Form_Load()
  i = 1
End Sub

Private Sub Label1_Click()
Dim OpenFile As String
  On Error GoTo ErrHandler
  CommonDialog1.Filter = "所有文件 (*.*)|*.*|Excle文件 (*.xls)|*.xls"
  CommonDialog1.FilterIndex = 2
  CommonDialog1.ShowOpen
  OpenFile = CommonDialog1.FileName
  Text1.Text = OpenFile
  Label2.Enabled = True
  Exit Sub
ErrHandler:
  Exit Sub
End Sub

Private Sub Label2_Click()
  If MsgBox("是否执行添加？", vbOKCancel, "请确认您的选择") = vbOK Then
    frmAdd.Text3.Text = CommonDialog1.FileName
    frmAdd.Show
    Unload Me
  Else
    MsgBox "Cancel"
  End If
End Sub
