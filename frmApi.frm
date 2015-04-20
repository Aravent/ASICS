Dim URL, getStr, addStr, get1, get2, get3, get4, get5 As String
Dim i, a1, a2, a3, Del As Integer

Public Function Conn(a, b) As String
  
  If a = 0 Then
    i = 1
    getStr = Inet0.OpenURL(b)
      
      If Len(getStr) = 37 Then
        frmMain.ListView1.ListItems.Clear
        MsgBox "没有找到学生信息，请核对后重试！", 64, "提示"
        frmMain.Image12.Visible = False
        frmMain.Image13.Visible = False
        frmMain.Timer2.Enabled = False
      Else
        frmMain.ListView1.ListItems.Clear               '清空列表
        frmMain.ListView1.ColumnHeaders.Clear           '清空列表头
        frmMain.ListView1.View = lvwReport              '设置列表显示方式
        frmMain.ListView1.Gridlines = True              '显示网络线
        frmMain.ListView1.LabelEdit = lvwManual         '禁止标签编辑
        frmMain.ListView1.FullRowSelect = True          '选择整行
 
        frmMain.ListView1.ColumnHeaders.Add , , "学号", 2000 '给列表中添加列名
        frmMain.ListView1.ColumnHeaders.Add , , "姓名", 1500
        frmMain.ListView1.ColumnHeaders.Add , , "学分", 1500
        frmMain.ListView1.ColumnHeaders.Add , , "操作", 3500
        frmMain.ListView1.ColumnHeaders.Add , , "备注", 3500
        
        Do
        On Error GoTo over
          DoEvents
          
          get1 = Fun_GetStr(getStr, "id"":""", """,")     '提取信息
          get2 = Fun_GetStr(getStr, "name"":""", """,")
          get3 = Fun_GetStr(getStr, "credit"":""", """,")
          get4 = Fun_GetStr(getStr, "option"":""", """,")
          get5 = Fun_GetStr(getStr, "note"":""", """")
         
          frmMain.ListView1.ListItems.Add , , get1
          frmMain.ListView1.ListItems(i).SubItems(1) = get2
          frmMain.ListView1.ListItems(i).SubItems(2) = get3
          frmMain.ListView1.ListItems(i).SubItems(3) = get4
          frmMain.ListView1.ListItems(i).SubItems(4) = get5
          
          Del = InStr(getStr, """}")
          getStr = Mid(getStr, Del + 2, Len(getStr) - Del)   '判断信息是否结束，用以结束循环
          i = i + 1
          
          If getStr = "" Then
          Exit Do
          End If
          
      Loop
over:
          frmMain.Image12.Visible = False
          frmMain.Image13.Visible = False
          frmMain.Timer2.Enabled = False
          Call frmMain.Label2_Click
          frmMain.Frame1.Visible = True
      End If
  
  ElseIf a = 1 Then
    
    addStr = Inet1.OpenURL(b)
    'frmMain.Text7.Text = b
    If addStr = "{""status"":200,""message"":""ok"",""data"":{""information"":""success""}}" Then
      Call frmLoad.Timer1_Timer
      Call frmMain.Command2_Click
    Else
      MsgBox "学生信息添加失败，请重试！", 64, "提示"
    End If
    
    frmMain.Image12.Visible = False
    frmMain.Image13.Visible = False
    frmMain.Timer2.Enabled = False
    Call frmMain.Label3_Click
    frmMain.Frame3.Visible = True
  
  ElseIf a = 2 Then  '学分修改单个

    getStr = Inet2.OpenURL(b)
      
      If Len(getStr) = 37 Then
        MsgBox "没有找到学生信息，请核对后重试！", 64, "提示"
        frmSurEdit.Text1.Text = ""
      Else
          
          frmSurEdit.Text1.Text = Fun_GetStr(getStr, "id"":""", """,")      '提取信息
          frmSurEdit.Label18.Caption = Fun_GetStr(getStr, "id"":""", """,")
          frmSurEdit.Label17.Caption = Fun_GetStr(getStr, "name"":""", """,")
          frmSurEdit.Label19.Caption = Fun_GetStr(getStr, "credit"":""", """,")
         
      End If
      
  ElseIf a = 3 Then

  ElseIf a = 5 Then

  ElseIf a = 6 Then

  End If
  
End Function
