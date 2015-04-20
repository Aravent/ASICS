Dim Str As String
Dim a, b, C As Integer

Private Sub Command1_Click()
  Text2.Text = MD5_32Bits(Text1.Text)
End Sub

Private Sub Command2_Click()
  Str = Inet1.OpenURL("http://www.anycen.com/api/student/getstudentbeta/?state=5&username=" & Text3.Text & "&key=e1152c54e099a972f1471a")
  a = InStr(2, Str, "}")
  b = InStr(a + 1, Str, "}")
  Text4.Text = Mid(Str, 2, a - 2)
  Text5.Text = Mid(Str, a + 2, b - a - 2)
End Sub

Private Sub Command3_Click()
Dim st1, st2, st3, st4, Zo As Integer
Open App.Path & "\123.txt" For Input As #1
Input #1, Zo
Input #1, st1
Input #1, st2
Input #1, st3
Text1 = st1
Text2 = st2
Text3 = st3
Close
End Sub

Private Sub Form_Load()
  Me.Caption = "11"
End Sub
