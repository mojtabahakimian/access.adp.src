Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Current()
                 On Error Resume Next
  DoCmd.RunCommand acCmdSelectRecord
  Forms![USER_LOGIN]![usname] = Me.SAL_NAME1
  Forms![USER_LOGIN]![cusname] = Me.SAL_NAME
  DoCmd.RunCommand acCmdSelectRecord

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                 On Error Resume Next
   If KeyAscii = 13 Then
         Forms![USER_LOGIN]![PSWORD].SetFocus
   End If
   If KeyAscii = 27 Then
      DoCmd.Quit
   End If
End Sub

Private Sub Form_Load()
                     On Error Resume Next
  Dim uts As String
  Dim fs, a
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.OPENTextFile("c:\uswintm.ini", 1)
    If Err <> 0 Then
    Else
       uts = a.readline
       Me.IDD.SetFocus
       DoCmd.FindRecord uts, acEntire, , acSearchAll, , acCurrent, True

    End If
a.Close

 ' DoCmd.RunCommand acCmdSelectRecord
'  SendKeys ("{ENTER}")
End Sub
