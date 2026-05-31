Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command4_Click()
    Call Shell("calc", 1)
    SendkeyUS (Me.Text2)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
    If KeyAscii = 27 Then
       DoCmd.Close
    End If
End Sub

Private Sub Form_Load()
                     On Error Resume Next
   Dim i, ST, le As Long, frm As Form, CTL As Control, JAM As Double
   Set frm = Screen.ActiveDatasheet
   Set CTL = Screen.ActiveControl
   ST = frm.SelTop
   le = frm.SelHeight
   DoCmd.GoToRecord acDataForm, frm.NAME, acGoTo, ST
   JAM = 0
   For i = 1 To le
      JAM = JAM + frm.Form(CTL.NAME)
      DoCmd.GoToRecord acDataForm, frm.NAME, acNext
   Next i
   DoCmd.GoToRecord acDataForm, frm.NAME, acGoTo, ST
   Me.Text0 = JAM
   Me.Text2 = JAM

End Sub
