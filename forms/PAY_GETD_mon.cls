Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit




Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 49 Then
   Dim i, ST, le As Long, frm As Form, CTL As Control, JAM As Double
                    On Error Resume Next
   ST = Me.SelTop
   le = Me.SelHeight
   DoCmd.GoToRecord acActiveDataObject, , acGoTo, ST
   JAM = 0
   For i = 1 To le
      JAM = JAM + Me.mabl
      DoCmd.GoToRecord acActiveDataObject, , acNext
   Next i
   DoCmd.GoToRecord acActiveDataObject, , acGoTo, ST
   Forms![chek_mon].JJKOL.ControlSource = ""
   Forms![chek_mon].JJKOL = JAM
 End If
End Sub
