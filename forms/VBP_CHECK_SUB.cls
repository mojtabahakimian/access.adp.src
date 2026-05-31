Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Private Sub Form_Current()
                        On Error Resume Next
 Me.N_TAF.RowSource = "SELECT     TNUMBER , NAME FROM TDETA_HES WHERE     (N_KOL = " & Me.N_KOL & " AND NUMBER = " & Me.N_MOIN & ")"

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
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
   Forms![VBP_CHECK].JJKOL.ControlSource = ""
   Forms![VBP_CHECK].JJKOL = JAM
 End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
   Me.N_TAF.RowSource = "SELECT     TNUMBER , NAME FROM TDETA_HES WHERE     (N_KOL = " & Me.N_KOL & " AND NUMBER = " & Me.N_MOIN & ")"
   Err.Clear
End Sub
