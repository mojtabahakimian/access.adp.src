Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "F_MENU_ANBAR"
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      DoCmd.OpenForm "TOLID_COUNTALL_MP2", acFormDS, , "SAIZE = " & Me.SAIZE
   End If
  If KeyAscii = 27 Then
     DoCmd.Close acForm, Me.NAME
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
  Me.RecordSource = "SELECT SUM(MEGHT) AS MEGHK, SAIZE, SUM(MEGHKF) AS MEGHF FROM AMAR_TOLID_FAREE(" & Forms![F_MENU_DATE]![DT1] & "," & Forms![F_MENU_DATE]![DT2] & ") AMAR_TOLID_FAREE GROUP BY SAIZE"

End Sub


Private Sub SAIZE_DblClick(CANCEL As Integer)
      DoCmd.OpenForm "TOLID_COUNTALL_MP2", acFormDS, , "SAIZE = " & Me.SAIZE
End Sub
