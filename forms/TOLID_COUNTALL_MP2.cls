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
   If KeyAscii = 13 Then
      DoCmd.OpenForm "TOLID_COUNTALL_MP3", acFormDS, , "SAIZE = " & Me.SAIZE & " AND TARH = " & Me.TARH
   End If
  If KeyAscii = 27 Then
     DoCmd.Close acForm, Me.NAME
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
   Me.RecordSource = "SELECT SUM(MEGHT) AS MEGHk, SAIZE, SUM(MEGHKF) AS MEGHF, TARH, CAST(TARH AS INT) AS TTARH FROM AMAR_TOLID_FAREE(" & Forms![F_MENU_DATE]![DT1] & "," & Forms![F_MENU_DATE]![DT2] & ") AMAR_TOLID_FAREE GROUP BY SAIZE, TARH, CAST(TARH AS INT)"
End Sub

Private Sub ttarh_DblClick(CANCEL As Integer)
  DoCmd.OpenForm "TOLID_COUNTALL_MP3", acFormDS, , "SAIZE = " & Me.SAIZE & " AND TARH = " & Me.TARH
End Sub
