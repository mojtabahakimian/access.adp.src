Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub DEGR_DblClick(CANCEL As Integer)
      DoCmd.OpenForm "FROOSH_COUNTALL_MP4", acFormDS, , "SAIZE = " & Me.SAIZE & " AND TARH = " & Me.TARH & " AND DEGR = " & Me.DEGR
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      DoCmd.OpenForm "FROOSH_COUNTALL_MP4", acFormDS, , "SAIZE = " & Me.SAIZE & " AND TARH = " & Me.TARH & " AND DEGR = " & Me.DEGR
   End If
  If KeyAscii = 27 Then
     DoCmd.Close acForm, Me.NAME
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
  Me.RecordSource = "SELECT SUM(MEGHKF) AS MEGHF, SUM(MABK) AS MABKK, SUM(MEGHkK) AS MEGHk, SAIZE, TARH, CAST(TARH AS INT) AS TTARH, DEGR FROM FROOSH_COUNTALL_CITY_N_FANI(" & Forms![F_MENU_DATE]![DT1] & "," & Forms![F_MENU_DATE]![DT2] & ") FROOSH_COUNTALL_CITY_N_FANI GROUP BY SAIZE, TARH, CAST(TARH AS INT), DEGR"
End Sub
