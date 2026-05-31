Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
   DoCmd.OpenForm "FROOSH_COUNTALL_CITY2", acFormDS
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  Me.RecordSource = "SELECT SUM(MEGHKF) AS MEGHF, MPNAME, SUM(MABK) AS MABKK, SUM(MEGHkK) AS MEGHk FROM FROOSH_COUNTALL_CITY_N_FANI(" & Forms![F_MENU_DATE]![DT1] & "," & Forms![F_MENU_DATE]![DT2] & ") FROOSH_COUNTALL_CITY_N_FANI GROUP BY MPNAME ORDER BY MPNAME"
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub MEGHF_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "FROOSH_COUNTALL_CITY2", acFormDS, , "MPNAME = '" & Me.MPNAME & "'"
End Sub

Private Sub MEGHk_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "FROOSH_COUNTALL_CITY1_OSTAN", acFormDS
End Sub

Private Sub MPNAME_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "FROOSH_COUNTALL_CITY2", acFormDS
End Sub
