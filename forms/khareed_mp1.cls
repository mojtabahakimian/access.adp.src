Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "F_MENU_date"
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      DoCmd.OpenForm "KHAREED_MP2", acFormDS, , "col1 = " & Me.col1
   End If
  If KeyAscii = 27 Then
     DoCmd.Close acForm, Me.NAME
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Dim rst As New ADODB.Recordset
  Me.RecordSource = "SELECT SUM(MABL_K) AS mabkk, col1, COUNT(NUMBER) AS cont FROM KHAREED_MP(" & Forms![F_MENU_DATE]![DT1] & "," & Forms![F_MENU_DATE]![DT2] & ")  KHAREED_MP GROUP BY col1 ORDER BY SUM(MABL_K) DESC"
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col1l.CAPTION = rst.Fields("MPNAME")
        Me.col1l2.CAPTION = rst.Fields("MPNAME")
  End If

End Sub


Private Sub col1_DblClick(CANCEL As Integer)
      DoCmd.OpenForm "KHAREED_MP2", acFormDS, , "col1 = " & Me.col1
End Sub
