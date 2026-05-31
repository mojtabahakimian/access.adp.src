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
                       On Error Resume Next
'   If KeyAscii = 13 Then
'     DoCmd.OpenForm "froosh_mp10", acFormDS, , "col1 = " & Me.col1 & " and col2 = " & Me.col3 & " and col3 = " & Me.col3 & " and col4 = " & Me.col4 & " and col5 = " & Me.col5 & " and col6 = " & Me.col6 & " and col7 = " & Me.col7 & " and col8 = " & Me.col8
'   End If
  If KeyAscii = 27 Then
     DoCmd.Close acForm, Me.NAME
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
    Dim rst As New ADODB.Recordset
  Me.RecordSource = "SELECT SUM(MABL_K) AS mabkk, col1, col2, col3, col4, col5, col6, col7, col8, col9, COUNT(NUMBER) AS cont FROM FROOSH_MP(" & Forms![F_MENU_DATE]![DT1] & "," & Forms![F_MENU_DATE]![DT2] & ")  FROOSH_MP GROUP BY col1, col2, col3, col4, col5, col6, col7, col8, col9"
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col1l.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col2l2.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 3", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col3l2.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 4", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col4l2.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 5", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col5l2.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 6", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.coL6l2.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 7", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.coL7l2.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 8", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.coL8l2.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 9", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col9l.CAPTION = rst.Fields("MPNAME")
        Me.coL9l2.CAPTION = rst.Fields("MPNAME")
  End If

End Sub


Private Sub col2_DblClick(CANCEL As Integer)
  '    DoCmd.OpenForm "froosh_mp10", acFormDS, , "col1 = " & Me.col1 & " and col2 = " & Me.col3 & " and col3 = " & Me.col3 & " and col4 = " & Me.col4 & " and col5 = " & Me.col5 & " and col6 = " & Me.col6 & " and col7 = " & Me.col7
End Sub
