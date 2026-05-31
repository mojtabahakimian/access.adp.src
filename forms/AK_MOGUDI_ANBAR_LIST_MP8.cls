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
      DoCmd.OpenForm "AK_MOGUDI_ANBAR_LIST_MP9", acFormDS, , "col1 = " & Me.col1 & " AND col2 = " & Me.col2 & " AND col3 = " & Me.col3 & " AND col4 = " & Me.col4 & " AND col5 = " & Me.col5 & " AND col6 = " & Me.col6 & " AND col7 = " & Me.col7 & " AND col8 = " & Me.col8
   End If
  If KeyAscii = 27 Then
     DoCmd.Close acForm, Me.NAME
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Dim rst As New ADODB.Recordset, SANBAR As String
    If IsLoaded("ak_mogudi_anbar_list") Then
        SANBAR = Forms![ak_mogudi_anbar_list]![ANBAR]
    Else
        SANBAR = Forms![MOGUDI_SEARCH_MAIN]![ANBAR]
    End If
   Me.RecordSource = "SELECT SUM(MAND) AS MANDS, col1, ANBAR, SUM(MANDF) AS MANDFS, ANBARN,  col2, col3, col4,col5,col6,col7,col8 FROM AKMOGUDI_KOL_ANBAR_N_FANI(99999999, " & SANBAR & ") AKMOGUDI_KOL_ANBAR_N_FANI GROUP BY col1, ANBAR, ANBARN,  col2, col3, col4,col5,col6,col7,col8"
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col1l.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col2l.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 3", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col3l.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 4", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col4l.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 5", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col5l.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 6", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col6l.CAPTION = rst.Fields("MPNAME")
  End If
  Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 7", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col7l.CAPTION = rst.Fields("MPNAME")
  End If
    Set rst = New ADODB.Recordset
  rst.Open "select MPNAME from TCOD_MAP_GRP where MPP = 8", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
        Me.col8l.CAPTION = rst.Fields("MPNAME")
        Me.col5l2.CAPTION = rst.Fields("MPNAME")
  End If
  End Sub
