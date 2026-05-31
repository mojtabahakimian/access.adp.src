Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub DBLIST_DblClick(CANCEL As Integer)
                   On Error Resume Next
  If Me.OpenArgs = 2 Then
        Forms![getfirstmog]![YEA] = DBLIST.value
  Else
        DoCmd.Close acForm, "BASEKNOW", acSaveYes
        ChangeDB (DBLIST.value)
        DoCmd.OpenForm "BASEKNOW", , , , , acHidden
        DoCmd.Close acForm, "YE"
        DoCmd.OpenForm "YE"
  End If
  DoCmd.Close acForm, Me.NAME
End Sub

Private Sub DBLIST_KeyPress(KeyAscii As Integer)
                  On Error Resume Next
  If KeyAscii = 13 Then
    If Me.OpenArgs = 2 Then
          Forms![getfirstmog]![YEA] = DBLIST.value
    Else
          DoCmd.Close acForm, "BASEKNOW", acSaveYes
          ChangeDB (DBLIST.value)
          DoCmd.OpenForm "BASEKNOW", , , , , acHidden
          DoCmd.Close acForm, "YE"
          If (MID(Forms![BASEKNOW]![tindata], 20, 7) <> "CORRECT" Or IsNull(Forms![BASEKNOW]![tindata])) And False Then
            DoCmd.OpenForm "YE"
          End If

    End If
    DoCmd.Close acForm, Me.NAME
  End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.Close acForm, "F_SELECTYEAR"
   Dim rst As New ADODB.Recordset, SST As New ADODB.Recordset
   DBLIST.AddItem ("نام فايل;سال مالي;شرح")
   rst.Open "SELECT * FROM  master.dbo.sysdatabases ", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
   While Not rst.EOF
      Err.Clear
      SST.Open "SELECT YEA,NAME,SANAT FROM " & rst.Fields("NAME") & ".DBO.SAZMAN", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
      If Err = 0 Then
        DBLIST.AddItem (rst.Fields("name") & ";" & SST.Fields("yea") & ";" & SST.Fields("name"))
      End If
      rst.MoveNext
      Set SST = New ADODB.Recordset
   Wend
End Sub
