Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub CODE_AfterUpdate()
                     On Error Resume Next
  Dim rst As New ADODB.Recordset
   On Error GoTo ERRD
  If Me.CODE.OldValue <> Me.CODE Then
    rst.Open "stuf_def", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me![CODE] & "'"
    Set rst = rst
    If rst.RecordCount = 0 Then
      If Me.NewRecord Then
          DoCmd.RunCommand acCmdSaveRecord
      End If
    Else
        Err.Clear
        Err.Raise 3022
    End If
    Exit Sub
  End If
  Me.Refresh
  Exit Sub
ERRD:       errdetector (Err)
End Sub

Private Sub Command39_Click()
                      On Error Resume Next
   Me.OrderBy = "nc"
   Me.OrderByOn = True
End Sub


Private Sub Form_BeforeUpdate(CANCEL As Integer)
                        On Error Resume Next
     Dim rst As New ADODB.Recordset
     If IsNull(Me.CODE) Or Me.CODE = "" Then
         rst.Open "SELECT Max(cast(CODE as int)) AS Expr1 FROM STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If Not IsNull(rst.Fields(0)) And rst.RecordCount > 0 Then
            Me.CODE = Trim(STR(rst.Fields(0) + 1))
        Else
          Me.CODE = 1
        End If
     End If

End Sub

Private Sub Form_Delete(CANCEL As Integer)
                       On Error Resume Next
   Forms![BASEKNOW]![Text44] = False
   DoCmd.OpenForm "deleteform", , , , , acDialog
   If Forms![BASEKNOW]![Text44] Then
   Else
     CANCEL = True
   End If
   Me.Refresh
End Sub

Private Sub Form_Load()
                     On Error Resume Next
'  Me.lock2.Password = "neginsedighahmad"
 ' Me.lock2.Connected = True
 ' If Me.lock2.ErrorCode <> 0 And Me.lock2.ErrorCode <> 7 Then
 '    DoCmd.OpenForm "lockok"
'  End If
  DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
  Call SETSECURITY(Me.NAME, "KHAD", 3)
  Me.nam.SetFocus
End Sub
Private Sub Form_AfterUpdate()
                     On Error Resume Next
     Dim rst As New ADODB.Recordset
     rst.Open "select * from TDETA_HES where N_KOL = " & Forms![BASEKNOW]![DARAM] & " AND " & "NUMBER = " & Me.DEPART & " AND " & "TNUMBER = " & Val(Me.CODE), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount = 0 Then
         rst.AddNew
      End If
    rst.Fields("N_KOL") = Forms![BASEKNOW]![DARAM]
    rst.Fields("NUMBER") = Me.DEPART
    rst.Fields("TNUMBER") = Val(Me.CODE)
    rst.Fields("NAME") = Me.nam
    rst.Fields("BED_BES") = -1
    rst.update
    rst.Close
    Set rst = New ADODB.Recordset
     rst.Open "select * from STUF_FSK where CODE = '" & Me.CODE & "' AND ANBAR = 0 ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount = 0 Then
         rst.AddNew
      End If
    rst.Fields("CODE") = Me.CODE
    rst.Fields("ANBAR") = 0
    rst.update
    rst.Close
    Set rst = New ADODB.Recordset
     rst.Open "SELECT * FROM STUF_STK WHERE  CODE = '" & Me.CODE & "' AND ANBAR = 0 ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount = 0 Then
         rst.AddNew
      Else
      End If
    rst.Fields("CODE") = Me.CODE
    rst.Fields("ANBAR") = 0
    rst.update
    rst.Close
End Sub




Private Sub NAME_AfterUpdate()
                        On Error Resume Next
     Me![NAME] = Trim(Me![NAME])
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub MABL_F_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub VAHED_DblClick(CANCEL As Integer)
                    On Error Resume Next

       Dim VAHEDID As String
    If LETSGO("VAHED") Then
       DoCmd.OpenForm "TCOD_VAHEDS", , , , , acDialog, "GotoNew"
       Me.vahed.Undo
       Me.vahed.Requery
    Else
          DoCmd.OpenForm "mesag", , , , , acDialog, "شما اجازه لازم براي استفاده از اين بخش را نداريد.!"
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
