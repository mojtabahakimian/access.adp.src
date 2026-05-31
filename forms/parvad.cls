Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Dim nam As String

Private Sub Command2_Click()
    If IsNull(Me.COMP_COD) Or Me.COMP_COD = "0" Then
      DoCmd.OpenForm "mesageform", , , , , acDialog, " مشتري مشخص نشده است ....!"
      Exit Sub
    Else
       If BLOCKEDMK(Me.COMP_COD) Then
          DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
          Exit Sub
       End If
    End If
    DoCmd.OpenForm "TASKS_COMPANY", acFormDS, , , , , Me.COMP_COD
      DoCmd.OpenForm Me.NAME, , , , , acHidden

End Sub

Private Sub Command3_Click()
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub COMP_COD_NotInList(NewData As String, Response As Integer)
                 On Error Resume Next
Dim rst As New ADODB.Recordset, SPACELESSV As String
                      On Error Resume Next
    If COMP_COD.Text = "-" Or COMP_COD.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "47"
       Response = acDataErrContinue
    Else
        If IsNumeric(NewData) Then
          Err.Clear
          rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If Err = 0 Then
            If rst.RecordCount = 1 Then
                 Me.COMP_COD = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                 Response = acDataErrContinue
            End If
          End If
        Else
            SPACELESSV = Replace(NewData, " ", "")
            DoCmd.OpenForm "SERSNDTAF", , , , , , "47" & "NAME LIKE '%" & NewData & "%'  OR REPLACE(dbo.TDETA_HES.NAME, ' ', '')  LIKE '%" & SPACELESSV & "%'"
           Response = acDataErrContinue
        End If
    End If
  Response = acDataErrContinue

End Sub

Private Sub Form_Activate()
  DoCmd.restore
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                   On Error Resume Next
   If KeyAscii = 27 Then
      DoCmd.Close acForm, Me.NAME
   End If
End Sub

Private Sub Form_Load()
                  On Error Resume Next
  If IsNull(Me.COMP_COD) Then
    Me.COMP_COD = Me.OpenArgs
 End If
  DoCmd.restore
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
