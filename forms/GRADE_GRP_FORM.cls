Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_AfterUpdate()
    UpAZAE (Trim(Forms![AZAE].Form.HES))
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
              On Error Resume Next
  If IsNull(Me.GCGRPNAME) Or Me.GCGRPNAME = "" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " عنوان مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If IsNull(Me.GCGRPID) Then
    Dim rst As New ADODB.Recordset
 '***************************************************************
  Set rst = CurrentProject.Connection.Execute("SELECT Max(GCGRPID) AS MIDD FROM GRADE_CUST_GRP")
 '***************************************************************
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.GCGRPID = 1
    Else
        Me.GCGRPID = rst.Fields(0) + 1
    End If
 End If
 End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub GCVALUESCAL_AfterUpdate()
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT  GSCAGRADE FROM GSCADTL WHERE GSCADTCOD = " & Me.GCVALUESCAL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GCGRPGRADE = rst.Fields(0)
    End If
End Sub

Private Sub GCVALUESCAL_Enter()
    Me.GCVALUESCAL.RowSource = "SELECT GSCADTCOD, GSCANAME FROM GSCADTL WHERE GSCACOD = " & Me.GCVALUE
    Me.GCVALUESCAL.Requery
End Sub

Private Sub GCVALUESCAL_Exit(CANCEL As Integer)
    Me.GCVALUESCAL.RowSource = "SELECT GSCADTCOD, GSCANAME FROM GSCADTL "
    Me.GCVALUESCAL.Requery
End Sub
