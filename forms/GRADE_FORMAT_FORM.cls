Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command14_Click()
    DoCmd.OpenReport "GRADE_FORMAT_REP", acViewPreview, , "IDD=" & Me.IDD
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next
  If IsNull(Me.GFNAME) Or Me.GFNAME = "" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " عنوان مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If IsNull(Me.IDD) Then
    Dim rst As New ADODB.Recordset
 '***************************************************************
     '   CurrentProject.Connection.Execute ("set transaction isolation level SERIALIZABLE"): DoCmd.OpenForm "BUN"
     '   CurrentProject.Connection.Execute ("begin transaction")
        Set rst = CurrentProject.Connection.Execute("SELECT Max(IDD) AS MIDD FROM GRADE_FORMAT")
 '***************************************************************
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.IDD = 1
    Else
        Me.IDD = rst.Fields(0) + 1
    End If
 End If
  Me.GRADE_TAB_FT_FORM.Locked = False
  Me.GRADE_TAB_FT_FORM.Enabled = True
  Me.GRADE_TAB_FT_FORM.Requery

End Sub

Private Sub Form_Current()
                On Error Resume Next
    If Me.OpenArgs = "1" Then
       DoCmd.RunCommand (acCmdSelectRecord)
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
            On Error Resume Next
 If Me.OpenArgs = "1" Then
    Dim rst As New ADODB.Recordset, RSTS As New ADODB.Recordset, RSTSG As New ADODB.Recordset, RSTSGR As New ADODB.Recordset, GCTAB As Long
    If KeyAscii = 13 Then
       rst.Open "select * from grade_cust_tab where GCCUST_HES = '" & Trim(Forms![AZAE]![HES]) & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " !....گريد بندي قبلا انجام شده  "
       Else
            RSTS.Open "select * from grade_tab_ft where gfid = " & Me.IDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not RSTS.EOF
                rst.AddNew
                GCTAB = GetLGradetab()
                rst.Fields("GCTABID") = GCTAB
                rst.Fields("GCNAME") = RSTS.Fields("GFNAMEFT")
                rst.Fields("GCZARIB") = RSTS.Fields("GFGZARIB")
                rst.Fields("GCCUST_HES") = Trim(Forms![AZAE]![HES])
                rst.Fields("GCDATE") = FARSIDATE(DATE)
                rst.Fields("USERNAME") = UCurrentUser()
                rst.update
               Call INSERTGRPGRADE(GCTAB, RSTS.Fields("GFTID"))
                RSTS.MoveNext
            Wend
        End If
    End If
    DoCmd.Close acForm, Me.NAME
    Forms![AZAE].Form.GRADE_TAB_FORM.Requery
    UpAZAE (Trim(Forms![AZAE].Form.HES))
End If
End Sub

Private Sub INSERTGRPGRADE(GCTABIDP As Long, GFTIDP As Long)
                On Error Resume Next
    Dim rst As New ADODB.Recordset, RSTS As New ADODB.Recordset
    rst.Open "grade_cust_GRP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    RSTS.Open "select * from grade_GRP_ft where gfTid = " & GFTIDP, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not RSTS.EOF
        rst.AddNew
        rst.Fields("GCTABID") = GCTABIDP
        rst.Fields("GCGRPNAME") = RSTS.Fields("GFGRPNAMEFT")
        rst.Fields("GCGRPZARIB") = RSTS.Fields("GFGRPZARIB")
        rst.Fields("GCGRPGRADE") = RSTS.Fields("GFGRPGRADE")
        rst.Fields("GCGRPID") = GetLGradeGRP()
        rst.Fields("GCVALUE") = RSTS.Fields("GVALUESCALE")
        rst.update
        RSTS.MoveNext
    Wend

End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub GFDATE_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(GFDATE, Forms![BASEKNOW]![CTL_DT])
End Sub
