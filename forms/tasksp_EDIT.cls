Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command114_Click()
    Dim RETVAL
                  On Error Resume Next
  If IDNUM <> "" Then
    If (Dir(Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, vbDirectory) <> "") Then
        RETVAL = Shell("explorer.exe /e," & Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, 1)
    Else
        MkDir (Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM)
        RETVAL = Shell("explorer.exe /e," & Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, 1)
   End If
  End If
End Sub

Private Sub Command28_Click()
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub COMP_COD_DblClick(CANCEL As Integer)
  If IsNull(Me.COMP_COD) Or Me.COMP_COD = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " مشتري مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  Else
     If BLOCKEDMK(COMP_COD) Then
        DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
        Exit Sub
     End If
  End If
  DoCmd.OpenForm "TASKS_COMPANY", acFormDS, , , , , Me.COMP_COD
End Sub

Private Sub COMP_COD_Exit(CANCEL As Integer)
                 On Error Resume Next
  If IsNull(Me.COMP_COD) Or Me.COMP_COD = "0" Then
  Else
    If BLOCKEDCUST(Me.COMP_COD) And Me.NewRecord Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " حساب  مسدود گرديده است لطفا با مديريت مالي تماس بگيريد"
        CANCEL = True
    End If
  End If
End Sub

Private Sub COMP_COD_NotInList(NewData As String, Response As Integer)
                 On Error Resume Next
 PLUS = False
  If Me.COMP_COD.Text = "-" Or Me.COMP_COD.Text = "+" Then
     Me.COMP_COD.Text = ""
     DoCmd.OpenForm "SERCH", , , , , acDialog
     PLUS = True
  Else
    Me.COMP_COD = NewData
  End If
  Response = acDataErrContinue
  Me.COMP_COD.SetFocus
End Sub

Private Sub Form_Close()
                 On Error Resume Next
   Forms![main]![TASKS].Form.Requery
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
    Dim rst As New ADODB.Recordset, VS As String
    Me.RecordSource = "SELECT TASKS.* FROM TASKS WHERE IDNUM = " & Me.OpenArgs
    VS = ""
    rst.Open "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL=0) and idd <> 120", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        VS = VS & rst.Fields("IDD") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & Chr(34)
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    Me.PERSONEL.RowSource = VS
    DoCmd.RunSQL "DELETE FROM dbo.temp_us WHERE     (USERCOT = " & Forms![BASEKNOW]![USERCOD] & ")"
    DoCmd.RunSQL "INSERT INTO dbo.temp_us (USERCO, tic, USERCOT) SELECT     IDD, 0 AS Expr1, " & Forms![BASEKNOW]![USERCOD] & " AS Expr2 FROM dbo.SALA_DTL WHERE     (ENABL = 0)"
    Me.COMP_COD.DefaultValue = Chr(34) & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & Chr(34)
    Me.taskspevent.SetFocus
End Sub
Private Sub IDNUM_DblClick(CANCEL As Integer)
    Dim RETVAL
                  On Error Resume Next
    If (Dir(Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, vbDirectory) <> "") Then
        RETVAL = Shell("explorer.exe /e," & Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, 1)
    Else
        MkDir (Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM)
        RETVAL = Shell("explorer.exe /e," & Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, 1)
   End If
End Sub

Private Sub TASK_AfterUpdate()
   If Me.TASK = "+" Or Me.TASK = "++" Or Me.TASK = "-" Or Me.TASK = "--" Then
      DoCmd.OpenForm "sharh_list", acFormDS
   End If
End Sub

Private Sub TASK_Change()
    If Len(Me.TASK.Text) > 255 Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " متن طولاني است ادامه متن را بعداز ذخيره در ذيل بنويسيد"
    End If
End Sub

Private Sub TASK_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "zoom", acFormDS, , , , , "1" & Me.TASK
End Sub

Private Sub taskspevent_Enter()
                  On Error Resume Next
  DoCmd.RunCommand acCmdRecordsGoToNew

End Sub
