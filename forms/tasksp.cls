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

Private Sub Command26_Click()
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
            On Error Resume Next
    If IsNull(Me.PERSONEL) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " 'گيرنده مشخص نيست....!"
        Exit Sub
    End If
    If IsNull(Me.COMP_COD) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " 'تماس گيرنده مشخص نيست...!"
        Exit Sub
    End If
    If Me.IDNUM = "" Or IsNull(Me.IDNUM) Then
        rst.Open "SELECT     * FROM dbo.TASKS WHERE     (IDNUM = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
    Else
        rst.Open "SELECT     * FROM dbo.TASKS WHERE     (IDNUM = " & Me.IDNUM & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    End If
    rst.Fields("PERSONEL") = Me.PERSONEL
    rst.Fields("COMP_COD") = Me.COMP_COD
    rst.Fields("TASK") = RIGHT(Me.TASK, 255)
    rst.Fields("PERIORITY") = Me.PERIORITY
    rst.Fields("STATUS") = Me.STATUS
    rst.Fields("STDATE") = Me.STDATE
    rst.Fields("STTIME") = Me.STTIME
    rst.Fields("USERNAME") = Me.USERNAME
    rst.Fields("GR") = Me.GR
    rst.Fields("ENDATE") = Me.ENDATE
    rst.Fields("ENTIME") = Me.ENTIME
'    RST.Fields("PIC") = Me.PIC
    rst.update
RST2.Open "SELECT     * FROM dbo.temp_us WHERE     (tic = 1) AND (USERCOT = " & Forms![BASEKNOW]![USERCOD] & ") and (userco <> " & Me.PERSONEL & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If RST2.RecordCount > 0 Then
        While Not RST2.EOF
            rst.AddNew
            rst.Fields("PERSONEL") = RST2.Fields("userco")
            rst.Fields("COMP_COD") = Me.COMP_COD
            rst.Fields("TASK") = RIGHT(Me.TASK, 255)
            rst.Fields("PERIORITY") = Me.PERIORITY
            rst.Fields("STATUS") = Me.STATUS
            rst.Fields("STDATE") = Me.STDATE
            rst.Fields("STTIME") = Me.STTIME
            rst.Fields("USERNAME") = Me.USERNAME
            rst.Fields("GR") = Me.GR
            rst.Fields("ENDATE") = Me.ENDATE
            rst.Fields("ENTIME") = Me.ENTIME
        '    RST.Fields("PIC") = Me.PIC
            rst.update
            RST2.MoveNext
        Wend
    End If
    Me.IDNUM = rst.Fields("IDNUM")
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command27_Click()
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    If IsNull(Me.PERSONEL) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " 'گيرنده مشخص نيست....!"
        Exit Sub
    End If
    If IsNull(Me.COMP_COD) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " 'تماس گيرنده مشخص نيست...!"
        Exit Sub
    End If
    If Me.IDNUM = "" Or IsNull(Me.IDNUM) Then
        rst.Open "SELECT     * FROM dbo.TASKS WHERE     (IDNUM = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
    Else
        rst.Open "SELECT     * FROM dbo.TASKS WHERE     (IDNUM = " & Me.IDNUM & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    End If
    rst.Fields("PERSONEL") = Me.PERSONEL
    rst.Fields("COMP_COD") = Me.COMP_COD
    rst.Fields("TASK") = RIGHT(Me.TASK, 255)
    rst.Fields("PERIORITY") = Me.PERIORITY
    rst.Fields("STATUS") = Me.STATUS
    rst.Fields("STDATE") = Me.STDATE
    rst.Fields("STTIME") = Me.STTIME
    rst.Fields("USERNAME") = Me.USERNAME
    rst.Fields("GR") = Me.GR
    rst.Fields("ENDATE") = Me.ENDATE
    rst.Fields("ENTIME") = Me.ENTIME
'    RST.Fields("PIC") = Me.PIC
    rst.update
    RST2.Open "SELECT     * FROM dbo.temp_us WHERE     (tic = 1) AND (USERCOT = " & Forms![BASEKNOW]![USERCOD] & ") and (userco <> " & Me.PERSONEL & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If RST2.RecordCount > 0 Then
        While Not RST2.EOF
            rst.AddNew
            rst.Fields("PERSONEL") = RST2.Fields("userco")
            rst.Fields("COMP_COD") = Me.COMP_COD
            rst.Fields("TASK") = RIGHT(Me.TASK, 255)
            rst.Fields("PERIORITY") = Me.PERIORITY
            rst.Fields("STATUS") = Me.STATUS
            rst.Fields("STDATE") = Me.STDATE
            rst.Fields("STTIME") = Me.STTIME
            rst.Fields("USERNAME") = Me.USERNAME
            rst.Fields("GR") = Me.GR
            rst.Fields("ENDATE") = Me.ENDATE
            rst.Fields("ENTIME") = Me.ENTIME
           ' RST.Fields("PIC") = Me.PIC
            rst.update
            RST2.MoveNext
        Wend
    End If
    Me.IDNUM = ""
    Me.TASK = ""
  '  Me.PIC = Null
    Me.IDNUM = Null
    Me.Refresh
    Me.PERSONEL.SetFocus
End Sub

Private Sub Command28_Click()
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command29_Click()
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    If IsNull(Me.PERSONEL) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " 'گيرنده مشخص نيست....!"
        Exit Sub
    End If
    If IsNull(Me.COMP_COD) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " 'تماس گيرنده مشخص نيست...!"
        Exit Sub
    End If
    If Me.IDNUM = "" Or IsNull(Me.IDNUM) Then
        rst.Open "SELECT     * FROM dbo.TASKS WHERE     (IDNUM = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
    Else
        rst.Open "SELECT     * FROM dbo.TASKS WHERE     (IDNUM = " & Me.IDNUM & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    End If
    rst.Fields("PERSONEL") = Me.PERSONEL
    rst.Fields("COMP_COD") = Me.COMP_COD
    rst.Fields("TASK") = RIGHT(Me.TASK, 255)
    rst.Fields("PERIORITY") = Me.PERIORITY
    rst.Fields("STATUS") = Me.STATUS
    rst.Fields("STDATE") = Me.STDATE
    rst.Fields("STTIME") = Me.STTIME
    rst.Fields("USERNAME") = Me.USERNAME
    rst.Fields("GR") = Me.GR
    rst.Fields("ENDATE") = Me.ENDATE
    rst.Fields("ENTIME") = Me.ENTIME
    'RST.Fields("PIC") = Me.PIC
    rst.update
    RST2.Open "SELECT     * FROM dbo.temp_us WHERE     (tic = 1) AND (USERCOT = " & Forms![BASEKNOW]![USERCOD] & ") and (userco <> " & Me.PERSONEL & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If RST2.RecordCount > 0 Then
        While Not RST2.EOF
            rst.AddNew
            rst.Fields("PERSONEL") = RST2.Fields("userco")
            rst.Fields("COMP_COD") = Me.COMP_COD
            rst.Fields("TASK") = RIGHT(Me.TASK, 255)
            rst.Fields("PERIORITY") = Me.PERIORITY
            rst.Fields("STATUS") = Me.STATUS
            rst.Fields("STDATE") = Me.STDATE
            rst.Fields("STTIME") = Me.STTIME
            rst.Fields("USERNAME") = Me.USERNAME
            rst.Fields("GR") = Me.GR
            rst.Fields("ENDATE") = Me.ENDATE
            rst.Fields("ENTIME") = Me.ENTIME
       '     RST.Fields("PIC") = Me.PIC
            rst.update
            RST2.MoveNext
        Wend
    End If
    Me.IDNUM = rst.Fields("IDNUM")
    Me.taskspevent.Enabled = True
    Me.taskspevent.SetFocus
End Sub

Private Sub Command30_Click()
  Me.USS = ""
    DoCmd.OpenForm "USER_SELECT", acFormDS, , , , acDialog, 1
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
    If BLOCKEDCUST(Me.COMP_COD) Then
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

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Dim rst As New ADODB.Recordset, VS As String
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
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     TOP 1 PERSONEL FROM dbo.MESAGEP WHERE     (USERNAME = N'" & Me.USERNAME & "') ORDER BY IDNUM DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.PERSONEL = rst.Fields("PERSONEL")
    End If
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
Private Sub STATUS_AfterUpdate()
  Me.ENDATE = FARSIDATE(DATE)
  Me.ENTIME = Hour(Now()) * 100 + Minute(Now())
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
