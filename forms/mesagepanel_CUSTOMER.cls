Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command28_Click()
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command33_Click()
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, RETVAL
    If IsNull(Me.PERSONEL) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " 'گيرنده مشخص نيست....!"
        Exit Sub
    End If
    If IsNull(Me.COMP_COD) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " 'تماس گيرنده مشخص نيست...!"
        Exit Sub
    End If
    DoCmd.OpenForm "BUN"
    Forms![BUN].[num].CAPTION = GETUSERNAME(Me.PERSONEL)
    Forms![BUN].Form.Refresh
    If Me.IDNUM = "" Or IsNull(Me.IDNUM) Then
        rst.Open "SELECT     * FROM dbo.MESAGEP WHERE     (IDNUM = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
    Else
        rst.Open "SELECT     * FROM dbo.MESAGEP WHERE     (IDNUM = " & Me.IDNUM & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    End If
    rst.Fields("PERSONEL") = Me.PERSONEL
    rst.Fields("COMP_COD") = Me.COMP_COD
    rst.Fields("PAYAM") = RIGHT("SMS:" & Me.PAYAM, 255)
    rst.Fields("STATUS") = Me.STATUS
    rst.Fields("STDATE") = Me.STDATE
    rst.Fields("STTIME") = Me.STTIME
    rst.Fields("USERNAME") = Me.USERNAME
    rst.update
    Call ersal_sms(Me.COMP_COD, Me.PAYAM & " - " & Me.USERNAME, rst.Fields("IDNUM"), 1, False)
    Me.IDNUM = rst.Fields("IDNUM")
    DoCmd.OpenForm "mesageform", , , , , acDialog, "پيامك ارسال شد....!"
    DoCmd.Close acForm, Me.NAME
    DoCmd.Close acForm, "BUN"
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

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Load()
    Me.PERSONEL = Forms![BASEKNOW]![USERCOD]
    If Me.OpenArgs = "1" Then
        Me.COMP_COD = Forms![main]![TASKS]![COMP_COD]
        Me.PAYAM = Forms![main]![TASKS]![EVENTs]![EVENTs]
        Me.Command33.SetFocus
    End If
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
    If LETSGO("sms") Then
      Me.Command33.Enabled = True
    Else
      Me.Command33.Enabled = False
    End If
End Sub



Private Sub PAYAM_Change()
    If Len(Me.PAYAM.Text) > 255 Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " متن طولاني است ادامه متن را بعداز ذخيره در ذيل بنويسيد"
    End If
End Sub

Private Sub PAYAM_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "zoom", acFormDS, , , , , "1" & Me.PAYAM
End Sub
