Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
Private Sub Combo34_AfterUpdate()
  Me.Combo36 = Me.Combo34
  Me.Combo36.Requery
End Sub

Private Sub Combo34_Exit(CANCEL As Integer)
    Call Combo34_AfterUpdate
    If Not IsNull(Me.Combo34) Then
        If ISTAF(Me.Combo34) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If
End Sub

Private Sub Combo34_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If Me.Combo34.Text = "-" Or Me.Combo34.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , acDialog, "49"
  End If
  Response = acDataErrContinue
End Sub

Private Sub Combo36_AfterUpdate()
   Dim KOL, MOIN, taf, TAF2, TAF3, TAF4
   Me.Combo34 = Me.Combo36
   Me.Combo34.Requery
End Sub

Private Sub Combo36_Exit(CANCEL As Integer)
    If Not IsNull(Me.Combo34) Then
        If ISTAF(Me.Combo34) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If

End Sub

Private Sub Command120_Click()
               On Error Resume Next
  Dim sql   As String
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, PATH As String, SORTT As String, i As Integer, MAN As Double
  If IsNull(Me.HTAF2) Then
    Me.HTTAF = Me.HKOL & "-" & Me.HMOIN & "-" & Me.HTAF
  Else
    Me.HTTAF = Me.HKOL & "-" & Me.HMOIN & "-" & Me.HTAF & "-" & Me.HTAF2
  End If
  If IsNull(Me.DT1) Then
    DT1 = CLng(Forms![BASEKNOW]![YEA] & "0101")
  End If
  If IsNull(Me.DT2) Then
    DT1 = "99990101"
  End If
  If IsNull(Me.SNDNUM1) Then
    Me.SNDNUM1 = 0
  End If
  If IsNull(Me.SNDNUM2) Then
    Me.SNDNUM2 = 929292929
  End If


If IsNull(Me.HKOL) Or IsNull(Me.HMOIN) Or IsNull(Me.DT1) Or IsNull(Me.DT2) Then
  DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
  Exit Sub
End If
If BLOCKEDMK(Me.Combo34) Then
   DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
   Exit Sub
End If
If BLOCKEDMK(Me.TOHES) Then
   DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
   Exit Sub
End If
rst.Open "select n_s   FROM dbo.DEED_DTL WHERE     (HES = N'" & Me.Combo34 & "') and n_s in (select n_s from deed_hed where date_s between " & Me.DT1 & " and " & Me.DT2 & ") and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "تعداد سطر : " & rst.RecordCount
    CurrentProject.Connection.Execute ("UPDATE    DEED_DTL SET   HES = '" & Me.TOHES & "', HES_K = " & Me.HKOL & ", HES_M = " & Me.HMOIN & "  , HES_T = " & Me.HTAF & " , HES_T2 = " & IIf(IsNull(Me.HTAF2), "NULL", Me.HTAF2) & ", HES_T3 =  " & IIf(IsNull(Me.HTAF3), "NULL", Me.HTAF3) & ", HES_T4 = " & IIf(IsNull(Me.HTAF4), "NULL", Me.HTAF4) & "  FROM dbo.DEED_DTL WHERE     (HES = N'" & Me.Combo34 & "') and n_s in (select n_s from deed_hed where date_s between " & Me.DT1 & " and " & Me.DT2 & ") and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2)
End If
Set rst = New ADODB.Recordset
rst.Open "select ID   FROM dbo.PGET_LST WHERE     (FHES = N'" & Me.Combo34 & "') OR (THES = N'" & Me.Combo34 & "') and id in (select id from pget_hed where date between " & Me.DT1 & " and " & Me.DT2 & " and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2 & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در اسناد حسابداري انجام شد"
    CurrentProject.Connection.Execute ("UPDATE    PGET_LST SET   FHES = '" & Me.TOHES & "', FHES_K = " & Me.HKOL & ", FHES_M = " & Me.HMOIN & "  , FHES_T = " & Me.HTAF & " , FHES_T2 = " & IIf(IsNull(Me.HTAF2), "NULL", Me.HTAF2) & ", FHES_T3 =  " & IIf(IsNull(Me.HTAF3), "NULL", Me.HTAF3) & ", FHES_T4 = " & IIf(IsNull(Me.HTAF4), "NULL", Me.HTAF4) & "  FROM dbo.PGET_LST WHERE     (FHES = N'" & Me.Combo34 & "') and id in (select id from pget_hed where date between " & Me.DT1 & " and " & Me.DT2 & " and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2 & ")")
    CurrentProject.Connection.Execute ("UPDATE    PGET_LST SET   THES = '" & Me.TOHES & "', THES_K = " & Me.HKOL & ", THES_M = " & Me.HMOIN & "  , THES_T = " & Me.HTAF & " , THES_T2 = " & IIf(IsNull(Me.HTAF2), "NULL", Me.HTAF2) & ", THES_T3 =  " & IIf(IsNull(Me.HTAF3), "NULL", Me.HTAF3) & ", THES_T4 = " & IIf(IsNull(Me.HTAF4), "NULL", Me.HTAF4) & "  FROM dbo.PGET_LST WHERE     (THES = N'" & Me.Combo34 & "') and id in (select id from pget_hed where date between " & Me.DT1 & " and " & Me.DT2 & " and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2 & ")")
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در خزانه داري  انجام شد"
End If
Set rst = New ADODB.Recordset
rst.Open "select NUMBER   FROM dbo.HEAD_LST WHERE     (CUST_NO = N'" & Me.Combo34 & "') and n_s in (select n_s from deed_hed where date_s between " & Me.DT1 & " and " & Me.DT2 & ") and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    HEAD_LST SET   CUST_NO = '" & Me.TOHES & "' FROM dbo.HEAD_LST WHERE     (CUST_NO = N'" & Me.Combo34 & "') and n_s in (select n_s from deed_hed where date_s between " & Me.DT1 & " and " & Me.DT2 & ") and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2)
End If
Set rst = New ADODB.Recordset
rst.Open "select NUMBER   FROM dbo.HEAD_LST WHERE     (MOIN_VAR = N'" & Me.Combo34 & "') and n_s in (select n_s from deed_hed where date_s between " & Me.DT1 & " and " & Me.DT2 & ") and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    HEAD_LST SET    MOIN_VAR = '" & Me.TOHES & "' FROM dbo.HEAD_LST WHERE     ( MOIN_VAR = N'" & Me.Combo34 & "') and n_s in (select n_s from deed_hed where date_s between " & Me.DT1 & " and " & Me.DT2 & ") and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2)
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در حساب واريزي فاکتورها   انجام شد"
End If
Set rst = New ADODB.Recordset
rst.Open "select NUMBER   FROM dbo.HEAD_LST WHERE     (MOIN_HAV = N'" & Me.Combo34 & "') and n_s in  (select n_s from deed_hed where date_s between " & Me.DT1 & " and " & Me.DT2 & ") and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    HEAD_LST SET    MOIN_HAV = '" & Me.TOHES & "' FROM dbo.HEAD_LST WHERE     ( MOIN_HAV = N'" & Me.Combo34 & "') and n_s in (select n_s from deed_hed where date_s between " & Me.DT1 & " and " & Me.DT2 & ") and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2)
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در حساب حواله واريزي فاکتورها   انجام شد"
End If
Set rst = New ADODB.Recordset
rst.Open "select NUMBER   FROM dbo.HEAD_LST WHERE     (MOIN_HAZ = N'" & Me.Combo34 & "') and n_s in (select n_s from deed_hed where date_s between " & Me.DT1 & " and " & Me.DT2 & ") and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    HEAD_LST SET    MOIN_HAZ = '" & Me.TOHES & "' FROM dbo.HEAD_LST WHERE     ( MOIN_HAZ = N'" & Me.Combo34 & "') and n_s in (select n_s from deed_hed where date_s between " & Me.DT1 & " and " & Me.DT2 & ") and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2)
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در حساب هزينه فاکتورها   انجام شد"
End If
Set rst = New ADODB.Recordset
rst.Open "select NUMBER   FROM dbo.HEAD_LST WHERE     (HMBAA = N'" & Me.Combo34 & "') and n_s in (select n_s from deed_hed where date_s between " & Me.DT1 & " and " & Me.DT2 & ") and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    HEAD_LST SET    HMBAA = '" & Me.TOHES & "' FROM dbo.HEAD_LST WHERE     ( HMBAA = N'" & Me.Combo34 & "') and n_s in (select n_s from deed_hed where date_s between " & Me.DT1 & " and " & Me.DT2 & ") and  n_s between " & Me.SNDNUM1 & " and  " & Me.SNDNUM2)
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در حساب ماليات بر ارزش افزوده فاکتورها   انجام شد"
End If
Set rst = New ADODB.Recordset
rst.Open "select NUMBER   FROM dbo.HEAD_LST WHERE    (CUST_NO = N'" & Me.Combo34 & "') and  (date_n between " & Me.DT1 & " and " & Me.DT2 & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    HEAD_LST SET    CUST_NO = '" & Me.TOHES & "' FROM dbo.HEAD_LST WHERE  (CUST_NO = N'" & Me.Combo34 & "') and  (date_n between " & Me.DT1 & " and " & Me.DT2 & ")")
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در حساب  فاکتورها   انجام شد"
End If

Set rst = New ADODB.Recordset
rst.Open "select N_SERI   FROM dbo.PAY_GETD WHERE     (HES1 = N'" & Me.Combo34 & "') and (date between " & Me.DT1 & " and " & Me.DT2 & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    PAY_GETD SET   HES1 = '" & Me.TOHES & "', N_KOL = " & Me.HKOL & ", N_MOIN = " & Me.HMOIN & "  , N_TAF = " & Me.HTAF & "   FROM dbo.PAY_GETD WHERE     (HES1 = N'" & Me.Combo34 & "') and (date between " & Me.DT1 & " and " & Me.DT2 & ")")
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در واگذار به حساب چکهاي دريافتي   انجام شد"
End If

Set rst = New ADODB.Recordset
rst.Open "select N_SERI   FROM dbo.PAY_GETD WHERE     (HES2 = N'" & Me.Combo34 & "') and (date between " & Me.DT1 & " and " & Me.DT2 & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    PAY_GETD SET   HES2 = '" & Me.TOHES & "', N_KOL2 = " & Me.HKOL & ", N_MOIN2 = " & Me.HMOIN & "  , N_TAF2 = " & Me.HTAF & "   FROM dbo.PAY_GETD WHERE     (HES2 = N'" & Me.Combo34 & "') and (date between " & Me.DT1 & " and " & Me.DT2 & ")")
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در برگشت به حساب چکهاي دريافتي   انجام شد"
End If

Set rst = New ADODB.Recordset
rst.Open "select N_SERI   FROM dbo.PAY_GETD WHERE     (HES3 = N'" & Me.Combo34 & "') and (date between " & Me.DT1 & " and " & Me.DT2 & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    PAY_GETD SET   HES3 = '" & Me.TOHES & "', N_KOL3 = " & Me.HKOL & ", N_MOIN3 = " & Me.HMOIN & "  , N_TAF3 = " & Me.HTAF & "   FROM dbo.PAY_GETD WHERE     (HES3 = N'" & Me.Combo34 & "') and (date between " & Me.DT1 & " and " & Me.DT2 & ")")
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در وصول به حساب چکهاي دريافتي   انجام شد"
End If

Set rst = New ADODB.Recordset
rst.Open "select N_SERI   FROM dbo.PAY_GETP WHERE     (HES1 = N'" & Me.Combo34 & "') and (date between " & Me.DT1 & " and " & Me.DT2 & ") and (date between " & Me.DT1 & " and " & Me.DT2 & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    PAY_GETP SET   HES1 = '" & Me.TOHES & "', N_KOL = " & Me.HKOL & ", N_MOIN = " & Me.HMOIN & "  , N_TAF = " & Me.HTAF & "   FROM dbo.PAY_GETP WHERE     (HES1 = N'" & Me.Combo34 & "') and (date between " & Me.DT1 & " and " & Me.DT2 & ")")
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در پرداخت از حساب چکهاي پرداختي   انجام شد"
End If

Set rst = New ADODB.Recordset
rst.Open "select N_SERI   FROM dbo.PAY_GETP WHERE     (HES2 = N'" & Me.Combo34 & "') and (date between " & Me.DT1 & " and " & Me.DT2 & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    PAY_GETP SET   HES2 = '" & Me.TOHES & "', N_KOL2 = " & Me.HKOL & ", N_MOIN2 = " & Me.HMOIN & "  , N_TAF2 = " & Me.HTAF & "   FROM dbo.PAY_GETP WHERE     (HES2 = N'" & Me.Combo34 & "') and (date between " & Me.DT1 & " and " & Me.DT2 & ")")
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در برگشت به حساب چکهاي پرداختي   انجام شد"
End If

Set rst = New ADODB.Recordset
rst.Open "select N_SERI   FROM dbo.PAY_GETP WHERE     (HES3 = N'" & Me.Combo34 & "') and (date between " & Me.DT1 & " and " & Me.DT2 & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    PAY_GETP SET   HES3 = '" & Me.TOHES & "', N_KOL3 = " & Me.HKOL & ", N_MOIN3 = " & Me.HMOIN & "  , N_TAF3 = " & Me.HTAF & "   FROM dbo.PAY_GETP WHERE     (HES3 = N'" & Me.Combo34 & "') and (date between " & Me.DT1 & " and " & Me.DT2 & ")")
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در وصول از حساب چکهاي پرداختي   انجام شد"
End If


Set rst = New ADODB.Recordset

rst.Open "SELECT     dbo.INVO_LST.N_RASID FROM    dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG WHERE     (dbo.HEAD_LST.DATE_N BETWEEN " & Me.DT1 & " and " & Me.DT2 & ")  and   (N_RASID = N'" & Me.Combo34 & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    INVO_LST SET   N_RASID = '" & Me.TOHES & " '  FROM  dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG WHERE     (dbo.HEAD_LST.DATE_N BETWEEN " & Me.DT1 & " and " & Me.DT2 & ")  and   (N_RASID = N'" & Me.Combo34 & "')")
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در محل مصرف حواله خروج ساير انجام شد"
End If

Set rst = New ADODB.Recordset
rst.Open "select POWNER   FROM dbo.PTAMIRAT WHERE     (POWNER = N'" & Me.Combo34 & "') and  (PINDATE between " & Me.DT1 & " and " & Me.DT2 & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    PTAMIRAT SET   POWNER = '" & Me.TOHES & " '  FROM dbo.PTAMIRAT WHERE      (POWNER = N'" & Me.Combo34 & "') and  (PINDATE between " & Me.DT1 & " and " & Me.DT2 & ")")
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در نام مالک دستگاه در تعميرات انجام شد"
End If

Set rst = New ADODB.Recordset
rst.Open "select CUST_NO   FROM dbo.VISITOR_DTL WHERE     (CUST_NO = N'" & Me.Combo34 & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    CurrentProject.Connection.Execute ("UPDATE    VISITOR_DTL SET   CUST_NO = '" & Me.TOHES & " '  FROM          dbo.HEAD_LST INNER JOIN    dbo.VISITOR_DTL ON dbo.HEAD_LST.NUMBER = dbo.VISITOR_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.VISITOR_DTL.TAG WHERE      (dbo.VISITOR_DTL.CUST_NO = N'" & Me.Combo34 & "') and  (DATE_N between " & Me.DT1 & " and " & Me.DT2 & ")")
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "جايگزاري حساب در نام مالک دستگاه در تعميرات انجام شد"
End If



End Sub




Private Sub Command8_Click()
                     On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub DT1_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
                    On Error Resume Next: CANCEL = CHEKDATEM(DT1, False)
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
  DoCmd.restore
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub HTAF_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If Me.HTAF.Text = "-" Or Me.HTAF.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "05"
  End If
Response = acDataErrContinue
End Sub

Private Sub TOHES_AfterUpdate()
   Dim KOL, MOIN, taf, TAF2, TAF3, TAF4
   Me.TOHES2 = Me.TOHES
  If Not IsNull(Me.TOHES) Then
    Call GETTAF3(Me.TOHES, KOL, MOIN, taf, TAF2, TAF3, TAF4)
    Me.HKOL = KOL
    Me.HMOIN = MOIN
    Me.HTAF = taf
    Me.HTAF2 = TAF2
    Me.HTAF3 = TAF3
    Me.HTAF4 = TAF4
  End If
   Me.TOHES2.Requery

End Sub


Private Sub TOHES_Exit(CANCEL As Integer)
    Call TOHES_AfterUpdate
    If Not IsNull(Me.TOHES) Then
        If ISTAF(Me.TOHES) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If
End Sub

Private Sub TOHES_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If Me.TOHES.Text = "-" Or Me.TOHES.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , acDialog, "50"
  End If
  Response = acDataErrContinue

End Sub

Private Sub TOHES2_AfterUpdate()
   Dim KOL, MOIN, taf, TAF2, TAF3, TAF4
   Me.TOHES = Me.TOHES2
  If Not IsNull(Me.TOHES) Then
    Call GETTAF3(Me.TOHES, KOL, MOIN, taf, TAF2, TAF3, TAF4)
    Me.HKOL = KOL
    Me.HMOIN = MOIN
    Me.HTAF = taf
    Me.HTAF2 = TAF2
    Me.HTAF3 = TAF3
    Me.HTAF4 = TAF4
  End If
   Me.TOHES.Requery

End Sub

Private Sub TOHES2_Exit(CANCEL As Integer)
    If Not IsNull(Me.TOHES) Then
        If ISTAF(Me.TOHES) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If
End Sub
