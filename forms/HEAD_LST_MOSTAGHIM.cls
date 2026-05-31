Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Private Sub SANAD()
                    On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
'  If Forms![baseknow]![SANAD] = 1 Then
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, SHSH As String
  Dim max_ns, MABL_CHK, JAMF, JAMCH, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, MKOL, MMOIN, MTAF, MTAF2, MTAF3, MTAF4
    If Not IsNull(Me.CUST_NO) Then
         Call GETTAF3(Me.CUST_NO, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
    End If
    If Not IsNull(Me.MOIN_HAZ) Then
         Call GETTAF3(Me.MOIN_HAZ, MKOL, MMOIN, MTAF, MTAF2, MTAF3, MTAF4)
    End If
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

  DoCmd.OpenForm "BUN"
 SHSH = IIf(Forms![BASEKNOW]![SNDKH], Left(" فاكتورهاي خريد كالاي مصرف مستقيم مورخ " & Format(Me.DATE_N, "####/##/##"), 255), Left("فاكتور خريد كالاي مصرف مستقيم  شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255))
If Not IsNull(Me.N_S) Then
      CurrentProject.Connection.Execute ("UPDATE DEED_HED SET DATE_S = " & Me.DATE_N & ",SHARH_S = '" & SHSH & "',GHATEI = 0,NO_S = 15,OKF=-1,USER_NAME ='" & Me.USER_NAME & "' WHERE N_S =" & Me.N_S): DoCmd.OpenForm "BUN"
      max_ns = Me.N_S
Else
      If Forms![BASEKNOW]![SNDKH] Then
          Set SHRST = CurrentProject.Connection.Execute("SELECT N_S,NO_S,DATE_S FROM DEED_HED WHERE NO_S = 15 AND DATE_S = " & Me.DATE_N)
          If SHRST.RecordCount <= 0 Then
              max_ns = Createsanad(Me.DATE_N, SHSH, 0, 15, -1, Me.USER_NAME)
          Else
              max_ns = SHRST.Fields("n_S")
              Forms![BUN].Form.Refresh
              Me.N_S = max_ns
              CurrentProject.Connection.Execute ("UPDATE DEED_HED SET DATE_S = " & Me.DATE_N & ",SHARH_S = '" & SHSH & "',GHATEI = 0,NO_S = 15,OKF=-1,USER_NAME ='" & Me.USER_NAME & "' WHERE N_S = " & max_ns): DoCmd.OpenForm "BUN"
              Forms![BUN].Form.Refresh
          End If
      Else
          max_ns = Createsanad(Me.DATE_N, SHSH, 0, 15, -1, Me.USER_NAME)
      End If
End If
  If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
     Me.N_S = max_ns
  End If
  JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER & " ) AND ((INVO_LST.TAG)=15))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
       JAMF = JST.Fields(0)
  Else
     JAMF = 0
  End If
  JST.Close
 SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
  If JAMF <> 0 Then
    SDRST.AddNew 'مركز هزينه
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = MKOL
    SDRST.Fields("HES_M") = MMOIN
    SDRST.Fields("HES_T") = MTAF
    SDRST.Fields("HES_T2") = MTAF2
    SDRST.Fields("HES_T3") = MTAF3
    SDRST.Fields("HES_T4") = MTAF4
    SDRST.Fields("hes") = Me.MOIN_HAZ
    SDRST.Fields("SHARH") = RIGHT("رسيدش. " & Me.NUMBER & "-" & Me.FNUMCO & " " & Me.MOLAH & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BED") = JAMF
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
End If
  If JAMF <> 0 Then
        SDRST.AddNew 'كل بستانكاري شخص بابت فاكتور
        SDRST.Fields("N_S") = max_ns
       SDRST.Fields("HES_K") = CKOL
       SDRST.Fields("HES_M") = CMOIN
       SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = Me.CUST_NO2
        SDRST.Fields("SHARH") = RIGHT("رسيدش. " & Me.NUMBER & "-" & Me.FNUMCO & " " & Me.MOLAH & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
        SDRST.Fields("BES") = JAMF
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.Fields("RADIF") = Me.NUMBER
        SDRST.update
  End If

    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(CUST_NO) Then
           rst.Close
           Set rst = New ADODB.Recordset
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES  = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Then
                  Me.MANDAH = 0
                Else
                  Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
                End If
         Else
             Me.MANDAH = "مسدود است"
         End If
    End If
 DoCmd.Close acForm, "bun"
End Sub
Private Sub Command100_Click()
                      On Error Resume Next
    DoCmd.OpenReport "INVOICE_KHAREED", acPreview, "", "HEAD_LST.NUMBER =" & Me.NUMBER & " AND HEAD_LST.TAG =" & Me.HTAG
End Sub


Private Sub Command106_Click()
                       On Error Resume Next
    DoCmd.OpenReport "RASID_MOSTAGHIM_LARG", acPreview, "", "NUMBER =" & Me.NUMBER
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_MO_SUB.Locked = True
         Me.ESLAH.Enabled = True
    End If
    DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub Command108_Click()
                       On Error Resume Next
    DoCmd.OpenReport "RASID_MOSTAGHIM", acPreview, "", "NUMBER =" & Me.NUMBER
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_MO_SUB.Locked = True
         Me.ESLAH.Enabled = True
    End If
    DoCmd.RunCommand acCmdSaveRecord
End Sub


Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "04"
  End If
  Response = acDataErrContinue
  Err.Clear

End Sub


Private Sub CUST_NO2_AfterUpdate()
                     On Error Resume Next
  Me.CUST_NO.Requery
End Sub

Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "04"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub

Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub ESLAH_Click()
                  On Error Resume Next
 Dim dt As Date
 If Not IsNull(Me.NUMBER) Then
        dt = Now()
        If Forms![BASEKNOW]![TRANSF] Then
            DoCmd.RunSQL ("INSERT INTO dbo.TR_HEAD_LST   (NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, UP_TIME,UP_DATE,OKF,UP_USER_NAME,PC_NAME,IPADD) SELECT     NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,OKF,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'   FROM dbo.HEAD_LST WHERE (NUMBER = " & Me.NUMBER & " ) And (TAG = 15)")
            DoCmd.RunSQL ("INSERT INTO dbo.TR_INVO_LST   (UP_TIME, UP_DATE, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA) SELECT    " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K , FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA FROM dbo.INVO_LST WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 15)")
        End If
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.INVO_LST_MO_SUB.Enabled = True
        Me.INVO_LST_MO_SUB.Locked = False
      '  If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Me.USER_NAME And Left(UCurrentUser(), 10) <> ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
      '       Me.USER_NAME = UCurrentUser()
      '      DoCmd.RunCommand acCmdSaveRecord
      '  End If
        Me![INVO_LST_MO_SUB].Form.Refresh
 End If
   Call SETSECURITY(Me.NAME, "MOS", 3)

End Sub

Private Sub Form_AfterUpdate()
                      On Error Resume Next
  Call SANAD
  LETSANAD = False
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next
  If IsNull(Me.CUST_NO) Or Me.CUST_NO = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, "نام خريدار مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If IsNull(Me.MOIN_HAZ) Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, "مركز هزينه مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If Me.NUMBER = 0 Then
    Dim rst As New ADODB.Recordset
 '***************************************************************
      '  CurrentProject.Connection.Execute ("set transaction isolation level SERIALIZABLE"): DoCmd.OpenForm "BUN"
      '  CurrentProject.Connection.Execute ("begin transaction")
 '***************************************************************
    Set rst = CurrentProject.Connection.Execute("SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))")
'    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER = Forms![BASEKNOW]![STMO]
    Else
        Me.NUMBER = rst.Fields(0) + 1
    End If
       rst.Close
  End If
 '***************************************************************
     '    CurrentProject.Connection.Execute ("COMMIT TRANSACTION"): DoCmd.Close acForm, "bun"
 '***************************************************************

  Me.INVO_LST_MO_SUB.Locked = False
  Me.INVO_LST_MO_SUB.Requery
End Sub

Private Sub Form_Current()
                     On Error Resume Next
  Dim frm As Form
  Dim ghat As Boolean
                      On Error Resume Next
  Set frm = Me![INVO_LST_MO_SUB].Form
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command108.Enabled = True
    Me.Command106.Enabled = True
  Else
    Me.Command108.Enabled = False
    Me.Command106.Enabled = False
  End If
  If IsNull(Me.N_S) Then
     Me.AllowAdditions = True
     Me.AllowDeletions = True
     Me.AllowEdits = True
     Me.INVO_LST_MO_SUB.Locked = False
     Me![INVO_LST_MO_SUB].Form.Refresh
     Me.lsanad.ForeColor = 65535
     Me.INVO_LST_MO_SUB.Form.AllowDeletions = True
  Else
    Dim rst As New ADODB.Recordset
    rst.Open "DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "N_S = " & Me.N_S
    Set rst = rst
    If rst.RecordCount = 0 Then
    Else
        If rst.Fields("GHATEI") Then
            ghat = True
            Me.AllowDeletions = False
            Me.AllowEdits = False
            Me.INVO_LST_MO_SUB.Locked = True
            Me![INVO_LST_MO_SUB].Form.Refresh
            Me.INVO_LST_MO_SUB.Form.AllowDeletions = False
            Me.lsanad.ForeColor = 125
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_MO_SUB.Locked = False
            Me![INVO_LST_MO_SUB].Form.Refresh
            Me.INVO_LST_MO_SUB.Form.AllowDeletions = True
            Me.lsanad.ForeColor = 65535
        End If
     End If
   rst.Close
   Set rst = New ADODB.Recordset

 End If
    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(CUST_NO) Then
           rst.Close
           Set rst = New ADODB.Recordset
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES  = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Then
                  Me.MANDAH = 0
                Else
                  Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
                End If
         Else
             Me.MANDAH = "مسدود است"
         End If
    End If
   If Me.NewRecord Then
        Me.INVO_LST_MO_SUB.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
   Else
        If Not ghat Then
            Me.INVO_LST_MO_SUB.Locked = False
            Me![INVO_LST_MO_SUB].Form.Refresh
        Else
            Me.INVO_LST_MO_SUB.Locked = True
        End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.CUST_NO.TAG = Me.CUST_NO
                Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
   End If
   Err.Clear
   Call SETSECURITY(Me.NAME, "MOS", 3)
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_MO_SUB.Locked = True
         Me.ESLAH.Enabled = True
    End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                       On Error Resume Next
    Dim frm As Form
    Set frm = Me![INVO_LST_MO_SUB].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس فاكتور را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub

Private Sub Form_Load()
                     On Error Resume Next
  DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
  Me.FNUMCO.SetFocus
  Call SETSECURITY(Me.NAME, "MOS", 3)
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub M_NAGHD_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub


Private Sub Form_Open(CANCEL As Integer)
       Forms![BASEKNOW]![hhwin] = Me.hWnd
     If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
        Me.OKF.DefaultValue = True
     Else
        Me.OKF.DefaultValue = False
     End If
     If Not LETSGO("ESLAHM") Then
       Me.ESLAH.Visible = False
    Else
       Me.ESLAH.Visible = True
    End If
End Sub

Private Sub INVO_LST_MO_SUB_Exit(CANCEL As Integer)
                   On Error Resume Next
  If LETSANAD Then
    Call SANAD
    LETSANAD = False
 End If
End Sub

Private Sub MANDAH_DblClick(CANCEL As Integer)
                   On Error Resume Next
  If Not IsNull(Me.CUST_NO) Then
    Call SANAD
  End If

End Sub

Private Sub MOIN_HAZ_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If MOIN_HAZ.Text = "-" Or MOIN_HAZ.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "10"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub

Private Sub MOIN_HAZ1_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If MOIN_HAZ1.Text = "-" Or MOIN_HAZ1.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "10"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                      On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
