Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim LETSANAD As Boolean, Meidnum As Double
Private Sub Command100_Click()
                     On Error Resume Next
    Dim frm As Form
'  Call CheckCustomPage("INVOICE_KHAREED", frm.RecordsetClone.RecordCount * 0.9 + 20)
    DoCmd.OpenReport "INVOICE_KHAREED", acPreview, "", "HEAD_LST.NUMBER =" & Me.NUMBER & " AND HEAD_LST.TAG =" & Me.HTAG
End Sub

Private Sub Command106_Click()
                    On Error Resume Next
   Dim frm As Form
    DoCmd.OpenReport "HAVALAH_ANVAR_VROUD", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_RASID_SUB_OTHER.Locked = True
         Me.ESLAH.Enabled = True
    End If
    DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub CUST_NO_AfterUpdate()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  If MID(Forms![BASEKNOW]![OPTIONSS], 19, 1) = 5 Then
    rst.Open "SELECT     hes, CUST_COD FROM dbo.CUST_HESAB WHERE     (hes = N'" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.CUST_KIND = rst.Fields("CUST_COD")
    End If
  End If

End Sub

Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                  On Error Resume Next
  If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "27"
  End If
  Response = acDataErrContinue
  Err.Clear

End Sub

Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                  On Error Resume Next
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "27"
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
            DoCmd.RunSQL ("INSERT INTO dbo.TR_HEAD_LST   (NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, UP_TIME,UP_DATE,OKF,UP_USER_NAME,PC_NAME,IPADD) SELECT     NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,OKF,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'   FROM dbo.HEAD_LST WHERE (NUMBER = " & Me.NUMBER & " ) And (TAG = 24)")
            DoCmd.RunSQL ("INSERT INTO dbo.TR_INVO_LST   (UP_TIME, UP_DATE, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA) SELECT    " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K , FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA FROM dbo.INVO_LST WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 24)")
        End If
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.INVO_LST_RASID_SUB_OTHER.Enabled = True
        Me.INVO_LST_RASID_SUB_OTHER.Locked = False
        If Me.SGN1 Or SGN2 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
            Me.CUST_NO.Locked = True
            Me.INVO_LST_RASID_SUB_OTHER.Locked = True
            Me.DATE_N.Locked = True
            Me.NUMBER.Locked = True
            Me.FNUMCO.Locked = True
            Me.MOLAH.Locked = True
            Me.TAH.Locked = True
        Else
            Me.CUST_NO.Locked = False
            Me.INVO_LST_RASID_SUB_OTHER.Locked = False
            Me.DATE_N.Locked = False
            Me.NUMBER.Locked = False
            Me.FNUMCO.Locked = False
            Me.MOLAH.Locked = False
            Me.TAH.Locked = False
        End If
        Me![INVO_LST_RASID_SUB_OTHER].Form.Refresh
 End If
   Call SETSECURITY(Me.NAME, "RBRFR", 3)

End Sub

Private Sub Form_AfterUpdate()
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 24, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
  If IsNull(Me.CUST_NO) Or Me.CUST_NO = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " مشتري مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If Me.NUMBER = 0 Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER = Forms![BASEKNOW]![STFRB]
    Else
        Me.NUMBER = rst.Fields(0) + 1
    End If
       rst.Close
  End If
  Me.INVO_LST_RASID_SUB_OTHER.Locked = False
  Me.INVO_LST_RASID_SUB_OTHER.Requery
End Sub

Private Sub Form_Current()
                  On Error Resume Next
 Dim frm As Form
  Dim ghat As Boolean
                   On Error Resume Next
    Me.PERSONEL.Visible = True
  Set frm = Me![INVO_LST_RASID_SUB_OTHER].Form
  frm.Requery
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command106.Enabled = True
  Else
    Me.Command106.Enabled = False
  End If
  If Forms![BASEKNOW]![SIGN] And False Then
    If Me.SGN2 Then
        Me.Command106.Enabled = True
    Else
        Me.Command106.Enabled = False
    End If
  End If
  If Me.NewRecord Then
        Me.INVO_LST_RASID_SUB_OTHER.Locked = True
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me![INVO_LST_RASID_SUB_OTHER].Form.AllowAdditions = True
        Me![INVO_LST_RASID_SUB_OTHER].Form.Refresh
  Else
    Dim rst As New ADODB.Recordset
    rst.Open "select * from head_lst where tag = 25 and NUMBER =  " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or Left(UCurrentUser(), 10) = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me![INVO_LST_RASID_SUB_OTHER].Form.AllowAdditions = True
        Me.INVO_LST_RASID_SUB_OTHER.Locked = False
        Me![INVO_LST_RASID_SUB_OTHER].Form.Refresh
    Else
        Me.INVO_LST_RASID_SUB_OTHER.Locked = False
        Me.AllowDeletions = False
        Me![INVO_LST_RASID_SUB_OTHER].Form.AllowAdditions = False
        Me![INVO_LST_RASID_SUB_OTHER].Form.Refresh
    End If
   rst.Close
 End If
  If Forms![BASEKNOW]![SIGN] Then
        Me.SGN1.Visible = True
        Me.SGN2.Visible = True
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT    SGN0124,SGN0224,SGN0324 FROM dbo.SIGN WHERE     USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If rst.Fields("SGN0124") Then
              Me.SGN1.Enabled = True
            Else
              Me.SGN1.Enabled = False
            End If
            If rst.Fields("SGN0224") Then
              Me.SGN2.Enabled = True
            Else
              Me.SGN2.Enabled = False
            End If
        End If
        Me.NUMBER.Locked = False
        Me.DATE_N.Locked = False
        Me.FNUMCO.Locked = False
        Me.MOLAH.Locked = False
        Me.TAH.Locked = False
        If Me.SGN2 Then
            Me.SGN1.Enabled = False
            Me.AllowDeletions = False
            Me.INVO_LST_RASID_SUB_OTHER.Locked = True
            Me.INVO_LST_RASID_SUB_OTHER.Form.AllowDeletions = False
            Me.NUMBER.Locked = True
            Me.DATE_N.Locked = True
            Me.FNUMCO.Locked = True
            Me.MOLAH.Locked = True
            Me.TAH.Locked = True
        End If
   End If
   Err.Clear
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_RASID_SUB_OTHER.Locked = True
         Me.ESLAH.Enabled = True
    End If
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 24, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                     On Error Resume Next
    Dim frm As Form
    Set frm = Me![INVO_LST_RASID_SUB_OTHER].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين برگه داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس برگه را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  If Forms![BASEKNOW]![UPDDATE] Then
     Me.DATE_N.Locked = False
  Else
     Me.DATE_N.Locked = True
  End If
  If IsNull(Me.ServerFilter) Then
    DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
  End If
  Me.NUMBER.SetFocus
  Call SETSECURITY(Me.NAME, "RBRFR", 3)
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub M_NAGHD_KeyPress(KeyAscii As Integer)
                    On Error Resume Next
sendkeyu (KeyAscii)
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    DoCmd.Maximize
    Forms![BASEKNOW]![hhwin] = Me.hWnd
    If Not LETSGO("ESLAHA") Then
       Me.ESLAH.Visible = False
    Else
       Me.ESLAH.Visible = True
    End If
    Me.PERSONEL.RowSource = GetUserList
    Me.SGN1usid.RowSource = GetUserList
    Me.sgn2usid.RowSource = GetUserList
    Me.sgn3usid.RowSource = GetUserList
    Me.INVO_LST_RASID_SUB_OTHER.Form.CODE.ColumnWidth = 5610
    Me.INVO_LST_RASID_SUB_OTHER.Width = 14160
    If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Me.Width = Me.WindowWidth - Me.WindowWidth * 0.02
           Dim shw
           shw = 14160
           Me.INVO_LST_RASID_SUB_OTHER.Left = 50
           Me.INVO_LST_RASID_SUB_OTHER.Width = Me.Width - 20
           Me.INVO_LST_RASID_SUB_OTHER.Form.CODE.ColumnWidth = 5610 + (Me.INVO_LST_RASID_SUB_OTHER.Width - shw) / 2
           Me.Repaint
        End If
    End If
End Sub



Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 24)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",24," & Me.NUMBER & ",24 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'ساير رسيد انبار شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",24," & Me.NUMBER & ",24," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 24)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",24," & Me.NUMBER & ",24 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command106.Enabled = True
    Else
        Me.Command106.Enabled = False
    End If

End Sub

Private Sub SGN2_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 24)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",24," & Me.NUMBER & ",24 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'ساير رسيد انبار شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",24," & Me.NUMBER & ",24," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 24)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",24," & Me.NUMBER & ",24 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command106.Enabled = True
    Else
        Me.Command106.Enabled = False
    End If

End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 24)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",24," & Me.NUMBER & ",24 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'ساير رسيد انبار شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",24," & Me.NUMBER & ",24," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 24)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",24," & Me.NUMBER & ",24 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command106.Enabled = True
    Else
        Me.Command106.Enabled = False
    End If
End Sub


Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(24, Me.NUMBER, Me.PERSONEL, "'ساير رسيد انبار شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
        DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub
