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
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenReport "HAVALAH_ANVAR_VROUD", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_RASID_SUB.Locked = True
         Me.INVO_LST_RASID_SUB.Form.AllowDeletions = False
         Me.ESLAH.Enabled = True
    End If
End Sub

Private Sub Command113_Click()
                      On Error Resume Next
     Dim prt As Printer
     Dim rptname As String
     If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
          rptname = "IK" & Val(MID(Forms![BASEKNOW]![OPTIONSS], 11, 2))
     Else
          rptname = "IK1"
     End If
     If rptname = "IK1" Then
    '   DoCmd.OpenReport rptname, acViewPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG = 2"
       For Each prt In Application.Printers
           If Left(prt.DeviceName, 17) = "BTP-2002NP(U) 1 (" Then
              Application.Printer = prt
           End If
       Next
       DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER, , 2
     Else
'       Application.Printer = Nothing
        DoCmd.OpenReport rptname, acPreview, "", "TAG = 1 and NUMBER =" & Me.NUMBER, , 2
     End If
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    DoCmd.RunCommand acCmdSaveRecord
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_RASID_SUB.Locked = True
         Me.INVO_LST_RASID_SUB.Form.AllowDeletions = False
         Me.ESLAH.Enabled = True
    End If
End Sub

Private Sub CUST_NO_DblClick(CANCEL As Integer)
                         On Error Resume Next
       Dim cust_noID As String
    If LETSGO("KALA") Then
          If IsNull(Me![CUST_NO].Text) Then
              Me![CUST_NO].Text = ""
          Else
              cust_noID = Me![CUST_NO].Text
              Me![CUST_NO].Text = ""
          End If
          DoCmd.OpenForm "FCODE_CUSTOMER", , , , , acDialog, "GotoNew"
          Me![CUST_NO].Requery
          If CUST_NO <> "" Then
            Me![CUST_NO].Text = cust_noID
            End If
        Err.Clear
    Else
          DoCmd.OpenForm "mesag", , , , , acDialog, "شما اجازه لازم براي استفاده از اين بخش را نداريد.!"
    End If
End Sub

Private Sub CUST_NO_Exit(CANCEL As Integer)
   If Not IsNull(Me.CUST_NO) Then
        If ISTAF(Me.CUST_NO) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If


End Sub

Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "29"
  End If
  Response = acDataErrContinue
  Err.Clear

End Sub


Private Sub CUST_NO2_Exit(CANCEL As Integer)
   If Not IsNull(Me.CUST_NO2) Then
        If ISTAF(Me.CUST_NO2) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If

End Sub

Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "29"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub
Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub ersalbtn_Click()
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    If Sendbefor(Me.NUMBER, 12) Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين  رسيد  خريد  قبلا ارسال شده آيا مايليد مجددا ارسال شود ؟ "
        If Not Forms![BASEKNOW]![Text44] Then
            Exit Sub
        End If
    End If
    td = Now()
    SHARH = "' رسيد خريد شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
    CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",1," & Me.NUMBER & ",1," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
    If Not IsLoaded("main") Then
        DoCmd.OpenForm "main"
    End If
    Forms![main].Form.Requery
    DoCmd.Close acForm, Me.NAME


End Sub

Private Sub ESLAH_Click()
                  On Error Resume Next
 Dim dt As Date
 Dim rst As New ADODB.Recordset
 If Not IsNull(Me.NUMBER) Then
    rst.Open "select * from head_lst where tag = 12 and NUMBER =  " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or Left(UCurrentUser(), 10) = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
        dt = Now()
        Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 1)", dt, 1)
        Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 1)", dt, 1)
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        If MID(Forms![BASEKNOW]![OPTIONSS], 17, 1) = "5" Then
            Me![INVO_LST_RASID_SUB].Form.AllowAdditions = False
            Me.INVO_LST_RASID_SUB.Form.AllowDeletions = False
        Else
            Me![INVO_LST_RASID_SUB].Form.AllowAdditions = True
            Me.INVO_LST_RASID_SUB.Form.AllowDeletions = True
        End If
        Me.INVO_LST_RASID_SUB.Enabled = True
        Me.INVO_LST_RASID_SUB.Locked = False
      '  If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Me.USER_NAME And Left(UCurrentUser(), 10) <> ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
      '       Me.USER_NAME = UCurrentUser()
      '      DoCmd.RunCommand acCmdSaveRecord
      '  End If
        Me![INVO_LST_RASID_SUB].Form.Refresh
    Else
        If Signed(12, Me.NUMBER) Then
           DoCmd.OpenForm "mesageform", , , , , acDialog, "براي اين حواله فاکتور صادر شده و به امضاء رسيده است .اگر ميخواهيد آنرا اصلاح کنيد بايد امضاء دوم و سوم فاکتور  برداشته شود . به کارتابل ارسال کنيد و براي مدير مالي جهت برداشتن امضا ارسال کنيد"
        Else
          dt = Now()
          Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 1)", dt, 1)
          Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 1)", dt, 1)
          Me.AllowAdditions = True
          Me.AllowDeletions = True
          Me.AllowEdits = True
          If MID(Forms![BASEKNOW]![OPTIONSS], 17, 1) = "5" Then
              Me![INVO_LST_RASID_SUB].Form.AllowAdditions = False
              Me.INVO_LST_RASID_SUB.Form.AllowDeletions = False
          Else
              Me![INVO_LST_RASID_SUB].Form.AllowAdditions = True
              Me.INVO_LST_RASID_SUB.Form.AllowDeletions = True
          End If
          Me.INVO_LST_RASID_SUB.Enabled = True
          Me.INVO_LST_RASID_SUB.Locked = False
       '  If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Me.USER_NAME And Left(UCurrentUser(), 10) <> ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
       '        Me.USER_NAME = UCurrentUser()
       '       DoCmd.RunCommand acCmdSaveRecord
       '   End If
    '      Me.AllowEdits = True
    '      Me.AllowDeletions = False
    '      Me![INVO_LST_RASID_SUB].Form.AllowAdditions = False
    '      Me![INVO_LST_RASID_SUB].Form.Refresh
          DoCmd.OpenForm "mesageform", , , , , acDialog, " براي اين رسيد فاكتور صادر شده است و تغيرات داده شده در فاکتور نيز اعمال ميشود دقت کنيد که بعداز اصلاح رسيد , فاکتور را هم کنترل کنيد ....!"
    '      Me![INVO_LST_RASID_SUB].Form.AllowEdits = False
        End If
    End If
 End If
 Call SETSECURITY(Me.NAME, "RASID", 3)

End Sub

Private Sub Form_AfterUpdate()
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, num As Long

  If MID(Forms![BASEKNOW]![OPTIONSS], 17, 1) = "5" Then
        rst.Open "select * from invo_lst where tag = 23 and NUMBER = " & Me.NUMBER1, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        RST2.Open "select * from invo_lst where tag = 1 and NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount = 0 Then
            While Not rst.EOF
                  RST2.AddNew
                  RST2.Fields("NUMBER") = Me.NUMBER
                  RST2.Fields("TAG") = 1
                  RST2.Fields("CODE") = rst.Fields("CODE")
                  RST2.Fields("ANBAR") = rst.Fields("ANBAR")
                  RST2.Fields("VAHED_K") = rst.Fields("VAHED_K")
                  RST2.Fields("RADAH") = rst.Fields("ID")
                  RST2.update
                  rst.MoveNext
            Wend
            Me.INVO_LST_RASID_SUB.Requery
         End If
  End If
  Set rst = New ADODB.Recordset
 If Not IsNull(Me.NUMBER) And Not IsNull(Me.CUST_NO) Then
    rst.Open "select * from head_lst where tag = 12 and NUMBER =  " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("cust_no") <> Me.CUST_NO Then
            num = rst.Fields("NUMBER")
            rst.Fields("cust_no") = Me.CUST_NO
            rst.update
            rst.Close
            Call GENSANADKHAREED(num, num)
        End If
    End If
 End If
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 1, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
      Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
  If MID(Forms![BASEKNOW]![OPTIONSS], 17, 1) = "5" And IsNull(Me.NUMBER1) Then
      DoCmd.OpenForm "mesageform", , , , , acDialog, " شماره درخواست وارد نشده است ....!"
      CANCEL = True
  Else
    If Me.NUMBER = 0 Then
      rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
          Me.NUMBER = Forms![BASEKNOW]![STHKH]
      Else
          Me.NUMBER = rst.Fields(0) + 1
      End If
    End If
   ' If Me.USER_NAME <> UCurrentUser() Then
  '      Me.USER_NAME = UCurrentUser()
 '   End If
    Me.INVO_LST_RASID_SUB.Locked = False
    Me.INVO_LST_RASID_SUB.Requery
  End If
End Sub

Private Sub Form_Current()
                  On Error Resume Next
 Dim frm As Form
  Dim ghat As Boolean
                  On Error Resume Next
Me.PERSONEL.Visible = True
Me.ERJA.Visible = False
  Set frm = Me![INVO_LST_RASID_SUB].Form
  frm.Requery
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command106.Enabled = True
    Me.Command113.Enabled = True
  Else
    Me.Command106.Enabled = False
    Me.Command113.Enabled = False
  End If
  If Me.NewRecord Then
        Me.INVO_LST_RASID_SUB.Locked = False
        Me.INVO_LST_RASID_SUB.Enabled = True
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        If MID(Forms![BASEKNOW]![OPTIONSS], 17, 1) = "5" Then
            Me![INVO_LST_RASID_SUB].Form.AllowAdditions = False
        Else
            Me![INVO_LST_RASID_SUB].Form.AllowAdditions = True
        End If
        Me![INVO_LST_RASID_SUB].Form.Refresh
  Else
    Dim rst As New ADODB.Recordset
    rst.Open "select * from head_lst where tag = 12 and NUMBER =  " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or Left(UCurrentUser(), 10) = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        If MID(Forms![BASEKNOW]![OPTIONSS], 17, 1) = "5" Then
            Me![INVO_LST_RASID_SUB].Form.AllowAdditions = False
        Else
            Me![INVO_LST_RASID_SUB].Form.AllowAdditions = True
        End If
        Me.INVO_LST_RASID_SUB.Locked = False
        Me.INVO_LST_RASID_SUB.Enabled = True
        Me![INVO_LST_RASID_SUB].Form.Refresh
    Else
        Me.AllowDeletions = False
        Me![INVO_LST_RASID_SUB].Form.AllowAdditions = False
        Me![INVO_LST_RASID_SUB].Form.Refresh
    End If
   rst.Close
 End If
  If Forms![BASEKNOW]![SIGN] Then
        If Me.SGN1 Then
            Me.Command106.Enabled = True
            Me.Command113.Enabled = True
        Else
            Me.Command106.Enabled = False
            Me.Command113.Enabled = False
       End If
  End If
   Err.Clear
   Call SETSECURITY(Me.NAME, "RASID", 3)
   If Me.OKF And Not Me.NewRecord Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_RASID_SUB.Locked = True
        Me.INVO_LST_RASID_SUB.Form.AllowDeletions = False
        Me.ESLAH.Enabled = True
   Else
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.INVO_LST_RASID_SUB.Locked = False
        Me.INVO_LST_RASID_SUB.Form.AllowDeletions = True
        Me.ESLAH.Enabled = False
   End If
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 1, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                     On Error Resume Next
    Dim frm As Form
    Set frm = Me![INVO_LST_RASID_SUB].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين برگه داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس برگه را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, i As Integer, VS As String
  If IsNull(Me.OpenArgs) Then
    Err.Clear
    Me.ServerFilter = ""
    If LETSGO("chartfilter") Then ' اگر محدود به  زير مجموعه خودش هست
        VS = UserOnChart(Forms![BASEKNOW]![USERCOD])
    End If

    'محدود به زير مجموعه خودش باشد
    If LETSGO("chartfilter") Then
        If Me.ServerFilter = "" Then
           Me.ServerFilter = VS
        Else
           Me.ServerFilter = Me.ServerFilter & " and  (" & VS & ")"
        End If
    End If
    'محدود به دپارتمان خودش باشد
    If LETSGO("DEPEMAL") Then
        If Me.ServerFilter = "" Then
           Me.RecordSource = "SELECT TOP (100) PERCENT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 1) and (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") ORDER BY NUMBER "
        Else
           Me.RecordSource = "SELECT TOP (100) PERCENT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 1) and (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") and " & Me.ServerFilter & "ORDER BY NUMBER "
        End If
    End If
Else
   Me.RecordSource = "SELECT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 1)  AND number = " & Me.OpenArgs
End If
  If Forms![BASEKNOW]![UPDDATE] Then
     Me.DATE_N.Locked = False
  Else
     Me.DATE_N.Locked = True
  End If
  DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
  Me.NUMBER.SetFocus
  Call SETSECURITY(Me.NAME, "RASID", 3)
  If Not IsNull(Me.OpenArgs) Then
    Me.ServerFilter = Me.OpenArgs
    Me.Refresh
  End If
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
     If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
        Me.OKF.DefaultValue = True
     Else
        Me.OKF.DefaultValue = False
     End If
If Not LETSGO("ESLAHR") Then
   Me.ESLAH.Visible = False
Else
   Me.ESLAH.Visible = True
End If
If Forms![BASEKNOW]![SIGN] Then
   Me.SGN1.Visible = True
   Me.SGN2.Visible = True
  ' Me.SGN3.Visible = True
'   If GETSIGNER(1, Forms![baseknow]![USERCOD]) > 0 Then
        If IsLoaded("main") Then
            If Not IsNull(Forms![main]![TASKS]![IDNUM]) And (Forms![main]![TASKS]![tg] = 1) Then
              Meidnum = Forms![main]![TASKS]![IDNUM]
             ' ' Me.answerbtn.Visible = True
            End If
        End If
 '  End If
   ' Me.ersalbtn.Visible = True
End If

    Me.PERSONEL.RowSource = GetUserList
    Me.SGN1usid.RowSource = GetUserList
    Me.sgn2usid.RowSource = GetUserList
    Me.sgn3usid.RowSource = GetUserList
    Me.INVO_LST_RASID_SUB.Form.CODE.ColumnWidth = 6660
    If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Me.Width = Me.WindowWidth - Me.WindowWidth * 0.02
           Dim shw
           shw = Me.INVO_LST_RASID_SUB.Width
           Me.INVO_LST_RASID_SUB.Left = 50
           Me.INVO_LST_RASID_SUB.Width = Me.Width - 20
           Me.INVO_LST_RASID_SUB.Form.CODE.ColumnWidth = 6660 + (Me.INVO_LST_RASID_SUB.Width - shw) / 2
           Me.Repaint
           DoCmd.Minimize
           DoCmd.Maximize
        End If
    End If
End Sub

Private Sub INVO_LST_RASID_SUB_Exit(CANCEL As Integer)
Dim frm As Form
  Set frm = Me![INVO_LST_RASID_SUB].Form
  frm.Requery
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command106.Enabled = True
    Me.Command113.Enabled = True
  Else
    Me.Command106.Enabled = False
    Me.Command113.Enabled = False
  End If

End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "FERSAL_SANAD", , , , , acDialog, 2
End Sub

Private Sub NUMBER1_BeforeUpdate(CANCEL As Integer)
                  On Error Resume Next
  If Me.NUMBER > 0 Then
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    rst.Open "select count(NUMBER) as fnum from invo_lst where tag = 1 and NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.Fields("fnum") > 0 Then
        CANCEL = True
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين برگه داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس شماره درخواست جديد را وارد نماييد."
    End If
  End If
End Sub

Private Sub NUMBER1_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "HAVLISTDAR", acFormDS
End Sub

Private Sub NUMBER1_NotInList(NewData As String, Response As Integer)
    If NewData = "++" Or NewData = "+" Then
         DoCmd.OpenForm "HAVLISTDAR", acFormDS
    Else
        If NewData = "--" Or NewData = "-" Then
           DoCmd.OpenForm "HAVLISTDAR", acFormDS, , , , , 1
        End If
    End If
  Response = acDataErrContinue
  Err.Clear


End Sub

Private Sub PERSONEL_AfterUpdate()
    ERJA_Click
End Sub

Private Sub PERSONEL_DblClick(CANCEL As Integer)
    Call PERSONEL_AfterUpdate
End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 1)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & IIf(SGN1, "امضا شد1 ", ":امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",1," & Me.NUMBER & ",1 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'رسيد انبار شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",1," & Me.NUMBER & ",1," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 1)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & IIf(SGN1, "امضا شد1 ", ":امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",1," & Me.NUMBER & ",1 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command106.Enabled = True
        Me.Command113.Enabled = True
    Else
        Me.Command106.Enabled = False
        Me.Command113.Enabled = False
    End If
End Sub

Private Sub SGN2_Click()
                   On Error Resume Next
    If Me.SGN2 Then
        Me.Command106.Enabled = True
    Else
        Me.Command106.Enabled = False
    End If
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 1)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & IIf(SGN2, "امضا شد2 ", "امضا برداشته شد2 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",1," & Me.NUMBER & ",1 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'رسيد انبار شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",1," & Me.NUMBER & ",1," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 1)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & IIf(SGN2, "امضا شد2 ", "امضا برداشته شد2 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",1," & Me.NUMBER & ",1 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command106.Enabled = True
        Me.Command113.Enabled = True
    Else
        Me.Command106.Enabled = False
        Me.Command113.Enabled = False
    End If
End Sub
Private Sub answerbtn_Click()
    DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "1-" & Me.NUMBER & "+" & Meidnum
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub ERJA_Click()
           On Error Resume Next: Meidnum = PERSONELUpdate(1, Me.NUMBER, Me.PERSONEL, "'رسيد شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub

Private Sub PERSONEL_NotInList(NewData As String, Response As Integer)
    Dim rst As New ADODB.Recordset
                  On Error Resume Next
    If IsNumeric(NewData) Then
        rst.Open "select idd from sala_dtl where idd = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.PERSONEL = NewData
        End If
    Else
        DoCmd.OpenForm "SelectUser", acFormDS, , "sal_name like N'%" & CODESAL(NewData) & "%' or sal_name like N'%" & CODESAL(Fixp(NewData)) & "%' or sal_name like N'%" & CODESAL(Fixpi(NewData)) & "%'", , acDialog, 4
    End If
    Response = acDataErrContinue
End Sub
