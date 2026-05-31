Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim DTCHK As Boolean, DTDT As String
Dim Meidnum As Double
Private Sub SANAD()
                    On Error Resume Next
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset
  Dim max_ns, MABL_CHK As Double, SHART, SHRH As String
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  SHRH = "خزانه داري شماره " & Me.ID & "مورخ " & Format(Me.DATE, "####/##/##")
    If Not IsNull(Me.N_S) Then
          CurrentProject.Connection.Execute ("UPDATE DEED_HED SET DATE_S = " & Me.DATE & ",SHARH_S = '" & SHRH & "',GHATEI = 0,NO_S = 5,OKF=-1,USER_NAME ='" & Me.USER_NAME & "' WHERE N_S =" & Me.N_S): DoCmd.OpenForm "BUN"
          max_ns = Me.N_S
    Else
          max_ns = Createsanad(Me.DATE, SHRH, 0, 5, -1, Me.USER_NAME)
    End If
  If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
    Me.N_S = max_ns
  End If

  If Not IsNull(Me.N_S) Then
    DoCmd.RunSQL ("DELETE FROM dbo.DEED_DTL WHERE (N_S = " & Me.N_S & ")")
  End If
  Err.Clear
    If MID(Forms![BASEKNOW]![OPTIONSS], 55, 1) = "5" Then
       DoCmd.RunSQL "INSERT INTO dbo.DEED_DTL (HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, SHARH, BED, N_SERI, BANK, N_S, HES,ARZD,MHAZ_NO) SELECT     THES_K, THES_M, THES_T, THES_T2, THES_T3, THES_T4, LEFT(SHARH + ' - ' + '" & GETDEPART(Me.DEPATMAN) & "',100), MABL, N_SERI, BANK," & Me.N_S & " AS Expr1, THES,ARZD ,MHAZ_NO FROM dbo.PGET_LST WHERE  (ID = " & Me.ID & ")"
       DoCmd.RunSQL "INSERT INTO dbo.DEED_DTL (HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, SHARH, BES, N_SERI, BANK, N_S, HES,ARZD,MHAZ_NO) SELECT     FHES_K, FHES_M, FHES_T, FHES_T2, FHES_T3, FHES_T4, LEFT(SHARH + ' - ' + '" & GETDEPART(Me.DEPATMAN) & "',100), MABL, N_SERI, BANK," & Me.N_S & " AS Expr1, FHES,ARZD,MHAZ_NO FROM dbo.PGET_LST WHERE  (ID = " & Me.ID & ")"
    Else
      DoCmd.RunSQL "INSERT INTO dbo.DEED_DTL (HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, SHARH, BED, N_SERI, BANK, N_S, HES,ARZD,MHAZ_NO) SELECT     THES_K, THES_M, THES_T, THES_T2, THES_T3, THES_T4, SHARH, MABL, N_SERI, BANK," & Me.N_S & " AS Expr1, THES,ARZD ,MHAZ_NO FROM dbo.PGET_LST WHERE  (ID = " & Me.ID & ")"
      DoCmd.RunSQL "INSERT INTO dbo.DEED_DTL (HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, SHARH, BES, N_SERI, BANK, N_S, HES,ARZD,MHAZ_NO) SELECT     FHES_K, FHES_M, FHES_T, FHES_T2, FHES_T3, FHES_T4, SHARH, MABL, N_SERI, BANK," & Me.N_S & " AS Expr1, FHES,ARZD,MHAZ_NO FROM dbo.PGET_LST WHERE  (ID = " & Me.ID & ")"
    End If
  If Err <> 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "خطا در صدور سند رخ داده است بعضي لز حساب ها وجود ندارد "
  End If
  SHRST.Close
   DoCmd.Close acForm, "bun"

  DoCmd.RunCommand acCmdSave
  Me.PGET_LST_SUB.Requery
End Sub


Private Sub Command23_Click()
                   On Error Resume Next
  DoCmd.OpenReport "SANAD_DARYAFT", acPreview, "", "ID =" & Me.ID
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
     If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.PGET_LST_SUB.Enabled = False
       '  Me.ESLAH.Enabled = True
    End If
End Sub

Private Sub Command24_Click()
                   On Error Resume Next
  DoCmd.OpenReport "SANAD_PARDAKHT", acPreview, "", "ID =" & Me.ID
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
     If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.PGET_LST_SUB.Enabled = False
       '  Me.ESLAH.Enabled = True
    End If
End Sub

Private Sub DATE_AfterUpdate()
                     On Error Resume Next
    DTDT = Me.DATE.Text
  If IsNull(Me.DATE.Text) Then
    Me.PGET_LST_SUB.Enabled = False
  Else
    Me.PGET_LST_SUB.Enabled = True
    DTCHK = True
  End If
End Sub

Private Sub DATE_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(Me.DATE, Forms![BASEKNOW]![CTL_DT])
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟ در صورتيكه مايليد تغييرات ثبت نشود بعداز بستن اين پنجره كليد  اسكيپ را فشار دهيد."
        If Not Forms![BASEKNOW]![Text44] Then
            CANCEL = True
        End If
    End If

End Sub

Private Sub ESLAH_Click()
Dim dt As Date
                On Error Resume Next
  If Not IsNull(Me.ID) Then
  dt = Now()
  '  If Forms![baseknow]![TRANSF] Then
    Call TR("PGET_HED", "(ID = " & Me.ID & " )", dt, 1)
    Call TR("PGET_LST", "(ID = " & Me.ID & " )", dt, 2)
       ' DoCmd.RunSQL ("INSERT INTO dbo.TR_PGET_HED   (ID, DATE,MOLAH, N_S, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, KIND, IDK, OKF, UP_TIME, UP_DATE,UP_USER_NAME,PC_NAME,IPADD) SELECT  ID, DATE, MOLAH, N_S, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, KIND, IDK, OKF," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'   FROM dbo.PGET_HED WHERE (ID = " & Me.ID & " ) ")
  '  DoCmd.RunSQL ("INSERT INTO dbo.TR_PGET_LST   (ID, DATE, RADIF, NO_AM, NAHVA, FHES_K, FHES_M, FHES_T, THES_K, THES_M, THES_T, SHARH, MABL, N_SERI, BANK, FHES, THES, UP_TIME, UP_DATE) SELECT     ID, DATE, RADIF, NO_AM, NAHVA, FHES_K, FHES_M, FHES_T, THES_K, THES_M, THES_T, SHARH, MABL, N_SERI, BANK, FHES, THES," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2  FROM dbo.PGET_LST WHERE     (ID = " & Me.ID & ")")
 '   End If
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.PGET_LST_SUB.Enabled = True
        Me.PGET_LST_SUB.Form.AllowAdditions = True
        Me.AllowDeletions = True
        If Me.SGN1 Or SGN2 Or SGN3 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
            Me.PGET_LST_SUB.Locked = True
            Me.DATE.Locked = True
            Me.MOLAH.Locked = True
            Me.AllowEdits = True
        Else
            Me.PGET_LST_SUB.Locked = False
            Me.MOLAH.Locked = False
            Me.DATE.Locked = False
            Me.AllowEdits = True
            Me.PGET_LST_SUB.Form.AllowEdits = True
       End If
    '    If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Me.USER_NAME Then
    '        Me.USER_NAME = UCurrentUser()
    '        DoCmd.RunCommand acCmdSaveRecord
    '    End If
   End If
End Sub

Private Sub Form_AfterUpdate()
                       On Error Resume Next
    Call SANAD
    LETSANAD = False
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
    If IsNull(ID) Or (ID = 0) Then
      Dim rst As New ADODB.Recordset
 '***************************************************************
     '   CurrentProject.Connection.Execute ("set transaction isolation level SERIALIZABLE"): DoCmd.OpenForm "BUN"
     '   CurrentProject.Connection.Execute ("begin transaction")
        Set rst = CurrentProject.Connection.Execute("SELECT Max(pget_hed.id) AS MaxOfid FROM pget_hed ")
 '***************************************************************
      If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
          Me.ID = 1
      Else
          Me.ID = rst.Fields(0) + 1
      End If
      rst.Close
        If Me.ID > 0 Then
             Call LetSigneTick(Me, 34, Forms![BASEKNOW]![USERCOD])
        Else
             Me.SGN1.Locked = True
             Me.SGN2.Locked = True
             Me.SGN3.Locked = True
        End If
    End If
    If IsNull(IDK) Or (IDK = 0) Then
      Set rst = New ADODB.Recordset
      rst.Open "SELECT     MAX(IDK) AS MaxOfidK FROM dbo.PGET_HED WHERE     (KIND = " & Me.KIND & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
          Me.IDK = 1
      Else
          Me.IDK = rst.Fields(0) + 1
      End If
      rst.Close
    End If
 '***************************************************************
     '    CurrentProject.Connection.Execute ("COMMIT TRANSACTION"): DoCmd.Close acForm, "bun"
 '***************************************************************
End Sub

Private Sub Form_Close()
    Me.ServerFilter = ""
End Sub

Private Sub Form_Current()
  Dim rst As New ADODB.Recordset
                   On Error Resume Next
  If IsNull(Me.DATE) Then
    Me.PGET_LST_SUB.Enabled = False
  Else
    Me.PGET_LST_SUB.Enabled = True
  End If
  Me.MABNA = Null
  If IsNull(Me.DATE) Then
    Me.PGET_LST_SUB.Enabled = False
  Else
    Me.PGET_LST_SUB.Enabled = True
    If IsNull(Me.N_S) Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "خزانه داري جاري سند نخورده است بنابر اين براي آن سند جديد صادر ميگردد"
        Call SANAD
    Else
        rst.Open "select * from dbo.DEED_HED where N_S = " & Me.N_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.MABNA = rst.Fields("BASE")
            If rst.Fields("GHATEI") Then
                LETSANAD = False
                Me.AllowDeletions = False
                Me.AllowEdits = False
                Me.PGET_LST_SUB.Locked = True
                Me.PGET_LST_SUB.Form.AllowAdditions = False
                Me.PGET_LST_SUB.Form.AllowDeletions = False
                Me![PGET_LST_SUB].Form.Refresh
            Else
                LETSANAD = True
                Me.AllowDeletions = True
                Me.AllowEdits = True
                Me.PGET_LST_SUB.Locked = False
                Me.PGET_LST_SUB.Form.AllowEdits = True
                Me.PGET_LST_SUB.Form.AllowDeletions = True
                Me![PGET_LST_SUB].Form.Refresh
            End If
         End If
         rst.Close
     End If

 End If
 If Me.OKF Then
     Me.AllowDeletions = False
     Me.AllowEdits = False
     Me.PGET_LST_SUB.Enabled = False
     Me.ESLAH.Enabled = True
End If
If Not Me.NewRecord And UCurrentUser() <> Me.USER_NAME Then
    If Not LETSGO("DPSEE") Then
       Me.ESLAH.Enabled = False
    Else
       Me.ESLAH.Enabled = True
    End If
Else
       Me.ESLAH.Enabled = True
End If
    If Forms![BASEKNOW]![SIGN] Then
        If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
            Me.Command12.Enabled = True
            Me.Command23.Enabled = True
            Me.Command24.Enabled = True
        Else
            Me.Command12.Enabled = False
            Me.Command23.Enabled = False
            Me.Command24.Enabled = False
        End If
    End If
   Me.PERSONEL.Visible = True
   If Me.ID > 0 Then
        Call LetSigneTick(Me, 34, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                      On Error Resume Next
    Dim frm As Form
    Call SANAD
    Set frm = Me![PGET_LST_SUB].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين خزانه داري داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس خزانه داري را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فرم كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If


End Sub

Private Sub Form_Load()
                     On Error Resume Next
 ' Me.lock2.Password = "neginsedighahmad"
'  Me.lock2.Connected = True
 ' If Me.lock2.ErrorCode <> 0 And Me.lock2.ErrorCode <> 7 Then
 '    DoCmd.OpenForm "lockok"
 ' End If
   If MID(Forms![BASEKNOW]![OPTIONSS], 42, 1) = "5" Then
    Me.MANDB.Visible = True
    Me.MANDS.Visible = True
  End If
 LETSANAD = True
  Call SETSECURITY(Me.NAME, "PGETD", 3)
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Command12_Click()
                     On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    Dim stDocName As String
    stDocName = "R_DP_PRINT"
    DoCmd.OpenReport stDocName, acPreview, , "ID  = " & Me.ID
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
     If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.PGET_LST_SUB.Enabled = False
     '    Me.ESLAH.Enabled = True
    End If
    Err.Clear
End Sub

Private Sub Form_Open(CANCEL As Integer)
Dim rst As New ADODB.Recordset, i As Integer
    Dim VS As String
    On Error Resume Next
    DoCmd.GoToRecord acDataForm, Me.NAME, acLast
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
    If Not LETSGO("ESLAHK") Then
       Me.ESLAH.Visible = False
    Else
       Me.ESLAH.Visible = True
    End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
       Me.OKF.DefaultValue = True
    Else
       Me.OKF.DefaultValue = False
    End If
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
           Me.ServerFilter = "(DEPATMAN = " & Forms![Default]![TFSAZMAN] & ")"
        Else
           Me.ServerFilter = Me.ServerFilter & " and  (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ")"
        End If
    End If
Else
    Me.ServerFilter = Me.OpenArgs
End If
  '  If Not LETSGO("DPDEED") Then
  '     Me.RecordSource = "SELECT     PGET_HED.* FROM PGET_HED WHERE     (USER_NAME = N'" & UCurrentUser() & "') ORDER BY DATE, ID"
  '  End If
    If Forms![BASEKNOW]![SIGN] Then
       Me.SGN1.Visible = True
       Me.SGN2.Visible = True
       Me.SGN3.Visible = True
       Me.SGN1usid.Visible = True
       Me.sgn2usid.Visible = True
       Me.sgn3usid.Visible = True
        If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
            Me.Command12.Enabled = True
            Me.Command23.Enabled = True
            Me.Command24.Enabled = True
        Else
            Me.Command12.Enabled = False
            Me.Command23.Enabled = False
            Me.Command24.Enabled = False
        End If
    End If
    Me.PERSONEL.RowSource = GetUserList
    Me.SGN1usid.RowSource = GetUserList
    Me.sgn2usid.RowSource = GetUserList
    Me.sgn3usid.RowSource = GetUserList
    Me.PGET_LST_SUB.Form.SHARH.ColumnWidth = 7410
    If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Dim shw
           shw = Me.PGET_LST_SUB.Width
           Me.Width = Me.WindowWidth - Me.WindowWidth * 0.02
           Me.Detail.Height = Me.WindowHeight * 0.99
           Me.PGET_LST_SUB.Height = Me.Detail.Height * 0.7
           Me.PGET_LST_SUB.Left = 50
           Me.PGET_LST_SUB.Width = Me.Width - 20
           Me.Text10.Top = Me.PGET_LST_SUB.Height + Me.PGET_LST_SUB.Top + 50
           Me.Label11.Top = Me.Text10.Top: Me.MANDS.Top = Me.Text10.Top: Me.Label19.Top = Me.Text10.Top: Me.mabl.Top = Me.Text10.Top
           Me.Label9.Top = Me.Text10.Top + Me.Text10.Height + 50
           Me.MANDB.Top = Me.Label9.Top: Me.Text8.Top = Me.Label9.Top
           Me.LSGN1.Top = Me.Label9.Top + Me.Label9.Height + 100
           Me.LSGN2.Top = Me.LSGN1.Top: Me.LSGN3.Top = Me.LSGN1.Top: Me.SGN1.Top = Me.LSGN1.Top: Me.SGN2.Top = Me.LSGN1.Top: Me.SGN3.Top = Me.LSGN1.Top: Me.SGN1usid.Top = Me.LSGN1.Top: Me.sgn2usid.Top = Me.LSGN1.Top: Me.sgn3usid.Top = Me.LSGN1.Top: Me.LPERSONEL.Top = Me.LSGN1.Top: Me.PERSONEL.Top = Me.LSGN1.Top
           Me.PGET_LST_SUB.Form.SHARH.ColumnWidth = 7410 + Me.PGET_LST_SUB.Width - shw
           Me.Repaint
           DoCmd.Minimize
           DoCmd.Maximize
           DoCmd.Maximize
        End If
    End If
            Me.Refresh

End Sub

Private Sub id_DblClick(CANCEL As Integer)
    If Not IsNull(Me.ID) <> 0 Then
      Call SANAD
    End If

End Sub

Private Sub IDK_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "FERSAL_SANAD", , , , , acDialog, 11
End Sub

Private Sub KIND_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  If Me.KIND <> Me.KIND.OldValue Then
    If IsNull(IDK) Or (IDK = 0) Then
      Set rst = New ADODB.Recordset
      rst.Open "SELECT     MAX(IDK) AS MaxOfidK FROM dbo.PGET_HED WHERE     (KIND = " & Me.KIND & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
          Me.IDK = 1
      Else
          Me.IDK = rst.Fields(0) + 1
      End If
      rst.Close
    Else
      Forms![BASEKNOW]![Text44] = False
      DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "با تغيير نوع برگه شماره جديد به آن اختصاص مي يابد آيا اين عمل را تائيد مي نمائيد؟"
      If Forms![BASEKNOW]![Text44] Then
          Set rst = New ADODB.Recordset
          rst.Open "SELECT     MAX(IDK) AS MaxOfidK FROM dbo.PGET_HED WHERE     (KIND = " & Me.KIND & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
              Me.IDK = 1
          Else
              Me.IDK = rst.Fields(0) + 1
          End If
          rst.Close
       Else
          CANCEL = True
       End If
    End If
  End If
End Sub

Private Sub MABNA_DblClick(CANCEL As Integer)
    If IsLoaded("TARAZ_4_TAFZIL") Then
      DoCmd.Close acForm, "TARAZ_4_TAFZIL"
    End If
    DoCmd.OpenForm "FMENU_TARAZ_4", , , , , acHidden
    Forms![FMENU_TARAZ_4]![DT1] = 1
    Forms![FMENU_TARAZ_4]![DT2] = 99999999
    Forms![FMENU_TARAZ_4]![SNDNUM1] = Me.N_S
    Forms![FMENU_TARAZ_4]![SNDNUM2] = Me.N_S
    DoCmd.OpenForm "TARAZ_4_TAFZ_DP", acNormal

End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                       On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , , , , "N_S = " & Me.N_S
End Sub
Private Sub PGET_LST_SUB_Exit(CANCEL As Integer)
                   On Error Resume Next
  If LETSANAD Then
    Call SANAD
    LETSANAD = False
  End If
End Sub

Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(34, Me.ID, Me.PERSONEL, "'خزانه داري   شماره: " & Me.ID & " مورخ " & Format(Me.DATE, "####/##/##") & "  به نام: " & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & "','" & GETUSERHES(Forms![BASEKNOW]![USERCOD]) & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.ID, 34)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",34," & Me.ID & ",34 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'خزانه داري   شماره: " & Me.ID & " مورخ " & Format(Me.DATE, "####/##/##") & "  به نام: " & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & "','" & GETUSERHES(Forms![BASEKNOW]![USERCOD]) & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",34," & Me.ID & ",34," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.ID, 34)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",34," & Me.ID & ",34 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command12.Enabled = True
        Me.Command23.Enabled = True
        Me.Command24.Enabled = True
    Else
        Me.Command12.Enabled = False
        Me.Command23.Enabled = False
        Me.Command24.Enabled = False
    End If
End Sub

Private Sub SGN2_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.ID, 34)

    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",34," & Me.ID & ",34 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'خزانه داري   شماره: " & Me.ID & " مورخ " & Format(Me.DATE, "####/##/##") & "  به نام: " & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & "','" & GETUSERHES(Forms![BASEKNOW]![USERCOD]) & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",34," & Me.ID & ",34," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.ID, 34)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",34," & Me.ID & ",34 )")
   End If

   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command12.Enabled = True
        Me.Command23.Enabled = True
        Me.Command24.Enabled = True
    Else
        Me.Command12.Enabled = False
        Me.Command23.Enabled = False
        Me.Command24.Enabled = False
    End If
End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.ID, 34)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",34," & Me.ID & ",34 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'خزانه داري   شماره: " & Me.ID & " مورخ " & Format(Me.DATE, "####/##/##") & "  به نام: " & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & "','" & GETUSERHES(Forms![BASEKNOW]![USERCOD]) & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",34," & Me.ID & ",34," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.ID, 34)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",34," & Me.ID & ",34 )")
   End If
   Me.PERSONEL.Visible = True
   Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command12.Enabled = True
        Me.Command23.Enabled = True
        Me.Command24.Enabled = True
    Else
        Me.Command12.Enabled = False
        Me.Command23.Enabled = False
        Me.Command24.Enabled = False
    End If

End Sub
