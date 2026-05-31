Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
 Dim SDRST As New ADODB.Recordset, SHRST As New ADODB.Recordset, firstcreate As Boolean
Dim Meidnum As Double
Private Sub SANAD()
    If Me.USER_NAME <> UCurrentUser() Then
        Me.USER_NAME = UCurrentUser()
    End If
                     On Error Resume Next
 If Forms![BASEKNOW]![SANAD] <> 3 Then
  DoCmd.RunCommand acCmdSaveRecord
  Dim rst As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, JSTT As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH As Double, SHART As String, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4
  rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields(0)) Then
    max_ns = 1
  Else
    max_ns = rst.Fields(0) + 1
  End If
  Err.Clear
  SHART = "NO_S = 12 AND N_S = " & Me.N_S
  SHRST.Filter = SHART
  Set SHRST = SHRST
  If SHRST.RecordCount = 0 Or Err <> 0 Then
    SHRST.AddNew
    SHRST.Fields("N_S") = max_ns
    SHRST.Fields("DATE_S") = Me.DATE_N
    SHRST.Fields("SHARH_S") = Left(" حواله خروج ساير مواد از انبار شماره " & Me.NUMBER & "-" & Me.FNUMCO & "مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 12
    SHRST.Fields("USER_NAME") = UCurrentUser()
    SHRST.Fields("CRT") = Now()
    SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
    SHRST.update
  Else
        max_ns = SHRST.Fields("N_S")
        SHRST.Fields("DATE_S") = Me.DATE_N
        SHRST.Fields("SHARH_S") = Left(" حواله خروج ساير مواد از انبار شماره " & Me.NUMBER & "-" & Me.FNUMCO & "مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
        SHRST.Fields("GHATEI") = 0
        SHRST.Fields("NO_S") = 12
        SHRST.Fields("USER_NAME") = UCurrentUser()
        SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
        SHRST.update
  End If
  If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
     Me.N_S = max_ns
  End If
  DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
  JST.Open "SELECT     SANAD_NO, N_RASID, MABL_K, MEGHk, CODE, ANBAR FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 11)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not JST.EOF()
     If IsNumeric(JST.Fields("N_RASID")) Then
            JAMCH = 0
            JSTT.Open "SELECT     FNUMB, NUMBER, TNUMBER, N_KOL, NAMES FROM dbo.HEAD_MANF WHERE     (FNUMB = " & JST.Fields("N_RASID") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If JST.Fields("MABL_K") <> 0 Then
                   SDRST.AddNew 'كنترل كالاي در جريان سا
                   SDRST.Fields("N_S") = max_ns
                   SDRST.Fields("HES_K") = JSTT.Fields("N_KOL")
                   SDRST.Fields("HES_M") = JSTT.Fields("NUMBER")
                   SDRST.Fields("HES_T") = JSTT.Fields("TNUMBER")
                   SDRST.Fields("hes") = JSTT.Fields("N_KOL") & "-" & JSTT.Fields("NUMBER") & "-" & JSTT.Fields("TNUMBER")
                   SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
                   SDRST.Fields("SHARH") = Left("حواله خروج ساير شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(JSTT.Fields("NAMES")), 255)
                   SDRST.Fields("NUMBER") = Me.NUMBER
                   SDRST.Fields("TAG") = 11
                   SDRST.update
                   JAMCH = JST.Fields("MABL_K")
                   SDRST.AddNew 'انبار
                   SDRST.Fields("N_S") = max_ns
                   SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                   SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                   SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                   SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                   SDRST.Fields("SHARH") = Left("حواله خروج  ساير  مواد شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
                   SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                   SDRST.Fields("NUMBER") = Me.NUMBER
                   SDRST.Fields("TAG") = 11
                   SDRST.Fields("MHAZ_NO") = JST.Fields("SANAD_NO")
                   SDRST.update
             End If
             Set JSTT = New ADODB.Recordset
      Else
            Call GETTAF3(JST.Fields("N_RASID"), CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
            If JST.Fields("MABL_K") <> 0 Then
                SDRST.AddNew 'كنترل كالاي در جريان سا
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = CKOL
                SDRST.Fields("HES_M") = CMOIN
                SDRST.Fields("HES_T") = CTAF
                SDRST.Fields("HES_T2") = CTAF2
                SDRST.Fields("HES_T3") = CTAF3
                SDRST.Fields("HES_T4") = CTAF4
                SDRST.Fields("hes") = JST.Fields("N_RASID")
                SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
                SDRST.Fields("SHARH") = Left("حواله خروج ساير شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(GETTAFNAME(JST.Fields("N_RASID"))), 255)
                SDRST.Fields("NUMBER") = Me.NUMBER
                SDRST.Fields("TAG") = Me.HTAG
                SDRST.update
                JAMCH = JST.Fields("MABL_K")
                SDRST.AddNew 'انبار
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                SDRST.Fields("SHARH") = Left("حواله خروج  ساير  مواد شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
                SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                SDRST.Fields("NUMBER") = Me.NUMBER
                SDRST.Fields("TAG") = Me.HTAG
                SDRST.Fields("MHAZ_NO") = JST.Fields("SANAD_NO")
                SDRST.update
            End If
      End If
      JST.MoveNext
  Wend
  End If
  LETSANAD = False
End Sub

Private Sub Command103_Click()
      On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    If MOGUDI(Me.NUMBER, 11) Then
        DoCmd.OpenReport "HAVALAH_EXIT", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
    End If
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.HAVALAH_EXIT_SAYER_SUB.Enabled = False
         Me.ESLAH.Enabled = True
    End If
End Sub
Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "15"
  End If
  Response = acDataErrContinue
  Err.Clear


End Sub


Private Sub CUST_NO2_Exit(CANCEL As Integer)
                     On Error Resume Next
  If IsNull(Me.CUST_NO2) Or Me.CUST_NO2 = "+" Or Me.CUST_NO2 = "-" Then
     DoCmd.OpenForm "SERCH"
  End If
End Sub

Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If Me.CUST_NO2.Text = "-" Or Me.CUST_NO2.Text = "+" Then
     Me.CUST_NO2.Text = ""
     DoCmd.OpenForm "SERCH"
  End If
Response = acDataErrContinue
End Sub

Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub ESLAH_Click()
Dim dt As Date
                On Error Resume Next
  If Not IsNull(Me.NUMBER) Then
     dt = Now()
      Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & " ) AND (TAG = 11)", dt, 1)
      Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & " ) AND (TAG = 11)", dt, 2)
     Me.AllowDeletions = True
     Me.AllowEdits = True
     Me.HAVALAH_EXIT_SAYER_SUB.Enabled = True
     Me.HAVALAH_EXIT_SAYER_SUB.Form.AllowAdditions = True
     Me.AllowDeletions = True
    If Me.NUMBER > 0 Then
         Call LetSigneTick(Me, 38, Forms![BASEKNOW]![USERCOD])
    Else
         Me.SGN1.Locked = True
         Me.SGN2.Locked = True
         Me.SGN3.Locked = True
    End If
     If Me.SGN1 Or SGN2 Or SGN3 Then
         DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
         Me.HAVALAH_EXIT_SAYER_SUB.Locked = True
         Me.DATE_N.Locked = True
         Me.CUST_NO.Locked = True

         Me.MOLAH.Locked = True
         Me.AllowEdits = True
     Else
         Me.HAVALAH_EXIT_SAYER_SUB.Locked = False
         Me.MOLAH.Locked = False
         Me.DATE_N.Locked = False
         Me.CUST_NO.Locked = False
         Me.AllowEdits = True
         Me.HAVALAH_EXIT_SAYER_SUB.Form.AllowEdits = True
    End If
   End If
  Call SETSECURITY(Me.NAME, "HSAYER", 3)

End Sub

Private Sub FNUMCO_DblClick(CANCEL As Integer)
    If Me.NUMBER <> 0 Then
      Call SANAD
    End If

End Sub

Private Sub Form_AfterUpdate()
                       On Error Resume Next
    Call SANAD
    LETSANAD = False
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next
  If IsNull(Me.CUST_NO) Or Me.CUST_NO = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, "مسئول شيفت مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If Me.NUMBER = 0 Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER = Forms![BASEKNOW]![STKHS]
    Else
        Me.NUMBER = rst.Fields(0) + 1
    End If
    rst.Close
  End If
  Me.HAVALAH_EXIT_SAYER_SUB.Locked = False
  Me![HAVALAH_EXIT_SAYER_SUB].Form.AllowDeletions = True
  Me.HAVALAH_EXIT_SAYER_SUB.Requery
End Sub


Private Sub Form_Close()
  Me.ServerFilter = ""
End Sub

Private Sub Form_Current()
          On Error Resume Next
  Dim frm As Form
  Dim ghat As Boolean
  Set frm = Me![HAVALAH_EXIT_SAYER_SUB].Form
  frm.Requery
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command103.Enabled = True
  Else
    Me.Command103.Enabled = False
  End If
  If Forms![BASEKNOW]![SIGN] And False Then
    If Me.SGN2 Then
      Me.Command103.Enabled = True
    Else
      Me.Command103.Enabled = False
    End If
  End If
  Me.lsanad.ForeColor = 65535
  If Me.NUMBER = 0 Then
    Me.HAVALAH_EXIT_SAYER_SUB.Locked = True
    Me.HAVALAH_EXIT_SAYER_SUB.Enabled = True
  Else
    Me.HAVALAH_EXIT_SAYER_SUB.Locked = False
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
            Me.HAVALAH_EXIT_SAYER_SUB.Locked = True
            Me![HAVALAH_EXIT_SAYER_SUB].Form.AllowDeletions = False
            Me.lsanad.ForeColor = 125
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.HAVALAH_EXIT_SAYER_SUB.Locked = False
            Me![HAVALAH_EXIT_SAYER_SUB].Form.AllowDeletions = True
            Me.lsanad.ForeColor = 65535
        End If
     End If
        rst.Close
  End If
If Me.SGN2 Or Me.SGN3 Then
    Me.HAVALAH_EXIT_SAYER_SUB.Locked = True
End If
   If Me.NUMBER <> 0 Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.HAVALAH_EXIT_SAYER_SUB.Enabled = False
        Me.ESLAH.Enabled = True
   End If
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 38, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
  Err.Clear
  Call SETSECURITY(Me.NAME, "HSAYER", 3)
   If Me.OKF And Not Me.NewRecord Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.HAVALAH_EXIT_SAYER_SUB.Enabled = False
        Me.ESLAH.Enabled = True
   Else
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.HAVALAH_EXIT_SAYER_SUB.Enabled = True
   '     Me.ESLAH.Enabled = False
 End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                       On Error Resume Next
    Dim frm As Form
    Set frm = Me![HAVALAH_EXIT_SAYER_SUB].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين برگه  داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس برگه را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي برگه  كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub


Private Sub Form_Load()
                      On Error Resume Next
  Dim rst As New ADODB.Recordset, VS As String
    If Forms![BASEKNOW]![SANAD] <> 3 Then
      Set SDRST = Forms![recording].Form.RecordsetClone
      SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
             Me.RecordSource = "SELECT TOP (100) PERCENT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 11) and (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") ORDER BY NUMBER "
          Else
             Me.RecordSource = "SELECT TOP (100) PERCENT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 11) and (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") and " & Me.ServerFilter & "ORDER BY NUMBER "
          End If
      End If
    Else
          Me.RecordSource = "SELECT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 11) AND number = " & Me.OpenArgs
    End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
       Me.OKF.DefaultValue = True
    Else
       Me.OKF.DefaultValue = False
    End If
  Call SETSECURITY(Me.NAME, "HSAYER", 3)
  Me.CUST_NO.SetFocus
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                      On Error Resume Next
   Response = acDataErrContinue
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
    DoCmd.Maximize
    Forms![BASEKNOW]![hhwin] = Me.hWnd
    If Not LETSGO("ESLAHS") Then
       Me.ESLAH.Visible = False
    Else
       Me.ESLAH.Visible = True
    End If
    If Forms![BASEKNOW]![SIGN] Then
        Me.SGN1.Visible = True
        Me.SGN2.Visible = True
        Me.SGN3.Visible = True
        Me.SGN1usid.Visible = True
        Me.sgn2usid.Visible = True
        Me.sgn3usid.Visible = True
        Me.PERSONEL.Visible = True
        Me.PERSONEL.RowSource = GetUserList
        Me.SGN1usid.RowSource = GetUserList
        Me.sgn2usid.RowSource = GetUserList
        Me.sgn3usid.RowSource = GetUserList
    End If
    Me.HAVALAH_EXIT_SAYER_SUB.Form.CODE.ColumnWidth = 5610
    Me.HAVALAH_EXIT_SAYER_SUB.Form.N_RASID.ColumnWidth = 4245
    Me.HAVALAH_EXIT_SAYER_SUB.Height = 6825
    If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Me.Width = Me.WindowWidth - Me.WindowWidth * 0.02
           Dim shw
           shw = Me.HAVALAH_EXIT_SAYER_SUB.Width
           Me.HAVALAH_EXIT_SAYER_SUB.Left = 50
           Me.HAVALAH_EXIT_SAYER_SUB.Width = Me.Width - 20
           Me.HAVALAH_EXIT_SAYER_SUB.Form.CODE.ColumnWidth = 5610 + (Me.HAVALAH_EXIT_SAYER_SUB.Width - shw) / 2
           Me.HAVALAH_EXIT_SAYER_SUB.Height = Me.Detail.Height - 1500
           Me.Repaint
           DoCmd.Minimize
           DoCmd.Maximize
        End If
    End If
End Sub

Private Sub HAVALAH_EXIT_SAYER_SUB_Exit(CANCEL As Integer)
                   On Error Resume Next
  If LETSANAD Then
    Call SANAD
    LETSANAD = False
  End If
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                       On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub


Private Sub NUMBER_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "FERSAL_SANAD", , , , , acDialog, 5
End Sub

Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(38, Me.NUMBER, Me.PERSONEL, "'حواله خروج ساير  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    If MOGUDI(Me.NUMBER, 11) Or Not SGN1 Then
           On Error Resume Next: MID = Gettaskid(Me.NUMBER, 38)
               SHARH = "'حواله خروج ساير  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
          If MID > 0 Then
               CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",38," & Me.NUMBER & ",38 )")
               CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
               Me.PERSONEL = GETUSERTASK(MID)
          Else
               td = Now()
               CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",38," & Me.NUMBER & ",38," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
                On Error Resume Next: MID = Gettaskid(Me.NUMBER, 38)
               CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",38," & Me.NUMBER & ",38 )")
         End If
        Call PERSONELUpdate(38, Me.NUMBER, Me.PERSONEL, SHARH)
        Me.PERSONEL.Visible = True
        Meidnum = MID:
        Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
    Else
        Me.SGN1 = False
    End If
End Sub

Private Sub SGN2_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    If MOGUDI(Me.NUMBER, 11) Or Not SGN2 Then
          On Error Resume Next: MID = Gettaskid(Me.NUMBER, 38)
         SHARH = "'حواله خروج ساير  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
         If MID > 0 Then
              CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",38," & Me.NUMBER & ",38 )")
              CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
              Me.PERSONEL = GETUSERTASK(MID)
         Else
              td = Now()
              CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",38," & Me.NUMBER & ",38," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
               On Error Resume Next: MID = Gettaskid(Me.NUMBER, 38)
              CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",38," & Me.NUMBER & ",38 )")
         End If
        Call PERSONELUpdate(38, Me.NUMBER, Me.PERSONEL, SHARH)
        Me.PERSONEL.Visible = True
        Meidnum = MID:
        Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    Else
        Me.SGN2 = False
    End If
End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
       On Error Resume Next: MID = Gettaskid(Me.NUMBER, 38)
      SHARH = "'حواله خروج ساير  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
      If MID > 0 Then
          CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",38," & Me.NUMBER & ",38 )")
          CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
          Me.PERSONEL = GETUSERTASK(MID)
      Else
          td = Now()
          CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",38," & Me.NUMBER & ",38," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
           On Error Resume Next: MID = Gettaskid(Me.NUMBER, 38)
          CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",38," & Me.NUMBER & ",38 )")
     End If
   Call PERSONELUpdate(38, Me.NUMBER, Me.PERSONEL, SHARH)
   Me.PERSONEL.Visible = True
   Meidnum = MID:
   Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
End Sub
