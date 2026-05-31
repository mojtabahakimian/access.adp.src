Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim LETSANAD As Boolean
Dim Meidnum As Double
Private Sub Command100_Click()
                     On Error Resume Next
    Dim frm As Form
'  Call CheckCustomPage("INVOICE_KHAREED", frm.RecordsetClone.RecordCount * 0.9 + 20)
    DoCmd.OpenReport "INVOICE_KHAREED", acPreview, "", "HEAD_LST.NUMBER =" & Me.NUMBER & " AND HEAD_LST.TAG =" & Me.HTAG
End Sub

Private Sub Command106_Click()
    Dim rptname As String
                     On Error Resume Next
    If Forms![BASEKNOW]![LOCKFAP] Then
       Me.OKF = True
    End If
    rptname = "DARKHST_KHARID"
    DoCmd.RunCommand acCmdSaveRecord
    If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
         rptname = "DARKHST_KHARID_" & MID(Forms![BASEKNOW]![OPTIONSS], 11, 2)
    End If
    Err.Clear
    DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER
    If Err <> 0 Then
        DoCmd.OpenReport "DARKHST_KHARID", acPreview, "", "NUMBER =" & Me.NUMBER
    End If
    If Me.OKF Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_REQUEST_SUB.Locked = True
        Me.INVO_LST_REQUEST_SUB.Form.AllowDeletions = False
        Me.ESLAH.Enabled = True
    End If
End Sub

Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub ESLAH_Click()
       On Error Resume Next
 Dim dt As Date
 Dim rst As New ADODB.Recordset
    dt = Now()
    Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 23)", dt, 1)
    Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 23)", dt, 1)
 If Not IsNull(Me.NUMBER) Then
    rst.Open "select * from head_lst where tag = 1 and NUMBER1 =  " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me![INVO_LST_REQUEST_SUB].Form.AllowAdditions = True
        Me.INVO_LST_REQUEST_SUB.Enabled = True
        Me.INVO_LST_REQUEST_SUB.Locked = False
        Me![INVO_LST_REQUEST_SUB].Form.Refresh
        If Me.SGN1 Or SGN2 Or SGN3 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
            Me.INVO_LST_REQUEST_SUB.Locked = True
            Me.DATE_N.Locked = True
            Me.NUMBER.Locked = True
            Me.FNUMCO.Locked = True
            Me.MOLAH.Locked = True
        Else
            Me.INVO_LST_REQUEST_SUB.Locked = False
            Me.NUMBER.Locked = False
            Me.FNUMCO.Locked = False
            Me.MOLAH.Locked = False
            Me.DATE_N.Locked = False
        End If
    Else
        Me.AllowDeletions = False
        Me![INVO_LST_REQUEST_SUB].Form.AllowAdditions = False
        Me![INVO_LST_REQUEST_SUB].Form.Refresh
        DoCmd.OpenForm "mesageform", , , , , acDialog, " براي اين درخواست رسيد صادر شده است و قابل تغيير نمي باشد ....!"
    End If
 End If
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 36, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
 Call SETSECURITY(Me.NAME, "DARKHR", 3)

End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
      If Me.NUMBER = 0 Then
        Dim rst As New ADODB.Recordset
        rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
            Me.NUMBER = Forms![BASEKNOW]![STHKH]
        Else
            Me.NUMBER = rst.Fields(0) + 1
        End If
           rst.Close
      End If
       ' If Me.USER_NAME <> UCurrentUser() Then
       '     Me.USER_NAME = UCurrentUser()
       ' End If
  Me.INVO_LST_REQUEST_SUB.Locked = False
  Me.INVO_LST_REQUEST_SUB.Requery
End Sub

Private Sub Form_Current()
                  On Error Resume Next
 Dim frm As Form
  Dim ghat As Boolean
                   On Error Resume Next
  Set frm = Me![INVO_LST_REQUEST_SUB].Form
  frm.Requery
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command106.Enabled = True
  Else
    Me.Command106.Enabled = False
  End If
  If Forms![BASEKNOW]![SIGN] Then
    If Me.SGN2 Or Me.SGN3 Then
        Me.Command106.Enabled = True
    Else
        Me.Command106.Enabled = False
    End If
  End If
  If Me.NewRecord Then
        Me.INVO_LST_REQUEST_SUB.Locked = True
        Me.INVO_LST_REQUEST_SUB.Enabled = True
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me![INVO_LST_REQUEST_SUB].Form.AllowAdditions = True
        Me![INVO_LST_REQUEST_SUB].Form.Refresh
  Else
    Dim rst As New ADODB.Recordset
    rst.Open "select * from head_lst where tag = 12 and NUMBER =  " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or Left(UCurrentUser(), 10) = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me![INVO_LST_REQUEST_SUB].Form.AllowAdditions = True
        Me.INVO_LST_REQUEST_SUB.Locked = False
        Me![INVO_LST_REQUEST_SUB].Form.Refresh
    Else
        Me.AllowDeletions = False
        Me![INVO_LST_REQUEST_SUB].Form.AllowAdditions = False
        Me![INVO_LST_REQUEST_SUB].Form.Refresh
    End If
   rst.Close
 End If
  If Forms![BASEKNOW]![SIGN] And False Then
        Me.SGN1.Visible = True
        Me.SGN2.Visible = True
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT    ANB_RASID_ANB,ANB_RASID_QC FROM dbo.SIGN WHERE     USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If rst.Fields("ANB_RASID_ANB") Then
              Me.SGN1.Enabled = True
            Else
              Me.SGN1.Enabled = False
            End If
            If rst.Fields("ANB_RASID_QC") Then
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
            Me.INVO_LST_REQUEST_SUB.Locked = True
            Me.INVO_LST_REQUEST_SUB.Form.AllowDeletions = False
            Me.NUMBER.Locked = True
            Me.DATE_N.Locked = True
            Me.FNUMCO.Locked = True
            Me.MOLAH.Locked = True
            Me.TAH.Locked = True
        End If
   End If
   Err.Clear
   Call SETSECURITY(Me.NAME, "DARKHR", 3)
   If Me.OKF Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_REQUEST_SUB.Enabled = False
        Me.ESLAH.Enabled = True
   End If
  If Forms![BASEKNOW]![SIGN] Then
      If Me.SGN2 Or Me.SGN3 Then
         Me.Command106.Enabled = True
     Else
         Me.Command106.Enabled = False
    End If
  End If
  If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 36, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                     On Error Resume Next
    Dim frm As Form
    Set frm = Me![INVO_LST_REQUEST_SUB].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين برگه داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس برگه را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub

Private Sub Form_Load()
                   On Error Resume Next
'  Me.lock2.Password = "neginsedighahmad"
 ' Me.lock2.Connected = True
 ' If Me.lock2.ErrorCode <> 0 And Me.lock2.ErrorCode <> 7 Then
 '    DoCmd.OpenForm "lockok"
'  End If
  If Forms![BASEKNOW]![UPDDATE] Then
     Me.DATE_N.Locked = False
  Else
     Me.DATE_N.Locked = True
  End If
  If Not IsNull(Me.OpenArgs) Then
      Me.RecordSource = "SELECT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 23) AND " & Me.OpenArgs
  Else
    DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
  End If
  Me.NUMBER.SetFocus
  Call SETSECURITY(Me.NAME, "DARKHR", 3)
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
    If Forms![BASEKNOW]![SIGN] Then
       Me.SGN1.Visible = True
       Me.SGN2.Visible = True
       Me.SGN3.Visible = True
       Me.LSGN1.Visible = True
       Me.LSGN2.Visible = True
       Me.LSGN3.Visible = True
       Me.SGN1usid.Visible = True
       Me.sgn2usid.Visible = True
       Me.sgn3usid.Visible = True
       Me.PERSONEL.Visible = True
End If
    Me.PERSONEL.RowSource = GetUserList
    Me.SGN1usid.RowSource = GetUserList
    Me.sgn2usid.RowSource = GetUserList
    Me.sgn3usid.RowSource = GetUserList
    Me.INVO_LST_REQUEST_SUB.Form.CODE.ColumnWidth = 5610
    Me.INVO_LST_REQUEST_SUB.Width = 14160
    If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Me.Width = Me.WindowWidth - Me.WindowWidth * 0.02
           Dim shw
           shw = 14160
           Me.INVO_LST_REQUEST_SUB.Left = 50
           Me.INVO_LST_REQUEST_SUB.Width = Me.Width - 20
           Me.INVO_LST_REQUEST_SUB.Form.CODE.ColumnWidth = 5610 + (Me.INVO_LST_REQUEST_SUB.Width - shw) / 2
      '     Me.HAVALA_EXIT_SUB.Form.N_RASID.ColumnWidth = 4245 + (Me.HAVALA_EXIT_SUB.Width - shw) / 2
           Me.Repaint
        End If
    End If
End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "FERSAL_SANAD", , , , , acDialog, 2
End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 36)
    SHARH = "'درخواست خريد  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & Me.TAH & "','" & GETUSERHES(Forms![BASEKNOW]![USERCOD]) & "'"
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",36," & Me.NUMBER & ",36 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",36," & Me.NUMBER & ",36," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 36)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",36," & Me.NUMBER & ",36 )")
   End If
   Call PERSONELUpdate(36, Me.NUMBER, Me.PERSONEL, SHARH)
   Me.PERSONEL.Visible = True
   Meidnum = MID: If Not Me.OKF Then Me.OKF = True
   Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
   If Me.SGN2 Or Me.SGN3 Then
       Me.Command106.Enabled = True
   Else
       Me.Command106.Enabled = False
   End If
End Sub

Private Sub SGN2_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 36)

        SHARH = "'درخواست خريد  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & Me.TAH & "','" & GETUSERHES(Forms![BASEKNOW]![USERCOD]) & "'"
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",36," & Me.NUMBER & ",36 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",36," & Me.NUMBER & ",36," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 36)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",36," & Me.NUMBER & ",36 )")
   End If
   Call PERSONELUpdate(36, Me.NUMBER, Me.PERSONEL, SHARH)
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN2 Or Me.SGN3 Then
        Me.Command106.Enabled = True
    Else
        Me.Command106.Enabled = False
    End If
End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 36)
    SHARH = "'درخواست خريد  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & Me.TAH & "','" & GETUSERHES(Forms![BASEKNOW]![USERCOD]) & "'"
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",36," & Me.NUMBER & ",36 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",36," & Me.NUMBER & ",36," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 36)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",36," & Me.NUMBER & ",36 )")
   End If
   Call PERSONELUpdate(36, Me.NUMBER, Me.PERSONEL, SHARH)
   Me.PERSONEL.Visible = True
   Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN2 Or Me.SGN3 Then
        Me.Command106.Enabled = True
    Else
        Me.Command106.Enabled = False
    End If
End Sub
Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(36, Me.NUMBER, Me.PERSONEL, "'درخواست خريد  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & Me.TAH & "','" & GETUSERHES(Forms![BASEKNOW]![USERCOD]) & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub
