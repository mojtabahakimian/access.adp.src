Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim PSK1, PSM1, PST1, PST2, PST3, PST4, Meidnum As Double

Private Sub CODE_AfterUpdate()
        On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
  Call MORAKHDAY_AfterUpdate
  Call LetSigneTick(Me, 37, Forms![BASEKNOW]![USERCOD])
End Sub

Private Sub CODE_NotInList(NewData As String, Response As Integer)
        On Error Resume Next
  If Not IsNumeric(Me.CODE.Text) Then
     DoCmd.OpenForm "SERCHP", acFormDS, , , , , "NAME like '%" & Me.CODE.Text & "%'"
  Else
     Me.CODE = Me.CODE.Text
  End If
  Response = acDataErrContinue
End Sub

Private Sub Command54_Click()
                     On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenReport "morakhasi", acPreview, , "IDNUM = " & Me.IDNUM
End Sub

Private Sub DA_AfterUpdate()
                 On Error Resume Next
  Me.MORAKHDAY = Me.DA * 440 + Me.HOU * 60 + Me.minu
  Call MORAKHDAY_AfterUpdate
  Call MOenDATE_AfterUpdate
End Sub


Private Sub ESLAH_Click()
  Dim dt As Date
                   On Error Resume Next
  If Not IsNull(Me.CODE) Then
        dt = Now()
        Call TR("PMORAKH", "(IDNUM = " & Me.IDNUM & " )", dt, 1)
        If Me.SGN1 Or SGN2 Or SGN3 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
            Me.CODE.Locked = True
            Me.MODATE.Locked = True
            Me.KINDM.Locked = True
            Me.MOLAH.Locked = True
            Me.MOSTDATE.Locked = True
            Me.MOENDATE.Locked = True
            Me.DA.Locked = True
            Me.HOU.Locked = True
            Me.minu.Locked = True
        Else
            Me.CODE.Locked = False
            Me.MODATE.Locked = False
            Me.KINDM.Locked = False
            Me.MOLAH.Locked = False
            Me.MOSTDATE.Locked = False
            Me.MOENDATE.Locked = False
            Me.DA.Locked = False
            Me.HOU.Locked = False
            Me.minu.Locked = False
        End If
        Me.AllowDeletions = True
        Me.AllowEdits = True
  End If
  Call SETSECURITY(Me.NAME, "MORAKH", 3)
End Sub

Private Sub Form_Close()
Me.ServerFilter = ""
End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Me.MDA = Fix(Me.MAND / 440)
  Me.MHOU = Fix(((Me.MAND / 440) - Me.MDA) * 440 / 60)
  Me.MMINU = Fix(((((Me.MAND / 440) - Me.MDA) * 440 / 60) - Me.MHOU) * 60)
  If Me.MAND = 110110110 Then
    Me.MORAKHDAY = 0
    Me.MAND = 0
  End If
  Me.DA = Fix(Me.MORAKHDAY / 440)
  Me.HOU = Fix(((Me.MORAKHDAY / 440) - Me.DA) * 440 / 60)
  Me.minu = Fix(((((Me.MORAKHDAY / 440) - Me.DA) * 440 / 60) - Me.HOU) * 60)
   If Me.OKF Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
   End If
   If Me.IDNUM > 0 Then
        Call LetSigneTick(Me, 37, Forms![BASEKNOW]![USERCOD])
   Else
            Me.CODE.Locked = False
            Me.MODATE.Locked = False
            Me.KINDM.Locked = False
            Me.MOLAH.Locked = False
            Me.MOSTDATE.Locked = False
            Me.MOENDATE.Locked = False
            Me.DA.Locked = False
            Me.HOU.Locked = False
            Me.minu.Locked = False
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
    Me.Command54.Enabled = True
End If
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "MORAKH", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
             Forms![BASEKNOW]![hhwin] = Me.hWnd
If IsNull(Me.ServerFilter) Or Me.ServerFilter = "" Then
   DoCmd.GoToRecord acActiveDataObject, , acNewRec
End If
   Err.Clear
    If Forms![BASEKNOW]![SIGN] Then
       Me.SGN1.Visible = True
       Me.SGN2.Visible = True
       Me.SGN3.Visible = True
       Me.SGN1usid.Visible = True
       Me.sgn2usid.Visible = True
       Me.sgn3usid.Visible = True
       Me.Command54.Enabled = False
    End If
    Me.PERSONEL.RowSource = GetUserList
    Me.SGN1usid.RowSource = GetUserList
    Me.sgn2usid.RowSource = GetUserList
    Me.sgn3usid.RowSource = GetUserList
    Call GETTAF3(Forms![BASEKNOW]![PERSONEL], PSK1, PSM1, PST1, PST2, PST3, PST4)
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub HOU_AfterUpdate()
                     On Error Resume Next
  Me.MORAKHDAY = Me.DA * 440 + Me.HOU * 60 + Me.minu
  Call MORAKHDAY_AfterUpdate
End Sub

Private Sub minu_AfterUpdate()
                     On Error Resume Next
  Me.MORAKHDAY = Me.DA * 440 + Me.HOU * 60 + Me.minu
  Call MORAKHDAY_AfterUpdate

End Sub

Private Sub MODATE_AfterUpdate()
                     On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
  Me.MAND = MORAKH_MAND(Me.MODATE, Me.CODE)
  Me.MDA = Fix(Me.MAND / 440)
  Me.MHOU = Fix(((Me.MAND / 440) - Me.MDA) * 440 / 60)
  Me.MMINU = Fix(((((Me.MAND / 440) - Me.MDA) * 440 / 60) - Me.MHOU) * 60)
  If Me.MAND = 110110110 Then
    Me.MORAKHDAY = 0
    Me.MAND = 0
  End If
End Sub

Private Sub MODATE_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next: CANCEL = CHEKDATEM(MODATE, False)
End Sub

Private Sub MORAKHDAY_AfterUpdate()
                   On Error Resume Next
  Dim ho As Integer, DD As Integer, mi As Integer
  DoCmd.RunCommand acCmdSaveRecord
  Me.MAND = MORAKH_MAND(Me.MODATE, Me.CODE)
  If Me.MAND = 110110110 Then
  '  Me.MORAKHDAY = 0
    Me.MAND = 0
  End If
  Me.MDA = Fix(Me.MAND / 440)
  Me.MHOU = Fix(((Me.MAND / 440) - Me.MDA) * 440 / 60)
  Me.MMINU = Fix(((((Me.MAND / 440) - Me.MDA) * 440 / 60) - Me.MHOU) * 60)
End Sub

Private Sub MOSTDATE_AfterUpdate()
Dim DD As Integer
  If MOENDATE > 0 Then
    DD = DIFF(Me.MOSTDATE, Me.MOENDATE) + 1
    If Me.DA <> DD And Me.KINDM <> 2 Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "تعداد روزهاي مرخصي با محدوده تاريخي داده شده تطبيق ندارد آن را اصلاح كردم!"
       Me.DA = DD
       Me.MORAKHDAY = Me.DA * 440 + Me.HOU * 60 + Me.minu
       Call MORAKHDAY_AfterUpdate
    End If
  End If
End Sub

Private Sub MOenDATE_AfterUpdate()
Dim DD As Integer
  If MOSTDATE > 0 Then
    DD = DIFF(Me.MOSTDATE, Me.MOENDATE) + 1
    If Me.DA <> DD And Me.KINDM <> 2 Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "تعداد روزهاي مرخصي با محدوده تاريخي داده شده تطبيق ندارد آن را اصلاح كردم!"
       Me.DA = DD
       Me.MORAKHDAY = Me.DA * 440 + Me.HOU * 60 + Me.minu
       Call MORAKHDAY_AfterUpdate
    End If
  End If
End Sub


Private Sub Form_AfterUpdate()
   If Forms![BASEKNOW]![LOCKFAP] Then
        If Not Me.OKF Then
          Me.OKF = True
        End If
    End If
End Sub

Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(37, Me.IDNUM, Me.PERSONEL, "'برگه مرخصي شماره: " & Me.IDNUM & " مورخ " & Format(Me.MODATE, "####/##/##") & "  به نام: " & GETTAFNAME(PSK1 & "-" & PSM1 & "-" & Me.CODE) & "','" & PSK1 & "-" & PSM1 & "-" & Me.CODE & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
           On Error Resume Next
      On Error Resume Next: MID = Gettaskid(Me.IDNUM, 37)
    SHARH = "'برگه مرخصي شماره: " & Me.IDNUM & " مورخ " & Format(Me.MODATE, "####/##/##") & "  به نام: " & GETTAFNAME(PSK1 & "-" & PSM1 & "-" & Me.CODE) & "','" & PSK1 & "-" & PSM1 & "-" & Me.CODE & "'"
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, ":امضا شد1 ", ":امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",37," & Me.IDNUM & ",37 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
        Me.PERSONEL = GETUSERTASK(MID)
    Else
        td = Now()
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",37," & Me.IDNUM & ",37," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.IDNUM, 37)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, ":امضا شد1 ", ":امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",37," & Me.IDNUM & ",37 )")
    End If
    Call PERSONELUpdate(37, Me.IDNUM, Me.PERSONEL, SHARH)
    Me.PERSONEL.Visible = True
    If Not Me.OKF Then Me.OKF = True
    Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command54.Enabled = True
    Else
        Me.Command54.Enabled = False
    End If
End Sub

Private Sub SGN2_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
           On Error Resume Next
     On Error Resume Next: MID = Gettaskid(Me.IDNUM, 37)
    SHARH = "'برگه مرخصي شماره: " & Me.IDNUM & " مورخ " & Format(Me.MODATE, "####/##/##") & "  به نام: " & GETTAFNAME(PSK1 & "-" & PSM1 & "-" & Me.CODE) & "','" & PSK1 & "-" & PSM1 & "-" & Me.CODE & "'"
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",37," & Me.IDNUM & ",37 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
        Me.PERSONEL = GETUSERTASK(MID)
    Else
        td = Now()
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",37," & Me.IDNUM & ",37," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.IDNUM, 37)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",37," & Me.IDNUM & ",37 )")
    End If
    Call PERSONELUpdate(37, Me.IDNUM, Me.PERSONEL, SHARH)
    Me.PERSONEL.Visible = True
    If Not Me.OKF Then Me.OKF = True
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command54.Enabled = True
    Else
        Me.Command54.Enabled = False
    End If
End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
           On Error Resume Next
     On Error Resume Next: MID = Gettaskid(Me.IDNUM, 37)
    SHARH = "'برگه مرخصي شماره: " & Me.IDNUM & " مورخ " & Format(Me.MODATE, "####/##/##") & "  به نام: " & GETTAFNAME(PSK1 & "-" & PSM1 & "-" & Me.CODE) & "','" & PSK1 & "-" & PSM1 & "-" & Me.CODE & "'"
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",37," & Me.IDNUM & ",37 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
        Me.PERSONEL = GETUSERTASK(MID)
    Else
        td = Now()
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",37," & Me.IDNUM & ",37," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.IDNUM, 37)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",37," & Me.IDNUM & ",37 )")
    End If
    Call PERSONELUpdate(37, Me.IDNUM, Me.PERSONEL, SHARH)
    Me.PERSONEL.Visible = True
    If Not Me.OKF Then Me.OKF = True
    Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command54.Enabled = True
    Else
        Me.Command54.Enabled = False
    End If
End Sub
