Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim Meidnum As Double

Private Sub Command119_Click()
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    If Sendbefor(Me.PIDT, 90) Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين پذيرش قبلا ارسال شده آيا مايليد مجددا ارسال شود ؟ "
        If Not Forms![BASEKNOW]![Text44] Then
            Exit Sub
        End If
    End If
    td = Now()
    SHARH = "'پذيرش شماره: " & Me.PIDT & " مورخ " & Format(Me.PINDATE, "####/##/##") & "  به نام: " & GETTAFNAME(Me.POWNER) & "  دستگاه: " & Me.MODEL & "','" & Me.POWNER & "'"
    CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",90," & Me.PIDT & ",90," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
    If Not IsLoaded("main") Then
        DoCmd.OpenForm "main"
    End If
    Forms![main].Form.Requery
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command120_Click()
    DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "90-" & Me.PIDT & "+" & Meidnum
    DoCmd.Close acForm, Me.NAME, acSaveYes

End Sub

Private Sub Command54_Click()
                 On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
  DoCmd.OpenReport "PPAZIRESHTAM", acViewPreview, , "PIDT=" & Me.PIDT
End Sub

Private Sub Command55_Click()
                 On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
  DoCmd.OpenReport "PTAMIR_PAGE", acViewPreview, , "PIDT=" & Me.PIDT
End Sub

Private Sub Command56_Click()
                 On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "select * from head_lst where tag = 2 and  MAS = 1010 AND fnumco = " & Me.PIDT, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "با اين شماره پذيرش قبلا تعداد" & rst.RecordCount & "  حواله ثبت گرديده است آيا مايليد آن را مشاهده كنيد ؟"
        If Forms![BASEKNOW]![Text44] Then
           If IsLoaded("HEAD_LST_HAVL") Then
                DoCmd.Close acForm, "HEAD_LST_HAVL"
           End If
           If LETSGO("FRMOST") Or MID(Forms![BASEKNOW]![OPTIONSS], 18, 1) = "5" Then
                DoCmd.OpenForm "HEAD_LST_FROOSH22", acNormal, , , , , "TAG = 13 AND MAS = 1010 AND FNUMCO = " & Me.PIDT
                DoCmd.GoToRecord acDataForm, "HEAD_LST_FROOSH22", acPrevious
                Forms![HEAD_LST_FROOSH22]![CUST_NO] = Me.POWNER
                Forms![HEAD_LST_FROOSH22]![MAS] = 1010
                Forms![HEAD_LST_FROOSH22]![FNUMCO].Locked = True
                Forms![HEAD_LST_FROOSH22]![CUST_NO].Locked = True
                Forms![HEAD_LST_FROOSH22]![CUST_NO2].Locked = True
           Else
                DoCmd.OpenForm "HEAD_LST_HAVL", acNormal, , , , , "TAG = 2 AND MAS = 1010 AND FNUMCO = " & Me.PIDT
                DoCmd.GoToRecord acDataForm, "HEAD_LST_HAVL", acPrevious
                Forms![head_lst_havl]![CUST_NO] = Me.POWNER
                Forms![head_lst_havl]![MAS] = 1010
                Forms![head_lst_havl]![FNUMCO].Locked = True
                Forms![head_lst_havl]![CUST_NO].Locked = True
                Forms![head_lst_havl]![CUST_NO2].Locked = True
           End If
        End If
   Else
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "براي اين پذيرش حواله صادر نشده حواله جديد ايجاد شود؟ "
        If Forms![BASEKNOW]![Text44] Then
           If IsLoaded("HEAD_LST_HAVL") Then
                DoCmd.Close acForm, "HEAD_LST_HAVL"
           End If
           If LETSGO("FRMOST") Or MID(Forms![BASEKNOW]![OPTIONSS], 18, 1) = "5" Then
                DoCmd.OpenForm "HEAD_LST_FROOSH22", acNormal
                DoCmd.GoToRecord acDataForm, "HEAD_LST_FROOSH22", acNewRec
                Forms![HEAD_LST_FROOSH22]![FNUMCO] = Me.PIDT
                Forms![HEAD_LST_FROOSH22]![FNUMCO].Locked = True
                Forms![HEAD_LST_FROOSH22]![MAS] = 1010
                Forms![HEAD_LST_FROOSH22]![CUST_NO] = Me.POWNER
                Forms![HEAD_LST_FROOSH22]![TAH].SetFocus
                Forms![HEAD_LST_FROOSH22]![CUST_NO].Locked = True
                Forms![HEAD_LST_FROOSH22]![CUST_NO2].Locked = True
                DoCmd.RunCommand acCmdSaveRecord
            Else
                DoCmd.OpenForm "HEAD_LST_HAVL", acNormal
                DoCmd.GoToRecord acDataForm, "HEAD_LST_HAVL", acNewRec
                Forms![head_lst_havl]![FNUMCO] = Me.PIDT
                Forms![head_lst_havl]![FNUMCO].Locked = True
                Forms![head_lst_havl]![MAS] = 1010
                Forms![head_lst_havl]![CUST_NO] = Me.POWNER
                Forms![head_lst_havl]![TAH].SetFocus
                Forms![head_lst_havl]![CUST_NO].Locked = True
                Forms![head_lst_havl]![CUST_NO2].Locked = True
                DoCmd.RunCommand acCmdSaveRecord
            End If
        End If
   End If

End Sub

Private Sub Command57_Click()
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "select * from head_lst where tag = 14 AND MAS = 1010 and fnumco = " & Me.PIDT, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "با اين شماره پذيرش قبلا تعداد" & rst.RecordCount & "  فاكتور خدمات ثبت گرديده است آيا مايليد آن را مشاهده كنيد ؟"
        If Forms![BASEKNOW]![Text44] Then
           If IsLoaded("HEAD_LST_KHADAMAT") Then
                DoCmd.Close acForm, "HEAD_LST_KHADAMAT"
           End If
           DoCmd.OpenForm "HEAD_LST_KHADAMAT", acNormal, , , , , "FNUMCO"
           'DoCmd.GoToRecord acDataForm, "HEAD_LST_KHADAMAT", acPrevious
            Forms![HEAD_LST_KHADAMAT]![FNUMCO].Locked = True
            Forms![HEAD_LST_KHADAMAT]![CUST_NO].Locked = True
            Forms![HEAD_LST_KHADAMAT]![CUST_NO2].Locked = True
        End If
   Else
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "براي اين حواله فاكتور خدمات صادر نشده است فاكتور جديد ايجاد شود؟"
        If Forms![BASEKNOW]![Text44] Then
           If IsLoaded("HEAD_LST_KHADAMAT") Then
                DoCmd.Close acForm, "HEAD_LST_KHADAMAT"
           End If
            DoCmd.OpenForm "HEAD_LST_KHADAMAT", acNormal
            DoCmd.GoToRecord acDataForm, "HEAD_LST_KHADAMAT", acNewRec
            Forms![HEAD_LST_KHADAMAT]![FNUMCO] = Me.PIDT
            Forms![HEAD_LST_KHADAMAT]![MAS] = 1010
            Forms![HEAD_LST_KHADAMAT]![CUST_NO] = Me.POWNER
            Forms![HEAD_LST_KHADAMAT]![FNUMCO].Locked = True
            Forms![HEAD_LST_KHADAMAT]![CUST_NO].Locked = True
            Forms![HEAD_LST_KHADAMAT]![CUST_NO2].Locked = True
            Forms![HEAD_LST_KHADAMAT]![MOLAH].SetFocus
            DoCmd.RunCommand acCmdSaveRecord
         End If
   End If
End Sub

Private Sub Command58_Click()
  Dim MAN As Double, rst As New ADODB.Recordset
  If LETSGO("HTAF") Then
    If BLOCKED(GETKOL(Me.POWNER), GETMOIN(Me.POWNER), GETTAF(Me.POWNER)) Then
         DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
         Exit Sub
    Else
      If IsLoaded("R_DAFTAR_MOIN_LIST") Then
            DoCmd.Close acForm, "R_DAFTAR_MOIN_LIST"
      End If
        DoCmd.OpenForm "R_DAFTAR_MOIN_LIST", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
        With Forms!R_DAFTAR_MOIN_LIST
            .RecordSource = "EXEC usp_TafzilLedger  1 , 99999999 , '" & Me.POWNER & "'"
            .Requery                  ' بارگذاري مجدد داده
            .Visible = True           ' نمايش فرم
        End With
     End If
   End If

End Sub

Private Sub Command62_Click()
                 On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
  DoCmd.OpenReport "PPAZIRESHTAMM", acViewPreview, , "PIDT=" & Me.PIDT

End Sub

Private Sub Command63_Click()
                 On Error Resume Next
 DoCmd.RunCommand acCmdSaveRecord
  DoCmd.OpenReport "PPAZIRESHTAMM", , , "PIDT=" & Me.PIDT
End Sub

Private Sub ESLAH_Click()
Dim dt As Date
                 On Error Resume Next
  If Not IsNull(Me.PIDT) Then
    dt = Now()
    Call TR("PTAMIRAT", "(PIDT = " & Me.PIDT & " )", dt, 1)
    Call TR("PTVAZ", "(PIDT = " & Me.PIDT & " )", dt, 2)
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.PTVAZ_SUB.Enabled = True
        Me.PTVAZ_SUB.Form.AllowAdditions = True
   End If
End Sub

Private Sub Form_BeforeInsert(CANCEL As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, i As Integer
  If PIDT = 0 Or IsNull(PIDT) Then
    rst.Open "SELECT  MAX(PIDT) AS MaxOfID FROM  dbo.PTAMIRAT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.PIDT = 1
    Else
        Me.PIDT = rst.Fields(0) + 1
    End If
  End If
  rst.Close
  Set rst = New ADODB.Recordset
  rst.Open "SELECT   TOP 1 PERCENT PIDT, PKIND, IYALAT, CITY, PINPKIND, PTAHVKIND, PAYAM,GRCOMPA FROM dbo.PTAMIRAT ORDER BY PIDT DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
    Me.PKIND = rst.Fields("PKIND")
    Me.IYALAT = rst.Fields("IYALAT")
    Me.CITY = rst.Fields("CITY")
    Me.PINPKIND = rst.Fields("PINPKIND")
    Me.PTAHVKIND = rst.Fields("PTAHVKIND")
    Me.PAYAM = rst.Fields("PAYAM")
    Me.GRCOMPA = rst.Fields("GRCOMPA")
  End If
End Sub

Private Sub Form_Current()
                    On Error Resume Next
   If Not (IsNull(Me.PIDT) Or IsNull(Me.MODEL) Or IsNull(Me.POWNER) Or Me.POWNER = "" Or Me.MODEL = "") Then
    Me.Command54.Enabled = True
    Me.Command55.Enabled = True
    Me.Command56.Enabled = True
    Me.Command57.Enabled = True
    Me.Command58.Enabled = True
    Me.Command62.Enabled = True
    Me.Command63.Enabled = True
 Else
    Me.Command54.Enabled = False
    Me.Command55.Enabled = False
    Me.Command56.Enabled = False
    Me.Command57.Enabled = False
    Me.Command58.Enabled = False
    Me.Command62.Enabled = False
    Me.Command63.Enabled = False
End If
  If Forms![BASEKNOW]![SIGN] Then
        If Me.SGN1 Or Me.SGN3 Then
            Me.Command55.Enabled = True
            Me.Command62.Enabled = True
            Me.Command54.Enabled = True
            Me.Command63.Enabled = True
        Else
            Me.Command55.Enabled = False
            Me.Command62.Enabled = False
            Me.Command54.Enabled = False
            Me.Command63.Enabled = False
       End If
 End If
' If (Not ISNULL(Me.TAHV_DATE) Or Not ISNULL(Me.TAHV_TIME)) And Left(UCurrentUser(), 10) <> ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
'    Me.AllowEdits = False
'    Me.AllowDeletions = False
' Else
'    Me.AllowDeletions = True
'    Me.AllowEdits = True
'   Call SETSECURITY(Me.NAME, "PTM", 3)
' End If
   If Me.PIDT > 0 Then
        Call LetSigneTick(Me, 90, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
  If Not IsNull(Me.PIDT) Then
     Me.AllowDeletions = False
     Me.AllowEdits = False
     Me.PTVAZ_SUB.Enabled = False
     Me.ESLAH.Enabled = True
 End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  rst.Open "select fnumco from head_lst where fnumco = " & Me.PIDT & " and (tag = 13 OR TAG = 14 ) and TAMIR = 1010 ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
      DoCmd.OpenForm "mesag", , , , , acDialog, "براي اين پذيرش فاكتور فروش يا فاكتور خدمات  صادر گرديده است بنابر اين قابل حذف نيست مگر اينكه ابتدا فاكتور فروش صادر شده قبلي به نام مشتري قبلي حذف گردد ."
      CANCEL = True
  End If
  rst.Close
  Set rst = New ADODB.Recordset
  rst.Open "select fnumco from head_lst where fnumco = " & Me.PIDT & " and  TAG = 2 and TAMIR = 1010", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
      DoCmd.OpenForm "mesag", , , , , acDialog, "براي اين پذيرش  حواله صادر گرديده است بنابر اين قابل حذف نيست مگر اينكه ابتدا  حواله صادر شده  حذف گردد ."
      CANCEL = True
  End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
   Call SETSECURITY(Me.NAME, "PTM", 3)
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
    If Not IsNull(Me.OpenArgs) Then
        Me.RecordSource = "SELECT * FROM PTAMIRAT WHERE PIDT = " & Me.OpenArgs
    Else
        DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
    End If
   Me.PSERIAL.SetFocus
    If Forms![BASEKNOW]![SIGN] Then
       Me.SGN1.Visible = True
       Me.SGN2.Visible = True
       Me.SGN3.Visible = True
       Me.LSGN1.Visible = True
       Me.LSGN2.Visible = True
       Me.LSGN3.Visible = True
        If IsLoaded("main") Then
            If Not IsNull(Forms![main]![TASKS]![IDNUM]) And (Forms![main]![TASKS]![tg] = 90) Then
                Meidnum = Forms![main]![TASKS]![IDNUM]
                If GETSIGNER(90, Forms![BASEKNOW]![USERCOD]) > 0 Then
                    Me.Command120.Visible = True
                End If
          End If
        End If
       Me.Command119.Visible = True
    End If
    Set rst = New ADODB.Recordset
    Dim VS As String
    VS = ""
    rst.Open "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL=0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        VS = VS & rst.Fields("IDD") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & Chr(34)
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    Me.PERSONEL.RowSource = VS
    Me.SGN1usid.RowSource = VS
    Me.sgn2usid.RowSource = VS
    Me.sgn3usid.RowSource = VS
End Sub

Private Sub GRENDDATE_Exit(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(GRENDDATE, False)

End Sub

Private Sub MODEL_AfterUpdate()
                 On Error Resume Next
  If Me.MODEL = "-" Or Me.MODEL = "+" Or Me.MODEL = "--" Or Me.MODEL = "++" Then
     Me.MODEL = ""
     DoCmd.OpenForm "SERCHMODEL1"
  End If


End Sub

Private Sub MODEL_Exit(CANCEL As Integer)
                    On Error Resume Next
   If Not (IsNull(Me.PIDT) Or IsNull(Me.MODEL) Or IsNull(Me.POWNER) Or Me.POWNER = "" Or Me.MODEL = "") Then
    Me.Command54.Enabled = True
    Me.Command55.Enabled = True
    Me.Command56.Enabled = True
    Me.Command57.Enabled = True
    Me.Command58.Enabled = True
    Me.Command62.Enabled = True
    Me.Command63.Enabled = True
 Else
    Me.Command54.Enabled = False
    Me.Command55.Enabled = False
    Me.Command56.Enabled = False
    Me.Command57.Enabled = False
    Me.Command58.Enabled = False
    Me.Command62.Enabled = False
    Me.Command63.Enabled = False
 End If
  If Forms![BASEKNOW]![SIGN] Then
        If Me.SGN1 Or Me.SGN3 Then
            Me.Command55.Enabled = True
            Me.Command62.Enabled = True
            Me.Command54.Enabled = True
            Me.Command63.Enabled = True
        Else
            Me.Command55.Enabled = False
            Me.Command62.Enabled = False
            Me.Command54.Enabled = False
            Me.Command63.Enabled = False
       End If
 End If
End Sub

Private Sub PINDATE_Exit(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(PINDATE, True)

End Sub




Private Sub POWNER_AfterUpdate()
                  On Error Resume Next
  Dim rst As New ADODB.Recordset
  rst.Open "SELECT     ADDRESS, TEL FROM dbo.TDETA_HES WHERE (N_KOL = " & GETKOL(Me.POWNER) & ")  And (NUMBER = " & GETMOIN(Me.POWNER) & ") And (TNUMBER = " & GETTAF(Me.POWNER) & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
      If Me.ADRESS = "" Or IsNull(Me.ADRESS) Then
            If rst.Fields(0) & "  تلفن : " & rst.Fields(1) <> "  تلفن : " Then
                Me.ADRESS = rst.Fields(0) & "  تلفن : " & rst.Fields(1)
            End If
      End If
  End If
End Sub

Private Sub POWNER_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  rst.Open "select fnumco from head_lst where fnumco = " & Me.PIDT & " and (tag = 13 or tag = 14) and TAMIR = 1010", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
      DoCmd.OpenForm "mesag", , , , , acDialog, "براي اين پذيرش فاكتور فروش يا خدمات صادر گرديده است بنابر اين مالك دستگاه قابل تغيير نيست مگر اينكه ابتدا فاكتور فروش صادر شده قبلي به نام مشتري قبلي حذف گردد ."
      CANCEL = True
  Else
     rst.Close
     Set rst = New ADODB.Recordset
     rst.Open "select fnumco from head_lst where fnumco = " & Me.PIDT & " and (tag = 2 ) and TAMIR = 1010", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "براي اين پذيرش حواله صادر شده حتما بايد مجددا حواله باز شود تا نام مشتري در حواله نيز اصلاح گردد"
     End If
  End If
End Sub

Private Sub POWNER_Exit(CANCEL As Integer)
                    On Error Resume Next
   If Not (IsNull(Me.PIDT) Or IsNull(Me.MODEL) Or IsNull(Me.POWNER) Or Me.POWNER = "" Or Me.MODEL = "") Then
    Me.Command54.Enabled = True
    Me.Command55.Enabled = True
    Me.Command56.Enabled = True
    Me.Command57.Enabled = True
    Me.Command58.Enabled = True
    Me.Command62.Enabled = True
    Me.Command63.Enabled = True
 Else
    Me.Command54.Enabled = False
    Me.Command55.Enabled = False
    Me.Command56.Enabled = False
    Me.Command57.Enabled = False
    Me.Command58.Enabled = False
    Me.Command62.Enabled = False
    Me.Command63.Enabled = False
 End If
                   On Error Resume Next
  If Me.ADRESS = "" Or IsNull(Me.ADRESS) Then
        Dim rst As New ADODB.Recordset
        rst.Open "SELECT     ADDRESS, TEL FROM dbo.TDETA_HES WHERE (N_KOL = " & GETKOL(Me.POWNER) & ")  And (NUMBER = " & GETMOIN(Me.POWNER) & ") And (TNUMBER = " & GETTAF(Me.POWNER) & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If rst.Fields(0) & "  تلفن : " & rst.Fields(1) <> "  تلفن : " Then
                Me.ADRESS = rst.Fields(0) & "  تلفن : " & rst.Fields(1)
            End If
        End If
  End If
  If Forms![BASEKNOW]![SIGN] Then
        If Me.SGN1 Or Me.SGN3 Then
            Me.Command55.Enabled = True
            Me.Command62.Enabled = True
            Me.Command54.Enabled = True
            Me.Command63.Enabled = True
        Else
            Me.Command55.Enabled = False
            Me.Command62.Enabled = False
            Me.Command54.Enabled = False
            Me.Command63.Enabled = False
       End If
 End If

End Sub

Private Sub POWNER_NotInList(NewData As String, Response As Integer)
                 On Error Resume Next
  If Me.POWNER.Text = "-" Or Me.POWNER.Text = "+" Or Me.POWNER.Text = "--" Or Me.POWNER.Text = "++" Then
     Me.POWNER.Text = ""
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "18"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub

Private Sub PSERIAL_AfterUpdate()
                 On Error Resume Next
 Dim rst As New ADODB.Recordset
 rst.Open "SELECT     * FROM dbo.PTAMIRAT WHERE     (PSERIAL = '" & Me.PSERIAL & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

 If rst.RecordCount > 0 And Trim(Me.PSERIAL) <> "0" Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "با اين شماره سريال قبلاً پذيرش انجام شده است آيا اطلاعات پذيرش قبلي بازيابي شود؟"
    If Forms![BASEKNOW]![Text44] Then
       Me.PKIND = rst.Fields("PKIND")
       Me.MODEL = rst.Fields("MODEL")
       Me.GRENDDATE = rst.Fields("GRENDDATE")
       Me.POWNER = rst.Fields("POWNER")
       Me.POWNERNAM = rst.Fields("POWNERNAM")
       Me.GRCOMPA = rst.Fields("GRCOMPA")
       Me.IYALAT = rst.Fields("IYALAT")
       Me.CITY = rst.Fields("CITY")
       Me.PINPKIND = rst.Fields("PINPKIND")
       Me.PTAHVKIND = rst.Fields("PTAHVKIND")
       Me.EYBEZHAR = rst.Fields("EYBEZHAR")
       Me.IDIATEK = rst.Fields("IDIATEK")
       Me.ADRESS = rst.Fields("ADRESS")
       Me.TAMIRKAR1 = rst.Fields("TAMIRKAR1")
       Me.TAMIRKAR2 = rst.Fields("TAMIRKAR2")
    End If
   If Not (IsNull(Me.PIDT) Or IsNull(Me.MODEL) Or IsNull(Me.POWNER) Or Me.POWNER = "" Or Me.MODEL = "") Then
    Me.Command54.Enabled = True
    Me.Command55.Enabled = True
    Me.Command56.Enabled = True
    Me.Command57.Enabled = True
    Me.Command58.Enabled = True
    Me.Command62.Enabled = True
    Me.Command63.Enabled = True
 Else
    Me.Command54.Enabled = False
    Me.Command55.Enabled = False
    Me.Command56.Enabled = False
    Me.Command57.Enabled = False
    Me.Command58.Enabled = False
    Me.Command62.Enabled = False
    Me.Command63.Enabled = False
    End If
 End If
  If Forms![BASEKNOW]![SIGN] Then
         If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
            Me.Command54.Enabled = True
            Me.Command55.Enabled = True
            Me.Command56.Enabled = True
            Me.Command57.Enabled = True
            Me.Command58.Enabled = True
            Me.Command62.Enabled = True
            Me.Command63.Enabled = True
         Else
            Me.Command54.Enabled = False
            Me.Command55.Enabled = False
            Me.Command56.Enabled = False
            Me.Command57.Enabled = False
            Me.Command58.Enabled = False
            Me.Command62.Enabled = False
            Me.Command63.Enabled = False
        End If
 End If
End Sub

Private Sub TAHV_DATE_DblClick(CANCEL As Integer)
                 On Error Resume Next
If Me.AllowDeletions = True Then
    Me.TAHV_DATE = FARSIDATE(Now())
    Dim rst As New ADODB.Recordset
    rst.Open "PTVAZ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("PIDT") = Me.PIDT
    rst.Fields("PDATE") = FARSIDATE(Now())
    rst.Fields("PTIME") = Hour(Now()) * 100 + Minute(Now())
    rst.Fields("PVAZ") = 7
    rst.update
    rst.Close
End If

End Sub

Private Sub TAHV_DATE_Exit(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(TAHV_DATE, False)
End Sub

Private Sub TAHV_TIME_DblClick(CANCEL As Integer)
                 On Error Resume Next
If Me.AllowDeletions = True Then
  Me.TAHV_TIME = Hour(Now()) * 100 + Minute(Now())
End If
End Sub

Private Sub TAMIRKAR1_NotInList(NewData As String, Response As Integer)
                 On Error Resume Next
  If Me.TAMIRKAR1.Text = "-" Or Me.TAMIRKAR1.Text = "+" Or Me.TAMIRKAR1.Text = "--" Or Me.TAMIRKAR1.Text = "++" Then
     Me.TAMIRKAR1.Text = ""
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "19"
  Else
    Me.TAMIRKAR1 = NewData
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub

Private Sub TAMIRKAR2_NotInList(NewData As String, Response As Integer)
                 On Error Resume Next
  If Me.TAMIRKAR2.Text = "-" Or Me.TAMIRKAR2.Text = "+" Or Me.TAMIRKAR2.Text = "--" Or Me.TAMIRKAR2.Text = "++" Then
     Me.TAMIRKAR2.Text = ""
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "20"
  Else
    Me.TAMIRKAR2 = NewData
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub
Private Sub PIDT_DblClick(CANCEL As Integer)
                  On Error Resume Next
 DoCmd.OpenForm "PTAMIRAT_REPORT", acFormDS
End Sub

Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(90, Me.PIDT, Me.PERSONEL, "'پذيرش شماره: " & Me.PIDT & " مورخ " & Format(Me.PINDATE, "####/##/##") & "  به نام: " & GETTAFNAME(Me.POWNER) & "  دستگاه: " & Me.MODEL & "','" & Me.POWNER & "'")
        DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.PIDT, 90)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",90," & Me.PIDT & ",90 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
    SHARH = "'پذيرش شماره: " & Me.PIDT & " مورخ " & Format(Me.PINDATE, "####/##/##") & "  به نام: " & GETTAFNAME(Me.POWNER) & "  دستگاه: " & Me.MODEL & "','" & Me.POWNER & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",90," & Me.PIDT & ",90," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.PIDT, 90)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",90," & Me.PIDT & ",90 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID
    Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
 If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
    Me.Command54.Enabled = True
    Me.Command55.Enabled = True
    Me.Command56.Enabled = True
    Me.Command57.Enabled = True
    Me.Command58.Enabled = True
    Me.Command62.Enabled = True
    Me.Command63.Enabled = True
 Else
    Me.Command54.Enabled = False
    Me.Command55.Enabled = False
    Me.Command56.Enabled = False
    Me.Command57.Enabled = False
    Me.Command58.Enabled = False
    Me.Command62.Enabled = False
    Me.Command63.Enabled = False
End If
End Sub

Private Sub SGN2_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.PIDT, 90)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",90," & Me.PIDT & ",90 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
    SHARH = "'پذيرش شماره: " & Me.PIDT & " مورخ " & Format(Me.PINDATE, "####/##/##") & "  به نام: " & GETTAFNAME(Me.POWNER) & "  دستگاه: " & Me.MODEL & "','" & Me.POWNER & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",90," & Me.PIDT & ",90," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.PIDT, 90)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",90," & Me.PIDT & ",90 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
 If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
    Me.Command54.Enabled = True
    Me.Command55.Enabled = True
    Me.Command56.Enabled = True
    Me.Command57.Enabled = True
    Me.Command58.Enabled = True
    Me.Command62.Enabled = True
    Me.Command63.Enabled = True
 Else
    Me.Command54.Enabled = False
    Me.Command55.Enabled = False
    Me.Command56.Enabled = False
    Me.Command57.Enabled = False
    Me.Command58.Enabled = False
    Me.Command62.Enabled = False
    Me.Command63.Enabled = False
End If

End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.PIDT, 90)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",90," & Me.PIDT & ",90 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'پذيرش شماره: " & Me.PIDT & " مورخ " & Format(Me.PINDATE, "####/##/##") & "  به نام: " & GETTAFNAME(Me.POWNER) & "  دستگاه: " & Me.MODEL & "','" & Me.POWNER & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",90," & Me.PIDT & ",90," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.PIDT, 90)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",90," & Me.PIDT & ",90 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID
    Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
 If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
    Me.Command54.Enabled = True
    Me.Command55.Enabled = True
    Me.Command56.Enabled = True
    Me.Command57.Enabled = True
    Me.Command58.Enabled = True
    Me.Command62.Enabled = True
    Me.Command63.Enabled = True
 Else
    Me.Command54.Enabled = False
    Me.Command55.Enabled = False
    Me.Command56.Enabled = False
    Me.Command57.Enabled = False
    Me.Command58.Enabled = False
    Me.Command62.Enabled = False
    Me.Command63.Enabled = False
End If
End Sub
